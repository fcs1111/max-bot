from fastapi import FastAPI, Request, Header
from fastapi.responses import PlainTextResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

import json
import os
import shutil
import traceback
import zipfile
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from pptx import Presentation


app = FastAPI()

# Public URL of your Railway service. Set this in Railway variables after deploy.
BASE_URL = os.getenv("BASE_URL", "https://web-production-a9964.up.railway.app")

# Token from MAX Master Bot / business.max.ru. Do not paste it into code.
BOT_TOKEN = os.getenv("BOT_TOKEN", "")
MAX_API_URL = os.getenv("MAX_API_URL", "https://platform-api.max.ru")

# Optional secret for MAX webhook. Set the same value in Railway and webhook subscription.
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "")

TEMPLATES_DIR = Path("templates")
EXCEL_DIR = Path("excel")
OUTPUT_DIR = Path("output")
STATE_DIR = Path("state")

for directory in (TEMPLATES_DIR, EXCEL_DIR, OUTPUT_DIR, STATE_DIR):
    directory.mkdir(exist_ok=True)

app.mount("/files", StaticFiles(directory=str(OUTPUT_DIR)), name="files")


# ------------------ BASIC FILE HELPERS ------------------

def sanitize_filename(value: Any, fallback: str = "file") -> str:
    name = str(value or fallback).strip()
    bad_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    for char in bad_chars:
        name = name.replace(char, "")
    return name[:120] or fallback


def state_path(user_id: str) -> Path:
    return STATE_DIR / f"{user_id}.json"


def load_state(user_id: str) -> dict:
    path = state_path(user_id)
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_state(user_id: str, data: dict) -> None:
    state_path(user_id).write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def download_file(file_url: str, save_dir: Path, filename: str | None = None, force_ext: str | None = None):
    response = requests.get(file_url, timeout=60)
    response.raise_for_status()

    if not filename:
        filename = file_url.split("/")[-1].split("?")[0] or "file"

    filename = sanitize_filename(filename)

    if force_ext and not filename.lower().endswith(force_ext):
        filename = f"{filename}{force_ext}"

    path = save_dir / filename
    path.write_bytes(response.content)

    return filename, path


# ------------------ POWERPOINT GENERATION ------------------

def replace_text_in_shape(shape, row: pd.Series, columns) -> None:
    if not shape.has_text_frame:
        return

    for paragraph in shape.text_frame.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        if not full_text:
            continue

        replaced = False
        for col in columns:
            placeholder = str(col).strip()
            if placeholder in full_text:
                full_text = full_text.replace(placeholder, str(row[col]))
                replaced = True

        if replaced and paragraph.runs:
            for run in paragraph.runs:
                run.text = ""
            paragraph.runs[0].text = full_text


def generate_pptx(template_path: Path, excel_path: Path, user_id: str) -> str:
    user_output_dir = OUTPUT_DIR / user_id

    if user_output_dir.exists():
        shutil.rmtree(user_output_dir)
    user_output_dir.mkdir(exist_ok=True)

    df = pd.read_excel(excel_path)
    if df.empty:
        raise ValueError("Excel пустой. Добавь хотя бы одну строку с данными.")

    generated_files = []

    for _, row in df.iterrows():
        prs = Presentation(str(template_path))

        for slide in prs.slides:
            for shape in slide.shapes:
                replace_text_in_shape(shape, row, df.columns)

        safe_name = sanitize_filename(row[df.columns[0]], fallback="presentation")
        pptx_path = user_output_dir / f"{safe_name}.pptx"
        prs.save(str(pptx_path))
        generated_files.append(pptx_path)

    zip_name = f"{user_id}_result.zip"
    zip_path = OUTPUT_DIR / zip_name

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in generated_files:
            zipf.write(file, arcname=file.name)

    return zip_name


# ------------------ MAX API HELPERS ------------------

def max_headers() -> dict:
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN не задан в Railway Variables")
    return {"Authorization": BOT_TOKEN, "Content-Type": "application/json"}


def inline_keyboard(buttons: list[list[dict]]) -> dict:
    return {
        "type": "inline_keyboard",
        "payload": {"buttons": buttons},
    }


def main_menu_keyboard() -> dict:
    return inline_keyboard([
        [
            {"type": "callback", "text": "Загрузить шаблон", "payload": "upload_template"},
        ],
        [
            {"type": "callback", "text": "Создать грамоты", "payload": "upload_excel"},
        ],
        [
            {"type": "callback", "text": "Помощь", "payload": "help"},
        ],
    ])


def send_max_message(user_id: str, text: str, attachments: list[dict] | None = None) -> None:
    response = requests.post(
        f"{MAX_API_URL}/messages",
        params={"user_id": user_id},
        headers=max_headers(),
        json={
            "text": text,
            "attachments": attachments or [],
            "link": None,
        },
        timeout=30,
    )
    response.raise_for_status()


def answer_callback(callback_id: str, notification: str = "Готово") -> None:
    if not callback_id:
        return

    response = requests.post(
        f"{MAX_API_URL}/answers",
        params={"callback_id": callback_id},
        headers=max_headers(),
        json={"notification": notification, "message": None},
        timeout=30,
    )
    response.raise_for_status()


def extract_user_id(update: dict) -> str | None:
    if update.get("update_type") == "bot_started":
        return str((update.get("user") or {}).get("user_id") or "")

    if update.get("update_type") == "message_callback":
        return str(((update.get("callback") or {}).get("user") or {}).get("user_id") or "")

    message = update.get("message") or {}
    sender = message.get("sender") or {}
    return str(sender.get("user_id") or "")


def extract_text(update: dict) -> str:
    message = update.get("message") or {}
    body = message.get("body") or {}
    return (body.get("text") or "").strip()


def extract_file_attachment(update: dict, allowed_extensions: list[str]) -> tuple[str | None, str | None, str | None]:
    message = update.get("message") or {}
    body = message.get("body") or {}
    attachments = body.get("attachments") or []

    found = []
    for attachment in attachments:
        if not isinstance(attachment, dict):
            continue
        if attachment.get("type") != "file":
            continue

        payload = attachment.get("payload") or {}
        url = payload.get("url") or attachment.get("url") or ""
        filename = attachment.get("filename") or url.split("/")[-1].split("?")[0]
        clean_name = filename.lower()
        clean_url = url.split("?")[0].lower()

        for ext in allowed_extensions:
            if clean_name.endswith(ext) or clean_url.endswith(ext):
                found.append((url, filename, ext))

    return found[-1] if found else (None, None, None)


# ------------------ MAX BOT FLOW ------------------

async def handle_max_update(update: dict) -> None:
    user_id = extract_user_id(update)
    if not user_id:
        return

    update_type = update.get("update_type")

    if update_type == "bot_started":
        send_max_message(
            user_id,
            "Привет! Я помогу сделать грамоты из PPTX-шаблона и Excel-файла.\n\n"
            "1. Нажми «Загрузить шаблон» и отправь PPTX.\n"
            "2. Потом нажми «Создать грамоты» и отправь Excel.",
            [main_menu_keyboard()],
        )
        return

    if update_type == "message_callback":
        callback = update.get("callback") or {}
        payload = callback.get("payload")
        callback_id = callback.get("callback_id")
        state = load_state(user_id)

        if payload == "upload_template":
            state["mode"] = "waiting_template"
            save_state(user_id, state)
            answer_callback(callback_id, "Жду PPTX")
            send_max_message(user_id, "Отправь PPTX-шаблон одним файлом.")
            return

        if payload == "upload_excel":
            template_path = state.get("template_path")
            if not template_path or not Path(template_path).exists():
                answer_callback(callback_id, "Сначала нужен шаблон")
                send_max_message(user_id, "Сначала загрузи PPTX-шаблон.", [main_menu_keyboard()])
                return

            state["mode"] = "waiting_excel"
            save_state(user_id, state)
            answer_callback(callback_id, "Жду Excel")
            send_max_message(user_id, "Отправь Excel-файл .xlsx или .xls.")
            return

        if payload == "help":
            answer_callback(callback_id, "Помощь")
            send_max_message(
                user_id,
                "Как подготовить файлы:\n\n"
                "В PPTX напиши слова-плейсхолдеры, например: Имя, Класс, Дата.\n"
                "В Excel сделай такие же названия колонок: Имя, Класс, Дата.\n"
                "Каждая строка Excel станет отдельной грамотой.",
                [main_menu_keyboard()],
            )
            return

    if update_type != "message_created":
        return

    text = extract_text(update).lower()
    if text in {"/start", "старт", "меню"}:
        send_max_message(user_id, "Главное меню:", [main_menu_keyboard()])
        return

    state = load_state(user_id)
    mode = state.get("mode")

    if mode == "waiting_template":
        file_url, filename, _ = extract_file_attachment(update, [".pptx"])
        if not file_url:
            send_max_message(user_id, "Я не вижу PPTX-файл. Нажми скрепку и отправь именно .pptx.")
            return

        _, template_path = download_file(file_url, TEMPLATES_DIR, filename=filename, force_ext=".pptx")

        try:
            prs = Presentation(str(template_path))
            slide_count = len(prs.slides)
        except Exception as exc:
            send_max_message(user_id, f"Файл не похож на нормальный PPTX:\n{exc}")
            return

        state["template_path"] = str(template_path)
        state["template_name"] = template_path.name
        state["mode"] = None
        save_state(user_id, state)

        send_max_message(
            user_id,
            f"Шаблон загружен.\nФайл: {template_path.name}\nСлайдов: {slide_count}\n\n"
            "Теперь нажми «Создать грамоты» и отправь Excel.",
            [main_menu_keyboard()],
        )
        return

    if mode == "waiting_excel":
        template_path = state.get("template_path")
        if not template_path or not Path(template_path).exists():
            state["mode"] = None
            save_state(user_id, state)
            send_max_message(user_id, "Шаблон не найден. Загрузи PPTX заново.", [main_menu_keyboard()])
            return

        file_url, filename, ext = extract_file_attachment(update, [".xlsx", ".xls"])
        if not file_url:
            send_max_message(user_id, "Я не вижу Excel-файл. Отправь .xlsx или .xls.")
            return

        _, excel_path = download_file(file_url, EXCEL_DIR, filename=filename, force_ext=ext or ".xlsx")
        send_max_message(user_id, "Excel получил. Генерирую грамоты, это может занять немного времени.")

        zip_name = generate_pptx(
            template_path=Path(template_path),
            excel_path=excel_path,
            user_id=user_id,
        )

        state["mode"] = None
        save_state(user_id, state)

        full_url = f"{BASE_URL}/files/{zip_name}"
        send_max_message(user_id, f"Файлы готовы.\n\nСкачать ZIP:\n{full_url}", [main_menu_keyboard()])
        return

    send_max_message(user_id, "Выбери действие в меню:", [main_menu_keyboard()])


# ------------------ MAX WEBHOOK ENDPOINTS ------------------

@app.post("/max/webhook")
async def max_webhook(request: Request, x_max_bot_api_secret: str | None = Header(default=None)):
    try:
        if WEBHOOK_SECRET and x_max_bot_api_secret != WEBHOOK_SECRET:
            return JSONResponse({"ok": False, "error": "bad secret"}, status_code=403)

        update = await request.json()
        await handle_max_update(update)
        return {"ok": True}
    except Exception as exc:
        print("MAX webhook error:", exc)
        print(traceback.format_exc())
        return {"ok": True}


def register_max_webhook() -> str:
    webhook_url = f"{BASE_URL}/max/webhook"
    body = {
        "url": webhook_url,
        "update_types": ["bot_started", "message_created", "message_callback"],
    }
    if WEBHOOK_SECRET:
        body["secret"] = WEBHOOK_SECRET

    response = requests.post(
        f"{MAX_API_URL}/subscriptions",
        headers=max_headers(),
        json=body,
        timeout=30,
    )
    return f"MAX response {response.status_code}:\n{response.text}"


@app.get("/setup_max_webhook", response_class=PlainTextResponse)
def setup_max_webhook_from_browser():
    return register_max_webhook()


@app.post("/setup_max_webhook", response_class=PlainTextResponse)
def setup_max_webhook():
    return register_max_webhook()


# ------------------ OLD WATBOT COMPATIBILITY ENDPOINTS ------------------

def extract_watbot_file_url(variables: list, allowed_extensions: list[str]):
    found_urls = []

    for var in variables:
        if not var:
            continue

        if isinstance(var, dict):
            payload = var.get("payload") or {}
            url = payload.get("url") or var.get("url") or ""
        elif isinstance(var, str):
            url = var
        else:
            url = ""

        if url:
            clean_url = url.split("?")[0].lower()
            for ext in allowed_extensions:
                if clean_url.endswith(ext):
                    found_urls.append((url, ext))

    return found_urls[-1] if found_urls else (None, None)


@app.post("/upload_template")
async def upload_template(request: Request):
    try:
        data = await request.json()
        variables = data if isinstance(data, list) else data.get("variables") or []
        contact = {} if isinstance(data, list) else data.get("contact") or {}
        user_id = str(contact.get("id", "default"))

        file_url, _ = extract_watbot_file_url(variables, [".pptx"])
        if not file_url:
            return PlainTextResponse("PPTX не найден.")

        filename, template_path = download_file(file_url, TEMPLATES_DIR, force_ext=".pptx")

        try:
            prs = Presentation(str(template_path))
            slide_count = len(prs.slides)
        except Exception as exc:
            return PlainTextResponse(f"Файл не является валидным PPTX:\n{exc}")

        state = load_state(user_id)
        state["template_path"] = str(template_path)
        state["template_name"] = filename
        state["mode"] = None
        save_state(user_id, state)

        return PlainTextResponse(
            f"Шаблон загружен ✅\n"
            f"Файл: {filename} ({slide_count} слайдов)\n\n"
            f"Теперь отправь Excel файл (.xlsx)"
        )

    except Exception as exc:
        return PlainTextResponse(f"Ошибка upload_template:\n{exc}\n\n{traceback.format_exc()}")


@app.post("/upload_excel")
async def upload_excel(request: Request):
    try:
        data = await request.json()
        variables = data if isinstance(data, list) else data.get("variables") or []
        contact = {} if isinstance(data, list) else data.get("contact") or {}
        user_id = str(contact.get("id", "default"))

        state = load_state(user_id)
        template_path = state.get("template_path")

        if not template_path:
            return PlainTextResponse("Сначала загрузи шаблон PPTX")

        if not Path(template_path).exists():
            return PlainTextResponse("Шаблон не найден на диске. Загрузи шаблон заново.")

        file_url, ext = extract_watbot_file_url(variables, [".xlsx", ".xls"])
        if not file_url:
            return PlainTextResponse("Excel не найден.")

        _, excel_path = download_file(file_url, EXCEL_DIR, force_ext=ext or ".xlsx")
        zip_name = generate_pptx(Path(template_path), excel_path, user_id)

        full_url = f"{BASE_URL}/files/{zip_name}"
        return PlainTextResponse(f"Файлы готовы ✅\n\n{full_url}")

    except Exception as exc:
        return PlainTextResponse(f"Ошибка сервера:\n{exc}\n\n{traceback.format_exc()}")


# ------------------ STATUS ------------------

@app.get("/", response_class=PlainTextResponse)
def status():
    return "бот работает"
