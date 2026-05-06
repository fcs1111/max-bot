from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
from fastapi.staticfiles import StaticFiles

import requests
import os
import json
import pandas as pd
from pptx import Presentation
import zipfile
import shutil
import traceback

app = FastAPI()

BASE_URL = "https://web-production-a9964.up.railway.app"

TEMPLATES_DIR = "templates"
EXCEL_DIR = "excel"
OUTPUT_DIR = "output"
STATE_DIR = "state"

MAX_TEMPLATES = 5

os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(STATE_DIR, exist_ok=True)

app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")

# ==================== ШАБЛОНЫ ====================

def get_templates_path(user_id: str) -> str:
    return os.path.join(
        STATE_DIR,
        f"{user_id}_templates.json"
    )

def load_templates(user_id: str) -> dict:

    path = get_templates_path(user_id)

    if os.path.exists(path):

        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    return {}

def save_templates(user_id: str, templates: dict):

    path = get_templates_path(user_id)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(
            templates,
            f,
            ensure_ascii=False,
            indent=2
        )

def get_selected_template(user_id: str):

    templates = load_templates(user_id)

    selected = templates.get("selected")

    if not selected:
        return None

    slot = templates.get("slots", {}).get(str(selected))

    if not slot:
        return None

    return slot.get("path")

# ==================== СКАЧИВАНИЕ ====================

def download_file(file_url: str, save_dir: str, force_ext: str = None):

    response = requests.get(
        file_url,
        timeout=30
    )

    response.raise_for_status()

    filename = file_url.split("/")[-1].split("?")[0]

    if force_ext and not filename.lower().endswith(force_ext):
        filename += force_ext

    path = os.path.join(
        save_dir,
        filename
    )

    with open(path, "wb") as f:
        f.write(response.content)

    return filename, path

# ==================== ИЗВЛЕЧЕНИЕ URL ====================

def extract_file_url(
    variables: list,
    allowed_extensions: list
):

    found_urls = []

    for var in variables:

        if not var:
            continue

        if isinstance(var, dict):

            payload = var.get("payload") or {}

            url = (
                payload.get("url")
                or var.get("url")
                or ""
            )

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

# ==================== ГЕНЕРАЦИЯ PPTX ====================

def generate_pptx(
    template_path: str,
    excel_path: str,
    user_id: str
) -> str:

    user_output_dir = os.path.join(
        OUTPUT_DIR,
        user_id
    )

    if os.path.exists(user_output_dir):
        shutil.rmtree(user_output_dir)

    os.makedirs(user_output_dir, exist_ok=True)

    df = pd.read_excel(excel_path)

    generated_files = []

    for i, row in df.iterrows():

        prs = Presentation(template_path)

        for slide in prs.slides:

            for shape in slide.shapes:

                if not shape.has_text_frame:
                    continue

                for paragraph in shape.text_frame.paragraphs:

                    full_text = "".join(
                        run.text
                        for run in paragraph.runs
                    )

                    replaced = False

                    for col in df.columns:

                        placeholder = str(col).strip()

                        if placeholder in full_text:

                            full_text = full_text.replace(
                                placeholder,
                                str(row[col])
                            )

                            replaced = True

                    if replaced and paragraph.runs:

                        for run in paragraph.runs:
                            run.text = ""

                        paragraph.runs[0].text = full_text

        safe_name = str(row[df.columns[0]])

        bad_chars = [
            '/',
            '\\',
            ':',
            '*',
            '?',
            '"',
            '<',
            '>',
            '|'
        ]

        for char in bad_chars:
            safe_name = safe_name.replace(char, "")

        pptx_path = os.path.join(
            user_output_dir,
            f"{safe_name}.pptx"
        )

        prs.save(pptx_path)

        generated_files.append(pptx_path)

    zip_name = f"{user_id}_result.zip"

    zip_path = os.path.join(
        OUTPUT_DIR,
        zip_name
    )

    with zipfile.ZipFile(zip_path, "w") as zipf:

        for file in generated_files:

            zipf.write(
                file,
                arcname=os.path.basename(file)
            )

    return zip_name

# ==================== ЗАГРУЗКА ШАБЛОНА ====================

@app.post("/upload_template")
async def upload_template(request: Request):

    try:

        body = await request.body()

        try:
            data = await request.json()

        except Exception:

            return PlainTextResponse(
                f"Ошибка парсинга JSON:\n{body.decode()[:500]}"
            )

        if isinstance(data, list):

            variables = data
            user_id = "default"
            slot = 1

        else:

            variables = data.get("variables") or []

            contact = data.get("contact") or {}

            user_id = str(
                contact.get("id", "default")
            )

            slot = int(
                data.get("slot", 1)
            )

        if slot < 1 or slot > MAX_TEMPLATES:

            return PlainTextResponse(
                f"Слот должен быть от 1 до {MAX_TEMPLATES}"
            )

        file_url, ext = extract_file_url(
            variables,
            [".pptx"]
        )

        if not file_url:

            all_urls = []

            for var in variables:

                if isinstance(var, dict):

                    url = (
                        (var.get("payload") or {}).get("url")
                        or var.get("url")
                    )

                    if url:
                        all_urls.append(url)

            return PlainTextResponse(
                f"PPTX не найден.\nВсе URL: {all_urls or 'пусто'}"
            )

        filename, template_path = download_file(
            file_url,
            TEMPLATES_DIR,
            force_ext=".pptx"
        )

        try:

            prs = Presentation(template_path)

            slide_count = len(prs.slides)

        except Exception as e:

            return PlainTextResponse(
                f"Файл не является валидным PPTX:\n{e}"
            )

        templates = load_templates(user_id)

        if "slots" not in templates:
            templates["slots"] = {}

        templates["slots"][str(slot)] = {
            "path": template_path,
            "name": filename,
            "slides": slide_count
        }

        if "selected" not in templates:
            templates["selected"] = slot

        save_templates(
            user_id,
            templates
        )

        return PlainTextResponse(
            f"Шаблон сохранён в слот {slot} ✅\n"
            f"Файл: {filename} ({slide_count} слайдов)\n\n"
            f"Теперь отправь Excel файл (.xlsx)"
        )

    except Exception as e:

        return PlainTextResponse(
            f"Ошибка upload_template:\n{e}\n\n{traceback.format_exc()}"
        )

# ==================== СПИСОК ШАБЛОНОВ ====================

@app.post("/list_templates")
async def list_templates(request: Request):

    try:

        data = await request.json()

        if isinstance(data, dict):

            contact = data.get("contact") or {}

        else:

            contact = {}

        user_id = str(
            contact.get("id", "default")
        )

        templates = load_templates(user_id)

        slots = templates.get("slots", {})

        selected = templates.get("selected")

        if not slots:

            return PlainTextResponse(
                "У тебя нет сохранённых шаблонов"
            )

        lines = ["📁 Твои шаблоны:\n"]

        for i in range(1, MAX_TEMPLATES + 1):

            slot = slots.get(str(i))

            if slot:

                mark = ""

                if str(i) == str(selected):
                    mark = " ✅ (активный)"

                lines.append(
                    f"{i}. {slot['name']} — "
                    f"{slot['slides']} слайдов{mark}"
                )

            else:

                lines.append(f"{i}. — пусто")

        lines.append(
            f"\nОтправь номер шаблона "
            f"(1-{MAX_TEMPLATES})"
        )

        return PlainTextResponse(
            "\n".join(lines)
        )

    except Exception as e:

        return PlainTextResponse(
            f"Ошибка list_templates:\n{e}"
        )

# ==================== ВЫБОР ШАБЛОНА ====================

@app.post("/select_template")
async def select_template(request: Request):

    try:

        data = await request.json()

        contact = (
            data.get("contact") or {}
            if isinstance(data, dict)
            else {}
        )

        user_id = str(
            contact.get("id", "default")
        )

        slot = int(
            data.get("slot", 1)
        )

        templates = load_templates(user_id)

        slots = templates.get("slots", {})

        if str(slot) not in slots:

            return PlainTextResponse(
                f"Слот {slot} пустой"
            )

        templates["selected"] = slot

        save_templates(
            user_id,
            templates
        )

        name = slots[str(slot)]["name"]

        return PlainTextResponse(
            f"Выбран шаблон {slot}: {name} ✅"
        )

    except Exception as e:

        return PlainTextResponse(
            f"Ошибка select_template:\n{e}"
        )

# ==================== ЗАГРУЗКА EXCEL ====================

@app.post("/upload_excel")
async def upload_excel(request: Request):

    try:

        body = await request.body()

        try:

            data = await request.json()

        except Exception:

            return PlainTextResponse(
                f"Ошибка парсинга JSON:\n{body.decode()[:500]}"
            )

        if isinstance(data, list):

            variables = data
            user_id = "default"

        else:

            variables = data.get("variables") or []

            contact = data.get("contact") or {}

            user_id = str(
                contact.get("id", "default")
            )

        template_path = get_selected_template(user_id)

        if not template_path:

            return PlainTextResponse(
                "Сначала выбери шаблон"
            )

        if not os.path.exists(template_path):

            return PlainTextResponse(
                "Шаблон не найден на диске"
            )

        file_url, ext = extract_file_url(
            variables,
            [".xlsx", ".xls"]
        )

        if not file_url:

            all_urls = []

            for var in variables:

                if isinstance(var, dict):

                    url = (
                        (var.get("payload") or {}).get("url")
                        or var.get("url")
                    )

                    if url:
                        all_urls.append(url)

            return PlainTextResponse(
                f"Excel не найден.\n"
                f"Все URL: {all_urls or 'пусто'}"
            )

        filename, excel_path = download_file(
            file_url,
            EXCEL_DIR,
            force_ext=".xlsx"
        )

        zip_name = generate_pptx(
            template_path=template_path,
            excel_path=excel_path,
            user_id=user_id
        )

        full_url = (
            f"{BASE_URL}/files/{zip_name}"
        )

        return PlainTextResponse(
            f"Файлы готовы ✅\n\n{full_url}"
        )

    except Exception as e:

        return PlainTextResponse(
            f"Ошибка сервера:\n{e}\n\n{traceback.format_exc()}"
        )

# ==================== СТАТУС ====================

@app.get("/", response_class=PlainTextResponse)
def test():
    return "бот работает ✅"
