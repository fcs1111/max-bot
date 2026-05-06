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
import pymorphy2

app = FastAPI()
morph = pymorphy2.MorphAnalyzer()

BASE_URL = "https://web-production-a9964.up.railway.app"

TEMPLATES_DIR = "templates"
EXCEL_DIR    = "excel"
OUTPUT_DIR   = "output"
STATE_DIR    = "state"
MAX_TEMPLATES = 5

CASE_MAP = {
    "именительный": "nomn",
    "родительный":  "gent",
    "дательный":    "datv",
    "винительный":  "accs",
    "творительный": "ablt",
    "предложный":   "loct",
}

for d in [TEMPLATES_DIR, EXCEL_DIR, OUTPUT_DIR, STATE_DIR]:
    os.makedirs(d, exist_ok=True)

app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")


# ==================== СКЛОНЕНИЕ ====================

def is_cyrillic(text: str) -> bool:
    return any('а' <= ch.lower() <= 'я' or ch in 'ёЁ' for ch in text)

def inflect_word(word: str, case: str) -> str:
    if not is_cyrillic(word):
        return word
    parsed = morph.parse(word)
    if not parsed:
        return word
    inflected = parsed[0].inflect({case})
    if not inflected:
        return word
    result = inflected.word
    if word[0].isupper():
        result = result.capitalize()
    return result

def inflect_phrase(phrase: str, case: str) -> str:
    return " ".join(inflect_word(w, case) for w in phrase.split())

def apply_case_to_row(row: pd.Series, case_tag: str) -> pd.Series:
    """Склоняет все строковые значения строки."""
    new_row = row.copy()
    for col in row.index:
        val = str(row[col])
        if is_cyrillic(val):
            new_row[col] = inflect_phrase(val, case_tag)
    return new_row


# ==================== ШАБЛОНЫ (5 слотов) ====================

def get_templates_path(user_id: str) -> str:
    return os.path.join(STATE_DIR, f"{user_id}_templates.json")


def load_templates(user_id: str) -> dict:
    path = get_templates_path(user_id)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_templates(user_id: str, templates: dict):
    path = get_templates_path(user_id)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(templates, f, ensure_ascii=False, indent=2)


def get_selected_template(user_id: str) -> str | None:
    """Возвращает путь к выбранному шаблону."""
    templates = load_templates(user_id)
    selected = templates.get("selected")
    if not selected:
        return None
    slot = templates.get("slots", {}).get(str(selected))
    return slot.get("path") if slot else None


# ==================== СКАЧИВАНИЕ ====================

def download_file(file_url: str, save_dir: str, force_ext: str = None):
    response = requests.get(file_url, timeout=30)
    response.raise_for_status()
    filename = file_url.split("/")[-1].split("?")[0]
    if force_ext and not filename.lower().endswith(force_ext):
        filename = filename + force_ext
    path = os.path.join(save_dir, filename)
    with open(path, "wb") as f:
        f.write(response.content)
    return filename, path


def extract_file_url(variables: list, allowed_extensions: list):
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
            clean = url.split("?")[0].lower()
            for ext in allowed_extensions:
                if clean.endswith(ext):
                    found_urls.append((url, ext))
    return found_urls[-1] if found_urls else (None, None)


# ==================== ГЕНЕРАЦИЯ ====================

# generate_pptx — добавляем параметр case_name
def generate_pptx(template_path: str, excel_path: str, user_id: str, case_name: str = "именительный") -> str:

    user_output_dir = os.path.join(OUTPUT_DIR, user_id)
    if os.path.exists(user_output_dir):
        shutil.rmtree(user_output_dir)
    os.makedirs(user_output_dir, exist_ok=True)

    df = pd.read_excel(excel_path)

    case_tag = CASE_MAP.get(case_name.lower().strip())  # None = именительный

    generated_files = []

    for i, row in df.iterrows():
        # Применяем падеж если выбран не именительный
        if case_tag:
            row = apply_case_to_row(row, case_tag)

        prs = Presentation(template_path)

        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for paragraph in shape.text_frame.paragraphs:
                    full_text = "".join(run.text for run in paragraph.runs)
                    replaced = False
                    for col in df.columns:
                        placeholder = str(col).strip()
                        if placeholder in full_text:
                            full_text = full_text.replace(placeholder, str(row[col]))
                            replaced = True
                    if replaced and paragraph.runs:
                        for run in paragraph.runs:
                            run.text = ""
                        paragraph.runs[0].text = full_text

        safe_name = str(row[df.columns[0]])
        for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
            safe_name = safe_name.replace(char, "")

        pptx_path = os.path.join(user_output_dir, f"{safe_name}.pptx")
        prs.save(pptx_path)
        generated_files.append(pptx_path)

    zip_name = f"{user_id}_result.zip"
    zip_path = os.path.join(OUTPUT_DIR, zip_name)
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for file in generated_files:
            zipf.write(file, arcname=os.path.basename(file))

    return zip_name

# ==================== РОУТЫ ====================

@app.get("/", response_class=PlainTextResponse)
def root():
    return "бот работает ✅"


# --- Загрузка шаблона в слот ---
@app.post("/upload_template")
async def upload_template(request: Request):
    try:
        data = await request.json()
        if isinstance(data, list):
            variables, user_id, slot = data, "default", 1
        else:
            variables = data.get("variables") or []
            contact   = data.get("contact") or {}
            user_id   = str(contact.get("id", "default"))
            # номер слота передаётся в теле: "slot": 1
            slot = int(data.get("slot", 1))

        if slot < 1 or slot > MAX_TEMPLATES:
            return PlainTextResponse(f"Слот должен быть от 1 до {MAX_TEMPLATES}")

        file_url, _ = extract_file_url(variables, [".pptx"])
        if not file_url:
            return PlainTextResponse("PPTX файл не найден в запросе")

        filename, template_path = download_file(file_url, TEMPLATES_DIR, force_ext=".pptx")

        try:
            prs = Presentation(template_path)
            slide_count = len(prs.slides)
        except Exception as e:
            return PlainTextResponse(f"Невалидный PPTX:\n{e}")

        templates = load_templates(user_id)
        if "slots" not in templates:
            templates["slots"] = {}

        templates["slots"][str(slot)] = {
            "path": template_path,
            "name": filename,
            "slides": slide_count
        }
        # автовыбор если ещё ничего не выбрано
        if "selected" not in templates:
            templates["selected"] = slot

        save_templates(user_id, templates)

        return PlainTextResponse(
            f"Шаблон сохранён в слот {slot} ✅\n"
            f"Файл: {filename} ({slide_count} слайдов)\n\n"
            f"Теперь отправь Excel файл (.xlsx)"
        )

    except Exception as e:
        return PlainTextResponse(f"Ошибка upload_template:\n{e}\n\n{traceback.format_exc()}")


# --- Список шаблонов ---
@app.post("/list_templates")
async def list_templates(request: Request):
    try:
        data = await request.json()
        contact = data.get("contact") or {} if isinstance(data, dict) else {}
        user_id = str(contact.get("id", "default"))

        templates = load_templates(user_id)
        slots = templates.get("slots", {})
        selected = templates.get("selected")

        if not slots:
            return PlainTextResponse("У тебя нет сохранённых шаблонов.\n\nЗагрузи шаблон через /upload_template")

        lines = ["📁 Твои шаблоны:\n"]
        for i in range(1, MAX_TEMPLATES + 1):
            slot = slots.get(str(i))
            if slot:
                mark = " ✅ (активный)" if str(i) == str(selected) else ""
                lines.append(f"{i}. {slot['name']} — {slot['slides']} слайдов{mark}")
            else:
                lines.append(f"{i}. — пусто")

        lines.append(f"\nОтправь номер шаблона чтобы выбрать (1–{MAX_TEMPLATES})")
        return PlainTextResponse("\n".join(lines))

    except Exception as e:
        return PlainTextResponse(f"Ошибка list_templates:\n{e}")


# --- Выбор шаблона ---
@app.post("/select_template")
async def select_template(request: Request):
    try:
        data = await request.json()
        contact = data.get("contact") or {} if isinstance(data, dict) else {}
        user_id = str(contact.get("id", "default"))
        slot    = int(data.get("slot", 1))

        templates = load_templates(user_id)
        slots = templates.get("slots", {})

        if str(slot) not in slots:
            return PlainTextResponse(f"Слот {slot} пустой. Сначала загрузи шаблон.")

        templates["selected"] = slot
        save_templates(user_id, templates)

        name = slots[str(slot)]["name"]
        return PlainTextResponse(f"Выбран шаблон {slot}: {name} ✅")

    except Exception as e:
        return PlainTextResponse(f"Ошибка select_template:\n{e}")


# --- Загрузка Excel и генерация ---
# upload_excel — принимает case из тела запроса
@app.post("/upload_excel")
async def upload_excel(request: Request):
    try:
        data = await request.json()
        if isinstance(data, list):
            variables, user_id, case_name = data, "default", "именительный"
        else:
            variables  = data.get("variables") or []
            contact    = data.get("contact") or {}
            user_id    = str(contact.get("id", "default"))
            case_name  = str(data.get("case", "именительный"))

        template_path = get_selected_template(user_id)
        if not template_path or not os.path.exists(template_path):
            return PlainTextResponse("Шаблон не найден. Выбери шаблон из меню.")

        file_url, _ = extract_file_url(variables, [".xlsx", ".xls"])
        if not file_url:
            return PlainTextResponse("Excel файл не найден")

        _, excel_path = download_file(file_url, EXCEL_DIR, force_ext=".xlsx")

        zip_name = generate_pptx(template_path, excel_path, user_id, case_name)
        full_url = f"{BASE_URL}/files/{zip_name}"

        case_label = case_name.capitalize() if case_name != "именительный" else "Именительный (без изменений)"

        return PlainTextResponse(
            f"Файлы готовы ✅\n"
            f"Падеж: {case_label}\n\n"
            f"📦 Скачай архив:\n{full_url}"
        )

    except Exception as e:
        return PlainTextResponse(f"Ошибка сервера:\n{e}\n\n{traceback.format_exc()}")
