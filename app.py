from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
from fastapi.staticfiles import StaticFiles

import requests
import os
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

os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(STATE_DIR, exist_ok=True)

app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")

# ------------------ СКАЧИВАНИЕ ------------------

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

# ------------------ ИЗВЛЕЧЕНИЕ URL ------------------

def extract_file_url(variables: list, allowed_extensions: list):
    found_urls = []

    for var in variables:
        if not var:
            continue

        # variables может быть списком или словарём
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

# ------------------ ГЕНЕРАЦИЯ PPTX ------------------

def generate_pptx(template_path: str, excel_path: str, user_id: str) -> str:

    user_output_dir = os.path.join(OUTPUT_DIR, user_id)

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

                    # Собираем полный текст параграфа из всех runs
                    full_text = "".join(run.text for run in paragraph.runs)

                    # Проверяем есть ли хоть один плейсхолдер
                    replaced = False
                    for col in df.columns:
                        placeholder = str(col).strip()
                        if placeholder in full_text:
                            full_text = full_text.replace(placeholder, str(row[col]))
                            replaced = True

                    if replaced and paragraph.runs:
                        # Очищаем все runs кроме первого
                        for run in paragraph.runs:
                            run.text = ""
                        # Кладём итоговый текст в первый run
                        paragraph.runs[0].text = full_text

        # Имя файла — первая колонка
        safe_name = str(row[df.columns[0]])
        bad_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        for char in bad_chars:
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

# ------------------ ЗАГРУЗКА ШАБЛОНА ------------------

@app.post("/upload_template")
async def upload_template(request: Request):
    try:
        body = await request.body()
        try:
            data = await request.json()
        except Exception:
            return PlainTextResponse(f"Ошибка парсинга JSON:\n{body.decode()[:500]}")

        if isinstance(data, list):
            variables = data
            user_id = "default"
        else:
            variables = data.get("variables") or []
            contact = data.get("contact") or {}
            user_id = str(contact.get("id", "default"))

        file_url, ext = extract_file_url(variables, [".pptx"])

        if not file_url:
            all_urls = []
            for var in variables:
                if isinstance(var, dict):
                    url = (var.get("payload") or {}).get("url") or var.get("url")
                    if url:
                        all_urls.append(url)
            return PlainTextResponse(
                f"PPTX не найден.\nВсе URL: {all_urls or 'пусто'}"
            )

        filename, template_path = download_file(file_url, TEMPLATES_DIR, force_ext=".pptx")

        # Валидация
        try:
            prs = Presentation(template_path)
            slide_count = len(prs.slides)
        except Exception as e:
            return PlainTextResponse(f"Файл не является валидным PPTX:\n{e}")

        state_file = os.path.join(STATE_DIR, f"{user_id}.txt")
        with open(state_file, "w", encoding="utf-8") as f:
            f.write(template_path)

        return PlainTextResponse(
            f"Шаблон загружен ✅\n"
            f"Файл: {filename} ({slide_count} слайдов)\n\n"
            f"Теперь отправь Excel файл (.xlsx)"
        )

    except Exception as e:
        return PlainTextResponse(f"Ошибка upload_template:\n{e}\n\n{traceback.format_exc()}")

# ------------------ ЗАГРУЗКА EXCEL ------------------

@app.post("/upload_excel")
async def upload_excel(request: Request):
    try:
        body = await request.body()
        try:
            data = await request.json()
        except Exception:
            return PlainTextResponse(f"Ошибка парсинга JSON:\n{body.decode()[:500]}")

        if isinstance(data, list):
            variables = data
            user_id = "default"
        else:
            variables = data.get("variables") or []
            contact = data.get("contact") or {}
            user_id = str(contact.get("id", "default"))

        state_file = os.path.join(STATE_DIR, f"{user_id}.txt")
        if not os.path.exists(state_file):
            return PlainTextResponse("Сначала загрузи шаблон PPTX")

        with open(state_file, "r", encoding="utf-8") as f:
            template_path = f.read().strip()

        if not os.path.exists(template_path):
            return PlainTextResponse("Шаблон не найден на диске. Загрузи шаблон заново.")

        file_url, ext = extract_file_url(variables, [".xlsx", ".xls"])

        if not file_url:
            all_urls = []
            for var in variables:
                if isinstance(var, dict):
                    url = (var.get("payload") or {}).get("url") or var.get("url")
                    if url:
                        all_urls.append(url)
            return PlainTextResponse(
                f"Excel не найден.\nВсе URL: {all_urls or 'пусто'}"
            )

        filename, excel_path = download_file(file_url, EXCEL_DIR, force_ext=".xlsx")

        zip_name = generate_pptx(
            template_path=template_path,
            excel_path=excel_path,
            user_id=user_id
        )

        full_url = f"{BASE_URL}/files/{zip_name}"
        return PlainTextResponse(f"Файлы готовы ✅\n\n{full_url}")

    except Exception as e:
        return PlainTextResponse(f"Ошибка сервера:\n{e}\n\n{traceback.format_exc()}")

# ------------------ СТАТУС ------------------

@app.get("/", response_class=PlainTextResponse)
def test():
    return "бот работает ✅"
