from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
from fastapi.staticfiles import StaticFiles

import requests
import os
import pandas as pd
from pptx import Presentation
import zipfile
import shutil

app = FastAPI()

# ------------------ НАСТРОЙКИ ------------------

BASE_URL = "https://ТВОЙ-ПРОЕКТ.up.railway.app"

TEMPLATES_DIR = "templates"
EXCEL_DIR = "excel"
OUTPUT_DIR = "output"
STATE_DIR = "state"

os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(STATE_DIR, exist_ok=True)

# static files
app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")

# ------------------ TEST ------------------

@app.get("/", response_class=PlainTextResponse)
def test():
    return "бот работает"

# ------------------ СКАЧИВАНИЕ ФАЙЛА ------------------

def download_file(file_url, save_dir):

    response = requests.get(file_url)

    filename = file_url.split("/")[-1].split("?")[0]

    path = os.path.join(save_dir, filename)

    with open(path, "wb") as f:
        f.write(response.content)

    return filename, path

# ------------------ ГЕНЕРАЦИЯ PPTX ------------------

def generate_pptx(template_path, excel_path, user_id):

    user_output_dir = os.path.join(
        OUTPUT_DIR,
        user_id
    )

    # очистка старых файлов
    if os.path.exists(user_output_dir):
        shutil.rmtree(user_output_dir)

    os.makedirs(user_output_dir, exist_ok=True)

    # excel
    df = pd.read_excel(excel_path)

    generated_files = []

    for _, row in df.iterrows():

        prs = Presentation(template_path)

        for slide in prs.slides:

            for shape in slide.shapes:

                if not shape.has_text_frame:
                    continue

                for paragraph in shape.text_frame.paragraphs:

                    text = paragraph.text

                    for col in df.columns:

                        placeholder = str(col)

                        if placeholder in text:

                            value = str(row[col])

                            text = text.replace(
                                placeholder,
                                value
                            )

                    paragraph.text = text

        # имя файла
        if "Название" in df.columns:
            safe_name = str(row["Название"])
        else:
            safe_name = f"file_{_+1}"

        # очистка имени
        bad_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']

        for char in bad_chars:
            safe_name = safe_name.replace(char, "")

        pptx_filename = f"{safe_name}.pptx"

        pptx_path = os.path.join(
            user_output_dir,
            pptx_filename
        )

        prs.save(pptx_path)

        generated_files.append(pptx_path)

    # ZIP
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

# ------------------ ЗАГРУЗКА ШАБЛОНА ------------------

@app.post("/upload_template")
async def upload_template(request: Request):

    data = await request.json()

    variables = data.get("variables", [])
    contact = data.get("contact", {})

    user_id = str(contact.get("id"))

    file_url = None

    for var in variables:

        if var.get("name") == "file":

            payload = var.get("payload", {})
            file_url = payload.get("url")

    if not file_url:

        return PlainTextResponse(
            "Файл не найден"
        )

    filename, template_path = download_file(
        file_url,
        TEMPLATES_DIR
    )

    # сохраняем последний шаблон пользователя
    state_file = os.path.join(
        STATE_DIR,
        f"{user_id}.txt"
    )

    with open(state_file, "w", encoding="utf-8") as f:
        f.write(template_path)

    return PlainTextResponse(
        "Шаблон загружен ✅\n\nТеперь отправь Excel файл (.xlsx)"
    )

# ------------------ ЗАГРУЗКА EXCEL ------------------

@app.post("/upload_excel")
async def upload_excel(request: Request):

    data = await request.json()

    variables = data.get("variables", [])
    contact = data.get("contact", {})

    user_id = str(contact.get("id"))

    # шаблон
    state_file = os.path.join(
        STATE_DIR,
        f"{user_id}.txt"
    )

    if not os.path.exists(state_file):

        return PlainTextResponse(
            "Сначала загрузи шаблон"
        )

    with open(state_file, "r", encoding="utf-8") as f:
        template_path = f.read()

    # excel
    file_url = None

    for var in variables:

        if var.get("name") == "file":

            payload = var.get("payload", {})
            file_url = payload.get("url")

    if not file_url:

        return PlainTextResponse(
            "Excel файл не найден"
        )

    filename, excel_path = download_file(
        file_url,
        EXCEL_DIR
    )

    # генерация
    zip_name = generate_pptx(
        template_path=template_path,
        excel_path=excel_path,
        user_id=user_id
    )

    full_url = f"{BASE_URL}/files/{zip_name}"

    return PlainTextResponse(
        f"Файлы готовы ✅\n\n{full_url}"
    )
