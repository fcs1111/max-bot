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

os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# static files
app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")

# ------------------ ПАМЯТЬ ------------------

templates_db = {}
user_state = {}

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

    # папка пользователя
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

    # генерация
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
            safe_name = str(row[df.columns[0]])

        # очистка имени
        safe_name = safe_name.replace("/", "")
        safe_name = safe_name.replace("\\", "")
        safe_name = safe_name.replace(":", "")
        safe_name = safe_name.replace("*", "")
        safe_name = safe_name.replace("?", "")
        safe_name = safe_name.replace('"', "")
        safe_name = safe_name.replace("<", "")
        safe_name = safe_name.replace(">", "")
        safe_name = safe_name.replace("|", "")

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
        return PlainTextResponse("Файл не найден")

    filename, path = download_file(
        file_url,
        TEMPLATES_DIR
    )

    if user_id not in templates_db:
        templates_db[user_id] = []

    templates_db[user_id].append({
        "name": filename,
        "path": path
    })

    return PlainTextResponse(
        f"Шаблон загружен ✅\n\n{filename}"
    )

# ------------------ СПИСОК ШАБЛОНОВ ------------------

@app.post("/get_templates")
async def get_templates(request: Request):

    data = await request.json()

    contact = data.get("contact", {})
    user_id = str(contact.get("id"))

    user_templates = templates_db.get(user_id, [])

    if not user_templates:
        return PlainTextResponse(
            "У тебя пока нет шаблонов"
        )

    text = "Твои шаблоны:\n\n"

    for i, template in enumerate(user_templates, start=1):

        text += f"{i}. {template['name']}\n"

    text += "\nОтправь номер шаблона"

    return PlainTextResponse(text)

# ------------------ ВЫБОР ШАБЛОНА ------------------

@app.post("/select_template")
async def select_template(request: Request):

    data = await request.json()

    text = data.get("message", {}).get("text", "")
    contact = data.get("contact", {})

    user_id = str(contact.get("id"))

    user_templates = templates_db.get(user_id, [])

    if not user_templates:
        return PlainTextResponse(
            "Шаблоны не найдены"
        )

    try:
        template_number = int(text)

    except:
        return PlainTextResponse(
            "Отправь номер шаблона"
        )

    if template_number < 1 or template_number > len(user_templates):

        return PlainTextResponse(
            "Неверный номер шаблона"
        )

    selected_template = user_templates[
        template_number - 1
    ]

    if user_id not in user_state:
        user_state[user_id] = {}

    user_state[user_id]["selected_template"] = selected_template["path"]

    return PlainTextResponse(
        f"Выбран шаблон ✅\n\n{selected_template['name']}\n\nТеперь отправь Excel файл (.xlsx)"
    )

# ------------------ ЗАГРУЗКА EXCEL ------------------

@app.post("/upload_excel")
async def upload_excel(request: Request):

    data = await request.json()

    variables = data.get("variables", [])
    contact = data.get("contact", {})

    user_id = str(contact.get("id"))

    # проверка шаблона
    state = user_state.get(user_id, {})

    selected_template = state.get("selected_template")

    if not selected_template:
        return PlainTextResponse(
            "Сначала выбери шаблон"
        )

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

    user_state[user_id]["excel"] = excel_path

    # генерация
    zip_name = generate_pptx(
        template_path=selected_template,
        excel_path=excel_path,
        user_id=user_id
    )

    full_url = f"{BASE_URL}/files/{zip_name}"

    return PlainTextResponse(
        f"Файлы готовы ✅\n\n{full_url}"
    )
