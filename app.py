from fastapi import FastAPI
from fastapi.responses import FileResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles

import requests
import os
import pandas as pd
from pptx import Presentation
import zipfile
import shutil
import uuid

app = FastAPI()

# ------------------ НАСТРОЙКИ ------------------

BASE_URL = "https://web-production-a9964.up.railway.app"

TEMP_DIR = "temp"
OUTPUT_DIR = "output"

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")

# ------------------ TEST ------------------

@app.get("/", response_class=PlainTextResponse)
def test():
    return "бот работает"

# ------------------ СКАЧАТЬ ФАЙЛ ------------------

def download_file(url, path):

    response = requests.get(url)

    with open(path, "wb") as f:
        f.write(response.content)

# ------------------ ГЕНЕРАЦИЯ PPTX ------------------

def generate_pptx(template_path, excel_path, output_folder):

    df = pd.read_excel(excel_path)

    generated_files = []

    for index, row in df.iterrows():

        prs = Presentation(template_path)

        for slide in prs.slides:

            for shape in slide.shapes:

                if not hasattr(shape, "text"):
                    continue

                text = shape.text

                for col in df.columns:

                    placeholder = str(col).strip()
                    value = str(row[col])

                    text = text.replace(
                        placeholder,
                        value
                    )

                shape.text = text

        # имя файла
        if "Название" in df.columns:
            safe_name = str(row["Название"])
        else:
            safe_name = f"file_{index + 1}"

        # очистка
        bad_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']

        for char in bad_chars:
            safe_name = safe_name.replace(char, "")

        pptx_name = f"{safe_name}.pptx"

        pptx_path = os.path.join(
            output_folder,
            pptx_name
        )

        prs.save(pptx_path)

        generated_files.append(pptx_path)

    return generated_files

# ------------------ GENERATE ------------------

@app.post("/generate")
async def generate(data: dict):

    try:

        pptx_url = data.get("pptx_url")
        excel_url = data.get("excel_url")
        user_id = str(data.get("user_id"))

        if not pptx_url:
            return {"error": "pptx_url missing"}

        if not excel_url:
            return {"error": "excel_url missing"}

        # уникальная папка
        uid = str(uuid.uuid4())

        user_temp = os.path.join(TEMP_DIR, uid)
        user_output = os.path.join(OUTPUT_DIR, uid)

        os.makedirs(user_temp, exist_ok=True)
        os.makedirs(user_output, exist_ok=True)

        # пути
        template_path = os.path.join(user_temp, "template.pptx")
        excel_path = os.path.join(user_temp, "data.xlsx")

        # скачиваем файлы
        download_file(pptx_url, template_path)
        download_file(excel_url, excel_path)

        # генерируем pptx
        generated_files = generate_pptx(
            template_path,
            excel_path,
            user_output
        )

        # zip
        zip_name = f"{user_id}_result.zip"

        zip_path = os.path.join(
            OUTPUT_DIR,
            zip_name
        )

        with zipfile.ZipFile(
            zip_path,
            "w",
            zipfile.ZIP_DEFLATED
        ) as zipf:

            for file in generated_files:

                zipf.write(
                    file,
                    arcname=os.path.basename(file)
                )

        zip_url = f"{BASE_URL}/files/{zip_name}"

        return {
            "success": True,
            "zip_url": zip_url
        }

    except Exception as e:

        return {
            "success": False,
            "error": str(e)
        }
