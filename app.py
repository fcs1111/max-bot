from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse, FileResponse
import requests
import os
import glob
import pandas as pd
from pptx import Presentation
import zipfile

app = FastAPI()

# =========================================================
# ПАПКИ
# =========================================================

TEMPLATES_DIR = "templates"
EXCEL_DIR = "excel"
OUTPUT_DIR = "output"

os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# =========================================================
# ТЕСТ
# =========================================================

@app.get("/", response_class=PlainTextResponse)
def test():
    return "бот работает"

# =========================================================
# ЗАГРУЗКА ШАБЛОНА
# =========================================================

@app.post("/upload_template")
async def upload_template(request: Request):

    try:

        data = await request.json()

        print("UPLOAD TEMPLATE DATA:", data)

        variables = data.get("variables", [])

        file_url = None

        # ищем файл
        for var in variables:

            if var.get("name") == "file":

                payload = var.get("payload", {})

                file_url = payload.get("url")

        if not file_url:

            return PlainTextResponse(
                "Файл шаблона не найден"
            )

        # скачиваем
        response = requests.get(file_url)
        response.raise_for_status()

        filename = file_url.split("/")[-1].split("?")[0]

        file_path = os.path.join(
            TEMPLATES_DIR,
            filename
        )

        # сохраняем
        with open(file_path, "wb") as f:
            f.write(response.content)

        return PlainTextResponse(
            f"Шаблон {filename} загружен ✅"
        )

    except Exception as e:

        print("UPLOAD TEMPLATE ERROR:", str(e))

        return PlainTextResponse(
            f"Ошибка: {str(e)}"
        )

# =========================================================
# ЗАГРУЗКА EXCEL
# =========================================================

@app.post("/upload_excel")
async def upload_excel(request: Request):

    try:

        data = await request.json()

        print("UPLOAD EXCEL DATA:", data)

        variables = data.get("variables", [])

        file_url = None

        # ищем excel
        for var in variables:

            if var.get("name") == "file":

                payload = var.get("payload", {})

                file_url = payload.get("url")

        if not file_url:

            return PlainTextResponse(
                "Excel файл не найден"
            )

        # скачиваем
        response = requests.get(file_url)
        response.raise_for_status()

        filename = file_url.split("/")[-1].split("?")[0]

        file_path = os.path.join(
            EXCEL_DIR,
            filename
        )

        # сохраняем excel
        with open(file_path, "wb") as f:
            f.write(response.content)

        return PlainTextResponse(
            f"Excel {filename} загружен ✅"
        )

    except Exception as e:

        print("UPLOAD EXCEL ERROR:", str(e))

        return PlainTextResponse(
            f"Ошибка: {str(e)}"
        )

# =========================================================
# ГЕНЕРАЦИЯ
# =========================================================

@app.post("/generate")
async def generate():

    try:

        # берем последний шаблон
        template_files = sorted(
            glob.glob("templates/*.pptx"),
            key=os.path.getmtime
        )

        if not template_files:

            return PlainTextResponse(
                "Сначала загрузи шаблон"
            )

        latest_template = template_files[-1]

        # берем последний excel
        excel_files = sorted(
            glob.glob("excel/*.xlsx"),
            key=os.path.getmtime
        )

        if not excel_files:

            return PlainTextResponse(
                "Excel не найден"
            )

        latest_excel = excel_files[-1]

        # читаем excel
        df = pd.read_excel(latest_excel)

        generated_files = []

        # генерация
        for _, row in df.iterrows():

            prs = Presentation(latest_template)

            for slide in prs.slides:

                for shape in slide.shapes:

                    if not shape.has_text_frame:
                        continue

                    for paragraph in shape.text_frame.paragraphs:

                        text = paragraph.text

                        # замена плейсхолдеров
                        for col in df.columns:

                            placeholder = f"%{col}%"

                            if placeholder in text:

                                text = text.replace(
                                    placeholder,
                                    str(row[col])
                                )

                        paragraph.text = text

            # имя файла
            safe_name = str(
                row[df.columns[0]]
            ).replace("/", "").replace("\\", "")

            output_file = os.path.join(
                OUTPUT_DIR,
                f"{safe_name}.pptx"
            )

            prs.save(output_file)

            generated_files.append(output_file)

        # создаем zip
        zip_filename = "result.zip"

        zip_path = os.path.join(
            OUTPUT_DIR,
            zip_filename
        )

        with zipfile.ZipFile(zip_path, "w") as zipf:

            for file in generated_files:

                zipf.write(
                    file,
                    os.path.basename(file)
                )

        # ссылка
        download_url = (
            "https://web-production-a9964.up.railway.app"
            "/download/result.zip"
        )

        return PlainTextResponse(
            f"ZIP готов ✅\n{download_url}"
        )

    except Exception as e:

        print("GENERATE ERROR:", str(e))

        return PlainTextResponse(
            f"Ошибка генерации: {str(e)}"
        )

# =========================================================
# СКАЧИВАНИЕ ZIP
# =========================================================

@app.get("/download/{filename}")
async def download_file(filename: str):

    file_path = os.path.join(
        OUTPUT_DIR,
        filename
    )

    if not os.path.exists(file_path):

        return PlainTextResponse(
            "Файл не найден"
        )

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/zip"
    )
