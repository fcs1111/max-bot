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

BASE_URL = "https://web-production-a9964.up.railway.app"

WATBOT_WEBHOOK_URL = "https://api.watbot.ru/hook/4727260:QbeXNM2mi0XbghRAI22tDKuONEZjErl6kEJVwZ1D67VGvfup"

TEMP_DIR = "temp"
OUTPUT_DIR = "output"
STATE_DIR = "state"

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(STATE_DIR, exist_ok=True)

app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")

# ------------------ TEST ------------------

@app.get("/", response_class=PlainTextResponse)
def home():
    return "бот работает"

# ------------------ ОТПРАВКА В WATBOT ------------------

def send_message(chat_id, text):

    payload = {
        "chat_id": str(chat_id),
        "text": text
    }

    response = requests.post(
        WATBOT_WEBHOOK_URL,
        json=payload
    )

    print("WATBOT STATUS:")
    print(response.status_code)

    print("WATBOT RESPONSE:")
    print(response.text)

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

        # очистка имени
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

# ------------------ WEBHOOK ------------------

@app.post("/")
async def webhook(request: Request):

    try:

        data = await request.json()

        print("WEBHOOK DATA:")
        print(data)

        message = data.get("message") or {}
        sender = message.get("sender") or {}

        chat_id = sender.get("user_id")

        attachments = (
            message.get("body", {})
            .get("attachments", [])
        )

        if not attachments:

            send_message(
                chat_id,
                "Отправь PPTX или Excel файл"
            )

            return {"ok": True}

        attachment = attachments[0]

        payload = attachment.get("payload") or {}

        file_url = payload.get("url")
        filename = payload.get("filename", "")

        if not file_url:

            send_message(
                chat_id,
                "Файл не найден"
            )

            return {"ok": True}

        # ------------------ PPTX ------------------

        if filename.lower().endswith(".pptx"):

            template_path = os.path.join(
                STATE_DIR,
                f"{chat_id}_template.pptx"
            )

            download_file(
                file_url,
                template_path
            )

            send_message(
                chat_id,
                "Шаблон сохранен ✅\n\nТеперь отправь Excel файл"
            )

            return {"ok": True}

        # ------------------ EXCEL ------------------

        if (
            filename.lower().endswith(".xlsx")
            or filename.lower().endswith(".xls")
        ):

            template_path = os.path.join(
                STATE_DIR,
                f"{chat_id}_template.pptx"
            )

            if not os.path.exists(template_path):

                send_message(
                    chat_id,
                    "Сначала отправь PPTX шаблон"
                )

                return {"ok": True}

            user_temp = os.path.join(
                TEMP_DIR,
                str(chat_id)
            )

            user_output = os.path.join(
                OUTPUT_DIR,
                str(chat_id)
            )

            if os.path.exists(user_temp):
                shutil.rmtree(user_temp)

            if os.path.exists(user_output):
                shutil.rmtree(user_output)

            os.makedirs(user_temp, exist_ok=True)
            os.makedirs(user_output, exist_ok=True)

            excel_path = os.path.join(
                user_temp,
                "data.xlsx"
            )

            download_file(
                file_url,
                excel_path
            )

            generated_files = generate_pptx(
                template_path,
                excel_path,
                user_output
            )

            zip_name = f"{chat_id}_result.zip"

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

            send_message(
                chat_id,
                f"Готово ✅\n\nСкачать ZIP:\n{zip_url}"
            )

            return {"ok": True}

        send_message(
            chat_id,
            "Неподдерживаемый формат файла"
        )

        return {"ok": True}

    except Exception as e:

        print("ERROR:")
        print(str(e))

        return {
            "ok": False,
            "error": str(e)
        }
