from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import FileResponse, JSONResponse
import pandas as pd
from pptx import Presentation
import os
import zipfile
import shutil
import uuid
import requests

app = FastAPI()

from fastapi.responses import PlainTextResponse

@app.get("/", response_class=PlainTextResponse)
def test():
    return "бот работает"
# ------------------ Директории ------------------
TEMP_DIR = "temp"
OUTPUT_DIR = "output"

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ------------------ Хранилище сессий ------------------
sessions = {}

# ------------------ Токен и Webhook ------------------
MAX_TOKEN = "f9LHodD0cOIivtqm-8fTlt28_L0RokxyCNiOPUhjiWD2JxYKIIxDgLLUOFKGnUujUpEP63GWeppEZH302YfZ"
WEBHOOK_URL = "https://web-production-a9964.up.railway.app/webhook"

# Регистрируем webhook при старте сервера
def register_webhook():
    try:
        response = requests.post(
            "https://api.max.ru/bot/setWebhook",
            json={"url": WEBHOOK_URL},
            headers={"Authorization": f"Bearer {MAX_TOKEN}"}
        )
        print("Webhook registration status:", response.json())
    except Exception as e:
        print("Webhook registration failed:", e)

register_webhook()

# ------------------ Функция отправки сообщений ------------------
def send_message(chat_id, text):
    url = "https://platform-api.max.ru/messages"
    headers = {"Authorization": f"Bearer {MAX_TOKEN}"}
    data = {"chat_id": chat_id, "text": text}
    try:
        requests.post(url, json=data, headers=headers)
    except Exception as e:
        print("Failed to send message:", e)

# ------------------ Webhook endpoint ------------------
@app.post("/webhook")
async def webhook(request: Request):
    data = await request.json()
    print("Incoming:", data)

    message = data.get("message", {})
    text = message.get("text", "")
    chat_id = message.get("chat_id")

    if not chat_id:
        return {"status": "no chat_id"}

    # Простейшая логика
    if text == "/start":
        send_message(chat_id, "Привет! Отправь шаблон PPTX через API.")
    else:
        send_message(chat_id, f"Ты написал: {text}")

    return {"status": "ok"}

# ------------------ Загрузка шаблона ------------------
@app.post("/upload_template")
async def upload_template(file: UploadFile = File(...)):
    user_id = str(uuid.uuid4())
    path = os.path.join(TEMP_DIR, f"{user_id}_template.pptx")

    with open(path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    sessions[user_id] = {"template": path}
    return {"user_id": user_id, "message": "Шаблон загружен"}

# ------------------ Загрузка Excel ------------------
@app.post("/upload_excel")
async def upload_excel(user_id: str, file: UploadFile = File(...)):
    if user_id not in sessions:
        return JSONResponse({"error": "user_id не найден"}, status_code=400)

    path = os.path.join(TEMP_DIR, f"{user_id}_data.xlsx")
    with open(path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    sessions[user_id]["excel"] = path
    return {"message": "Excel загружен"}

# ------------------ Генерация PPTX ------------------
@app.post("/generate")
async def generate(user_id: str):
    if user_id not in sessions:
        return JSONResponse({"error": "user_id не найден"}, status_code=400)
    if "template" not in sessions[user_id] or "excel" not in sessions[user_id]:
        return JSONResponse({"error": "не все файлы загружены"}, status_code=400)

    template_path = sessions[user_id]["template"]
    excel_path = sessions[user_id]["excel"]

    df = pd.read_excel(excel_path)
    generated_files = []

    for _, row in df.iterrows():
        prs = Presentation(template_path)
        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for paragraph in shape.text_frame.paragraphs:
                    full_text = "".join(run.text for run in paragraph.runs)
                    for col in df.columns:
                        placeholder = str(col)
                        if placeholder in full_text:
                            new_text = full_text.replace(placeholder, str(row[col]))
                            for run in paragraph.runs:
                                run.text = ""
                            if paragraph.runs:
                                paragraph.runs[0].text = new_text

        safe_name = str(row[df.columns[0]]).replace("/", "").replace("\\", "")
        filename = os.path.join(OUTPUT_DIR, f"{safe_name}.pptx")
        prs.save(filename)
        generated_files.append(filename)

    zip_path = os.path.join(OUTPUT_DIR, f"{user_id}_result.zip")
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in generated_files:
            zipf.write(file)

    return FileResponse(zip_path, filename="result.zip")
