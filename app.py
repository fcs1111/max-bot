from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
import requests
import os
import pandas as pd
from pptx import Presentation
import zipfile

app = FastAPI()

# папки
TEMPLATES_DIR = "templates"
OUTPUT_DIR = "output"

os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# "база"
templates_db = {}
user_state = {}

# ------------------ TEST ------------------

@app.get("/", response_class=PlainTextResponse)
def test():
    return "бот работает"

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
        return {"message": "Файл не найден"}

    response = requests.get(file_url)
    filename = file_url.split("/")[-1].split("?")[0]

    path = os.path.join(TEMPLATES_DIR, filename)

    with open(path, "wb") as f:
        f.write(response.content)

    if user_id not in templates_db:
        templates_db[user_id] = []

    templates_db[user_id].append({
        "name": filename,
        "path": path
    })

    return {"message": "Шаблон сохранён"}

# ------------------ СПИСОК ШАБЛОНОВ ------------------

@app.post("/get_templates")
async def get_templates(request: Request):
    data = await request.json()
    contact = data.get("contact", {})
    user_id = str(contact.get("id"))

    user_templates = templates_db.get(user_id, [])

    if not user_templates:
        return {"message": "У тебя нет шаблонов"}

    names = [t["name"] for t in user_templates]

    return {"templates": names}

# ------------------ ВЫБОР ШАБЛОНА ------------------

@app.post("/select_template")
async def select_template(request: Request):
    data = await request.json()

    text = data.get("message", {}).get("text", "")
    contact = data.get("contact", {})
    user_id = str(contact.get("id"))

    user_templates = templates_db.get(user_id, [])

    for t in user_templates:
        if t["name"] == text:
            if user_id not in user_state:
                user_state[user_id] = {}

            user_state[user_id]["selected_template"] = t["path"]
            return {"message": f"Выбран шаблон {text}"}

    return {"message": "Шаблон не найден"}

# ------------------ ЗАГРУЗКА EXCEL ------------------

@app.post("/upload_excel")
async def upload_excel(request: Request):
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
        return {"message": "Excel не найден"}

    response = requests.get(file_url)
    filename = file_url.split("/")[-1].split("?")[0]

    path = os.path.join(TEMPLATES_DIR, filename)

    with open(path, "wb") as f:
        f.write(response.content)

    if user_id not in user_state:
        user_state[user_id] = {}

    user_state[user_id]["excel"] = path

    return {"message": "Excel загружен"}

# ------------------ ГЕНЕРАЦИЯ ------------------

@app.post("/generate")
async def generate(request: Request):
    data = await request.json()
    contact = data.get("contact", {})
    user_id = str(contact.get("id"))

    state = user_state.get(user_id, {})

    template_path = state.get("selected_template")
    excel_path = state.get("excel")

    if not template_path or not excel_path:
        return {"message": "Сначала выбери шаблон и загрузи Excel"}

    df = pd.read_excel(excel_path)
    generated_files = []

    for _, row in df.iterrows():
        prs = Presentation(template_path)

        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue

                for paragraph in shape.text_frame.paragraphs:
                    for col in df.columns:
                        if col in paragraph.text:
                            paragraph.text = paragraph.text.replace(
                                col, str(row[col])
                            )

        safe_name = str(row[df.columns[0]]).replace("/", "").replace("\\", "")
        filename = os.path.join(OUTPUT_DIR, f"{safe_name}.pptx")

        prs.save(filename)
        generated_files.append(filename)

    zip_path = os.path.join(OUTPUT_DIR, f"{user_id}_result.zip")

    with zipfile.ZipFile(zip_path, "w") as zipf:
        for file in generated_files:
            zipf.write(file)

    return {"message": "Файлы сгенерированы"}
