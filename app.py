from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
import pandas as pd
from pptx import Presentation
import os
import zipfile
import shutil
import uuid

app = FastAPI()

TEMP_DIR = "temp"
OUTPUT_DIR = "output"

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# временное хранилище пользователей
sessions = {}


@app.post("/upload_template")
async def upload_template(file: UploadFile = File(...)):
    user_id = str(uuid.uuid4())

    path = os.path.join(TEMP_DIR, f"{user_id}_template.pptx")

    with open(path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    sessions[user_id] = {"template": path}

    return {"user_id": user_id, "message": "Шаблон загружен"}


@app.post("/upload_excel")
async def upload_excel(user_id: str, file: UploadFile = File(...)):
    if user_id not in sessions:
        return JSONResponse({"error": "user_id не найден"}, status_code=400)

    path = os.path.join(TEMP_DIR, f"{user_id}_data.xlsx")

    with open(path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    sessions[user_id]["excel"] = path

    return {"message": "Excel загружен"}


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

    for i, row in df.iterrows():
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
