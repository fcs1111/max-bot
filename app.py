from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import pandas as pd
from pptx import Presentation
import os
import zipfile
import shutil

app = FastAPI()

TEMP_DIR = "temp"
OUTPUT_DIR = "output"

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.post("/generate")
async def generate(template: UploadFile = File(...), excel: UploadFile = File(...)):
    template_path = os.path.join(TEMP_DIR, template.filename)
    excel_path = os.path.join(TEMP_DIR, excel.filename)

    # сохраняем файлы
    with open(template_path, "wb") as f:
        shutil.copyfileobj(template.file, f)

    with open(excel_path, "wb") as f:
        shutil.copyfileobj(excel.file, f)

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

    zip_path = os.path.join(OUTPUT_DIR, "result.zip")

    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in generated_files:
            zipf.write(file)

    return FileResponse(zip_path, filename="result.zip")