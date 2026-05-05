from fastapi import FastAPI, Form
from fastapi.responses import PlainTextResponse
import requests
import os

app = FastAPI()

# папки
TEMPLATES_DIR = "templates"
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# тест (чтобы бот отвечал)
@app.get("/", response_class=PlainTextResponse)
def test():
    return "бот работает"


# загрузка шаблона
@app.post("/upload_template")
async def upload_template(file_url: str = Form(...)):
    try:
        # скачиваем файл по ссылке
        response = requests.get(file_url)
        response.raise_for_status()

        # имя файла
        filename = file_url.split("/")[-1]

        # путь сохранения
        path = os.path.join(TEMPLATES_DIR, filename)

        # сохраняем файл
        with open(path, "wb") as f:
            f.write(response.content)

        return {"message": f"Шаблон {filename} загружен"}

    except Exception as e:
        return {"message": f"Ошибка загрузки: {str(e)}"}
