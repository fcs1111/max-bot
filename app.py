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
from fastapi import FastAPI, Request
import requests
import os

app = FastAPI()

os.makedirs("templates", exist_ok=True)

@app.post("/upload_template")
async def upload_template(request: Request):
    data = await request.json()

    file_url = data.get("contact", {}).get("last_file_url")

    if not file_url:
        return {"message": "Файл не найден"}

    response = requests.get(file_url)

    filename = file_url.split("/")[-1]
    path = os.path.join("templates", filename)

    with open(path, "wb") as f:
        f.write(response.content)

    return {"message": f"Шаблон {filename} загружен"}
