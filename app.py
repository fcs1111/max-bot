from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
import requests
import os

app = FastAPI()

# папка
TEMPLATES_DIR = "templates"
os.makedirs(TEMPLATES_DIR, exist_ok=True)


# тест
@app.get("/", response_class=PlainTextResponse)
def test():
    return "бот работает"


# загрузка шаблона
@app.post("/upload_template")
async def upload_template(request: Request):
    data = await request.json()

    print("INCOMING:", data)

    file_url = None

    # 👇 ищем файл в variables
    variables = data.get("variables", [])

    for var in variables:
        if var.get("name") == "file":
    payload = var.get("payload", {})
    file_url = payload.get("url")

    if not file_url:
        return {"message": f"Файл не найден. Пришло: {data}"}

    try:
        response = requests.get(file_url)
        response.raise_for_status()

        filename = file_url.split("/")[-1]
        path = os.path.join("templates", filename)

        with open(path, "wb") as f:
            f.write(response.content)

        return {"message": f"Шаблон {filename} загружен"}

    except Exception as e:
        return {"message": f"Ошибка: {str(e)}"}
