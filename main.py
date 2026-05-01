import pandas as pd
from pptx import Presentation
import os
import zipfile
import shutil

TEMPLATES_DIR = "templates"
OUTPUT_DIR = "output"

os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


def load_template():
    path = input("Введи путь к шаблону (.pptx): ").strip()

    if not os.path.exists(path):
        print("❌ Файл не найден")
        return

    name = os.path.basename(path)
    dest = os.path.join(TEMPLATES_DIR, name)

    shutil.copy(path, dest)
    print(f"✅ Шаблон сохранён как {name}")


def list_templates():
    files = os.listdir(TEMPLATES_DIR)

    if not files:
        print("❌ Нет шаблонов")
        return []

    print("\n📁 Мои шаблоны:")
    for i, f in enumerate(files, 1):
        print(f"{i}. {f}")

    return files


def choose_template():
    files = list_templates()
    if not files:
        return None

    choice = int(input("Выбери номер шаблона: "))
    return os.path.join(TEMPLATES_DIR, files[choice - 1])


def generate_docs(template_path):
    excel_path = input("Введи путь к Excel: ").strip()

    if not os.path.exists(excel_path):
        print("❌ Excel не найден")
        return

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
                        placeholder = str(col).strip()  # теперь берём как есть (%ФИО%)

                        if placeholder in full_text:
                            new_text = full_text.replace(placeholder, str(row[col]))

                            for run in paragraph.runs:
                                run.text = ""

                            if paragraph.runs:
                                paragraph.runs[0].text = new_text

        filename = os.path.join(OUTPUT_DIR, f"{row[df.columns[0]]}.pptx")
        prs.save(filename)
        generated_files.append(filename)

    with zipfile.ZipFile("result.zip", 'w') as zipf:
        for file in generated_files:
            zipf.write(file)

    print("\n✅ Готово! Файл result.zip создан")


def main():
    while True:
        print("\n====== МЕНЮ ======")
        print("1. Загрузить шаблон")
        print("2. Мои шаблоны")
        print("3. Сгенерировать документы")
        print("0. Выход")

        choice = input("Выбери действие: ")

        if choice == "1":
            load_template()
        elif choice == "2":
            list_templates()
        elif choice == "3":
            template = choose_template()
            if template:
                generate_docs(template)
        elif choice == "0":
            break
        else:
            print("❌ Неверный выбор")


if __name__ == "__main__":
    main()