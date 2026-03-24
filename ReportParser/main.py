import json
import tempfile
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel
from typing import Optional, List

from docxParser import DocumentParser   # ← убедитесь, что импорт правильный

# Создаём папку для результатов
Path("Results").mkdir(exist_ok=True)

app = FastAPI(title="Проверка оформления DOCX")


class CheckResponse(BaseModel):
    comment_count: int
    filename: str
    download_url: str


class BatchCheckItem(BaseModel):
    original_filename: str
    comment_count: int
    checked_filename: str
    download_url: str


class BatchCheckResponse(BaseModel):
    results: List[BatchCheckItem]
    total_documents: int
    total_comments: int


# ====================== ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ======================
def generate_output_path(original_filename: str) -> Path:
    """Генерирует имя выходного файла по логике оригинального алгоритма"""
    basename = Path(original_filename).stem  # без расширения
    out_path = Path("Results") / f"{basename}_Проверенный.docx"

    count = 1
    while out_path.exists():
        out_path = Path("Results") / f"{basename}_Проверенный_{count}.docx"
        count += 1

    return out_path


# ====================== ОДИНОЧНАЯ ПРОВЕРКА ======================
@app.post("/check-docx", response_model=CheckResponse)
async def check_docx(
    file: UploadFile = File(..., description="Файл .docx"),
    criteria: Optional[str] = Form(None, description="JSON со словарями настроек")
):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(400, "Только .docx файлы")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(await file.read())
        input_path = tmp.name

    try:
        parser = DocumentParser()

        if criteria:
            try:
                crit_dict = json.loads(criteria)
                parser.set_settings(**crit_dict)
            except Exception as e:
                raise HTTPException(400, f"Неверный JSON критериев: {e}")

        comment_count, _ = parser.parse_document(input_path)   # игнорируем старый output_path

        # Генерируем правильное имя по оригинальной логике
        output_path = generate_output_path(file.filename)
        # Переименовываем файл, который создал parser
        Path(_).rename(output_path)   # _ — это старый путь от parser.parse_document

        checked_filename = output_path.name
        download_url = f"/download/{checked_filename}"

        return CheckResponse(
            comment_count=comment_count,
            filename=checked_filename,
            download_url=download_url
        )

    finally:
        Path(input_path).unlink(missing_ok=True)


# ====================== МАССОВАЯ ПРОВЕРКА ======================
@app.post("/check-docx-batch", response_model=BatchCheckResponse)
async def check_docx_batch(
    files: List[UploadFile] = File(..., description="Множество файлов .docx"),
    criteria: Optional[str] = Form(None, description="JSON с критериями (один на все файлы)"),
    max_files: int = Query(20, description="Максимальное количество файлов")
):
    if len(files) > max_files:
        raise HTTPException(400, f"Максимум {max_files} файлов за один запрос")

    results = []
    total_comments = 0

    for uploaded_file in files:
        if not uploaded_file.filename.lower().endswith(".docx"):
            continue

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(await uploaded_file.read())
            input_path = tmp.name

        try:
            parser = DocumentParser()

            if criteria:
                try:
                    crit_dict = json.loads(criteria)
                    parser.set_settings(**crit_dict)
                except Exception as e:
                    raise HTTPException(400, f"Неверный JSON критериев: {e}")

            comment_count, old_output_path = parser.parse_document(input_path)

            # Генерируем красивое имя по оригинальной логике
            output_path = generate_output_path(uploaded_file.filename)

            # Перемещаем/переименовываем файл, созданный парсером
            Path(old_output_path).rename(output_path)

            checked_filename = output_path.name
            download_url = f"/download/{checked_filename}"

            results.append(BatchCheckItem(
                original_filename=uploaded_file.filename,
                comment_count=comment_count,
                checked_filename=checked_filename,
                download_url=download_url
            ))
            total_comments += comment_count

        finally:
            Path(input_path).unlink(missing_ok=True)

    return BatchCheckResponse(
        results=results,
        total_documents=len(results),
        total_comments=total_comments
    )


# ====================== СКАЧИВАНИЕ ======================
@app.get("/download/{filename}")
async def download_file(filename: str):
    file_path = Path("Results") / filename
    if not file_path.exists():
        raise HTTPException(404, "Файл не найден")
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


# ====================== ДЕМО СТРАНИЦА ======================
@app.get("/demo", response_class=HTMLResponse)
async def get_demo_page():
    with open("demo_checker.html", encoding="utf-8") as f:
        return f.read()