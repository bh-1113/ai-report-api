from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
from PyPDF2 import PdfReader
from docx import Document
from pptx import Presentation
import pandas as pd
import tempfile, os
from openai import OpenAI

app = FastAPI()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# 업로드된 요약 저장 (메모리 캐싱)
last_summary = ""

def extract_text(file: UploadFile):
    ext = file.filename.split(".")[-1].lower()
    text = ""

    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp:
        tmp.write(file.file.read())
        tmp_path = tmp.name

    if ext == "pdf":
        reader = PdfReader(tmp_path)
        for page in reader.pages:
            text += page.extract_text() + "\n"
    elif ext == "docx":
        doc = Document(tmp_path)
        for p in doc.paragraphs:
            text += p.text + "\n"
    elif ext == "pptx":
        prs = Presentation(tmp_path)
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text += shape.text + "\n"
    elif ext in ["xlsx", "xls"]:
        df = pd.read_excel(tmp_path)
        text = df.to_string()
    elif ext == "txt":
        with open(tmp_path, "r", encoding="utf-8") as f:
            text = f.read()

    return text.strip()

@app.post("/upload_summary")
async def upload_summary(file: UploadFile = File(...)):
    global last_summary
    text = extract_text(file)

    # GPT 요약
    prompt = f"다음 문서를 한국어로 30% 분량으로 요약해줘:\n\n{text[:4000]}"
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}]
    )
    last_summary = response.choices[0].message.content

    return JSONResponse(content={"summary": last_summary})

@app.get("/download_summary")
def download_summary(format: str):
    global last_summary
    if not last_summary:
        return JSONResponse(content={"error": "No summary available"}, status_code=400)

    if format == "docx":
        doc = Document()
        doc.add_heading("문서 요약", level=1)
        doc.add_paragraph(last_summary)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(tmp.name)
        return FileResponse(tmp.name, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="summary.docx")

    elif format == "pptx":
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "문서 요약"
        slide.placeholders[1].text = last_summary
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        prs.save(tmp.name)
        return FileResponse(tmp.name, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename="summary.pptx")

    return JSONResponse(content={"error": "Invalid format"}, status_code=400)
