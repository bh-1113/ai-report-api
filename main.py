from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
import shutil

# 서버에 임시 ppt 파일 저장용 폴더
if not os.path.exists("static"):
    os.makedirs("static")

app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/make_ppt")
def make_ppt(topic: str):
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = f"{topic} 보고서"
    slide.placeholders[1].text = "자동 생성된 AI 보고서"

    report_text = ""
    for section in sections:
        text = generate_text(topic, section)
        report_text += f"{section}:\n{text}\n\n"

        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = section
        slide.placeholders[1].text = text

    tmp_path = f"static/{topic}_보고서.pptx"
    prs.save(tmp_path)

    # JSON으로 텍스트와 다운로드 URL 반환
    return JSONResponse({
        "report_text": report_text,
        "ppt_url": f"/static/{topic}_보고서.pptx"
    })
