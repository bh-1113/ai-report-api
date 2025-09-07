from fastapi import FastAPI, Query
from fastapi.responses import FileResponse
from pptx import Presentation

app = FastAPI()

@app.get("/")
def home():
    return {"message": "AI 보고서 생성 API 입니다. /make_ppt?topic=주제 형태로 호출하세요."}

@app.get("/make_ppt")
def make_ppt(topic: str = Query(..., description="보고서 주제")):
    prs = Presentation()

    # 제목 슬라이드
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = f"{topic} 보고서"
    subtitle.text = "자동 생성된 PowerPoint 파일"

    # 예시 본문 슬라이드
    contents = [
        f"{topic} 개요",
        f"{topic} 활용 사례",
        f"{topic} 장점과 한계",
        f"{topic} 미래 전망"
    ]

    for content in contents:
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = content
        slide.placeholders[1].text = f"{content}에 대한 설명이 자동 생성됩니다."

    file_name = "report.pptx"
    prs.save(file_name)

    return FileResponse(
        file_name,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=file_name
    )
