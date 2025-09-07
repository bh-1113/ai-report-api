from fastapi import FastAPI
from fastapi.responses import FileResponse
import os
from pptx import Presentation

app = FastAPI()

@app.get("/make_ppt")
def make_ppt(topic: str):
    prs = Presentation()
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = topic
    subtitle.text = f"{topic}에 대한 자동 생성 보고서"

    filename = "report.pptx"
    prs.save(filename)
    return FileResponse(filename, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename=filename)

# Render에서 포트 환경 변수 사용
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))  # Render 환경에 맞춤
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=True)
