from fastapi import FastAPI
from fastapi.responses import FileResponse
from pptx import Presentation
from pptx.util import Inches
import tempfile, os, requests
from openai import OpenAI

app = FastAPI()

# Render 환경변수에서 API Key 불러오기
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# 슬라이드 항목
sections = ["개요", "필요성", "활용 사례", "장점과 한계", "미래 전망"]

# GPT를 이용해 본문 텍스트 생성
def generate_text(topic, section):
    prompt = f"{topic}에 대해 '{section}' 파트의 발표 슬라이드 내용을 글머리표 3~4개로 작성해줘."
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

# DALL·E를 이용해 이미지 생성
def generate_image(topic, section):
    prompt = f"{topic} - {section} 발표 슬라이드용 삽화"
    response = client.images.generate(
        model="gpt-image-1",
        prompt=prompt,
        size="512x512"
    )
    return response.data[0].url

@app.get("/make_ppt")
def make_ppt(topic: str):
    prs = Presentation()

    # 표지
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = f"{topic} 보고서"
    slide.placeholders[1].text = "자동 생성된 AI 보고서"

    # 본문
    for section in sections:
        text = generate_text(topic, section)
        image_url = generate_image(topic, section)

        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = section
        slide.placeholders[1].text = text

        # 이미지 삽입
        img_data = requests.get(image_url).content
        tmp_img = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        tmp_img.write(img_data)
        tmp_img.close()
        slide.shapes.add_picture(tmp_img.name, Inches(5), Inches(2), Inches(3), Inches(3))

    # 파일 반환
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp.name)
    return FileResponse(
        tmp.name,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=f"{topic}_보고서.pptx"
    )
