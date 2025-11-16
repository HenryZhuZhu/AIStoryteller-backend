from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from typing import Dict, Any
import tempfile
import os

app = FastAPI(title="AIStoryteller Backend")

# 先全开放 CORS，之后你可以改成只允许 Netlify 的域名
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 比如部署后改成 ["https://aistroy.netlify.app"]
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def shape_type_to_str(shape_type) -> str:
    """把 python-pptx 的枚举类型转成字符串，避免前端难解析。"""
    try:
        return shape_type.name
    except Exception:
        return str(shape_type)


def extract_ppt_structure(prs: Presentation) -> Dict[str, Any]:
    """
    把任意用户 PPT 转成统一结构的 JSON，给前端 HTML5 生成器使用。
    这里故意做得“扁平简单”，便于前端理解：
      - meta: 全局信息（页数、宽高）
      - slides: 每一页的 shapes 列表
    """
    meta = {
        "slide_count": len(prs.slides),
        "slide_width_emu": int(prs.slide_width),
        "slide_height_emu": int(prs.slide_height),
    }

    slides_info = []

    for slide_idx, slide in enumerate(prs.slides):
        slide_dict = {
            "index": slide_idx,
            "layout_name": slide.slide_layout.name,
            "shapes": [],
        }

        for shape_idx, shape in enumerate(slide.shapes):
            shape_dict = {
                "index": shape_idx,
                "name": getattr(shape, "name", None),
                "shape_type": shape_type_to_str(getattr(shape, "shape_type", None)),
                "geometry": {
                    "left_emu": int(getattr(shape, "left", 0)),
                    "top_emu": int(getattr(shape, "top", 0)),
                    "width_emu": int(getattr(shape, "width", 0)),
                    "height_emu": int(getattr(shape, "height", 0)),
                },
                "has_text_frame": bool(getattr(shape, "has_text_frame", False)),
                "text": None,
            }

            # 如果有文本框，就把文字提出来
            if getattr(shape, "has_text_frame", False):
                text = shape.text.strip()
                shape_dict["text"] = text or None

            slide_dict["shapes"].append(shape_dict)

        slides_info.append(slide_dict)

    return {
        "meta": meta,
        "slides": slides_info,
    }


@app.get("/health")
def health_check():
    """简单健康检查接口，方便你测试服务是否正常启动"""
    return {"status": "ok"}


@app.post("/api/parse_ppt")
async def parse_ppt(file: UploadFile = File(...)):
    """
    接收用户上传的 PPTX，使用 python-pptx 解析结构，
    返回前端 HTML5 生成器可用的 JSON。
    """
    # 临时保存上传文件
    suffix = ".pptx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        contents = await file.read()
        tmp.write(contents)
        tmp_path = tmp.name

    try:
        prs = Presentation(tmp_path)
        data = extract_ppt_structure(prs)
        return data
    finally:
        # 删除临时文件，防止磁盘堆积
        os.remove(tmp_path)
