from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from typing import Dict, Any, List
import tempfile
import os
import re

app = FastAPI(title="AIStoryteller Backend")

# CORS：先全开放，之后可以缩到 Netlify 域名
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 比如之后改成 ["https://aistroy.netlify.app"]
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ========== 基础工具 ==========

def shape_type_to_str(shape_type) -> str:
    """把 python-pptx 的枚举类型转成字符串。"""
    try:
        return shape_type.name
    except Exception:
        return str(shape_type)


def safe_int(value, default=0) -> int:
    try:
        return int(value)
    except Exception:
        return default


# ========== 规则：slide 类型分类 ==========

AGENDA_KEYWORDS = ["agenda", "contents", "outline", "目录", "议程"]
ENDING_KEYWORDS = ["thank you", "thanks", "q&a", "questions", "谢谢", "结束"]
SECTION_KEYWORDS = ["section", "part", "chapter", "模块", "篇", "part "]


def is_bullet_line(line: str) -> bool:
    """判断一行是不是 bullet（•, -, 数字. 等）"""
    line = line.strip()
    if not line:
        return False
    # 常见 bullet 符号
    if line.startswith(("•", "-", "·", "· ", "●", "○", "▶")):
        return True
    # 数字 or 字母编号
    if re.match(r"^(\d+|[a-zA-Z])[\.\)]\s+", line):
        return True
    return False


def classify_slide(slide_dict: Dict[str, Any], meta: Dict[str, Any]) -> str:
    """
    根据当前页的 shapes 信息 + 全局 meta（宽高）来推断 slide_type。
    返回：'title' / 'agenda' / 'section' / 'content_bullets' /
         'content_image' / 'ending' / 'other'
    """
    layout_name = (slide_dict.get("layout_name") or "").lower()
    shapes = slide_dict.get("shapes", [])
    slide_h = meta.get("slide_height_emu") or 1
    slide_w = meta.get("slide_width_emu") or 1

    text_shapes: List[Dict[str, Any]] = []
    picture_shapes: List[Dict[str, Any]] = []

    total_text_len = 0
    total_lines = 0
    bullet_lines = 0

    # 收集文本 & 图片
    for s in shapes:
        if s.get("shape_type", "").upper() in ("PICTURE", "MEDIA"):
            picture_shapes.append(s)

        if s.get("has_text_frame") and s.get("text"):
            txt = s["text"].strip()
            if not txt:
                continue
            text_shapes.append(s)
            total_text_len += len(txt)

            lines = [ln for ln in txt.splitlines() if ln.strip()]
            total_lines += len(lines)
            bullet_lines += sum(1 for ln in lines if is_bullet_line(ln))

    bullet_ratio = (bullet_lines / total_lines) if total_lines > 0 else 0.0

    # 提取「候选大标题」
    title_candidate = None
    title_score = -1.0
    for s in text_shapes:
        geom = s.get("geometry", {})
        top = geom.get("top_emu") or 0
        height = geom.get("height_emu") or 0
        text = (s.get("text") or "").strip()

        y_center_ratio = (top + height / 2) / slide_h
        area = (geom.get("width_emu") or 0) * height

        score = area
        # 上方区域加点分
        if y_center_ratio < 0.35:
            score *= 1.3

        if score > title_score and 0 < len(text) <= 60:
            title_score = score
            title_candidate = text

    # 统一小写文本，用来做关键词判断
    all_text = " ".join([(s.get("text") or "") for s in text_shapes]).lower()

    # -------- 规则开始 --------

    # 1) 如果 layout 名字里本身就有强信号
    if "title" in layout_name and "agenda" not in layout_name:
        return "title"
    if "agenda" in layout_name or "目录" in layout_name:
        return "agenda"

    # 2) 封面 / 标题页：
    #    - 第 0 页优先考虑标题
    #    - 或者 整体文字不多，只有 1~2 个块，标题候选在 1/3 附近
    if slide_dict.get("index") == 0:
        if title_candidate and len(text_shapes) <= 3:
            return "title"

    if len(text_shapes) <= 2 and total_text_len <= 80 and title_candidate:
        return "title"

    # 3) Agenda / 目录页：
    #    - 包含 agenda / contents / 目录等关键词
    #    - 有明显的 bullets
    if any(k in all_text for k in AGENDA_KEYWORDS):
        return "agenda"
    if bullet_ratio >= 0.5 and any(k in all_text for k in ["agenda", "目录", "contents"]):
        return "agenda"

    # 4) 分节页：只有一个大标题 + 几乎没有其他正文
    if len(text_shapes) <= 2 and total_text_len <= 60:
        if any(k in (title_candidate or "").lower() for k in SECTION_KEYWORDS):
            return "section"

    # 5) 结束页：感谢 / Q&A
    if any(k in all_text for k in ENDING_KEYWORDS):
        return "ending"

    # 6) 图片为主的内容页
    if picture_shapes and total_text_len <= 120:
        return "content_image"

    # 7) bullets 为主的内容页
    if bullet_ratio >= 0.4:
        return "content_bullets"

    # 8) 其他普通内容页
    if total_text_len > 0:
        return "content"

    return "other"


# ========== PPT 解析：附带 slide_type ==========

def extract_ppt_structure(prs: Presentation) -> Dict[str, Any]:
    meta = {
        "slide_count": len(prs.slides),
        "slide_width_emu": safe_int(prs.slide_width),
        "slide_height_emu": safe_int(prs.slide_height),
    }

    slides_info: List[Dict[str, Any]] = []

    for slide_idx, slide in enumerate(prs.slides):
        slide_dict: Dict[str, Any] = {
            "index": slide_idx,
            "layout_name": slide.slide_layout.name,
            "shapes": [],
        }

        for shape_idx, shape in enumerate(slide.shapes):
            geom = {
                "left_emu": safe_int(getattr(shape, "left", 0)),
                "top_emu": safe_int(getattr(shape, "top", 0)),
                "width_emu": safe_int(getattr(shape, "width", 0)),
                "height_emu": safe_int(getattr(shape, "height", 0)),
            }

            shape_dict: Dict[str, Any] = {
                "index": shape_idx,
                "name": getattr(shape, "name", None),
                "shape_type": shape_type_to_str(getattr(shape, "shape_type", None)),
                "geometry": geom,
                "has_text_frame": bool(getattr(shape, "has_text_frame", False)),
                "text": None,
            }

            if getattr(shape, "has_text_frame", False):
                text = shape.text.strip()
                shape_dict["text"] = text or None

            slide_dict["shapes"].append(shape_dict)

        # 在这里做类型分类
        slide_dict["slide_type"] = classify_slide(slide_dict, meta)
        slides_info.append(slide_dict)

    return {
        "meta": meta,
        "slides": slides_info,
    }


# ========== 路由 ==========

@app.get("/health")
def health_check():
    return {"status": "ok"}


@app.post("/api/parse_ppt")
async def parse_ppt(file: UploadFile = File(...)):
    """
    接收用户 PPTX，解析结构并给每一页打上 slide_type 标签。
    """
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
        os.remove(tmp_path)
