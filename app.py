from __future__ import annotations

import argparse
import base64
import json
import locale
import os
import re
import subprocess
import traceback
import unicodedata
from copy import deepcopy
from pathlib import Path
from typing import Any

from openai import OpenAI
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "template"
PROMPTS_DIR = BASE_DIR / "prompts"
OUTPUT_DIR = BASE_DIR / "output"
IMAGES_DIR = OUTPUT_DIR / "images"

DEFAULT_PROMPT_FILE = PROMPTS_DIR / "lecture_prompt.txt"
DEFAULT_CURRICULUM_FILE = PROMPTS_DIR / "curriculum.txt"
DEFAULT_ENV_FILE = BASE_DIR / ".env"


JSON_FORMAT_GUIDE = """
Return JSON only.

{
  "deck_title": "string",
  "section_label": "string",
  "slides": [
    {
      "slide_number": 1,
      "slide_role": "objective | content | practice_problem | practice_answer | summary",
      "source_slide": 21,
      "title": "string",
      "why": "string",
      "bullets": ["string", "string"],
      "example": "string",
      "practice_prompt": "string",
      "transition": "string",
      "image_prompt": "string",
      "table": {
        "headers": ["string", "string"],
        "rows": [["string", "string"]]
      }
    }
  ]
}

Rules:
- Keep exactly 20 slides unless the prompt explicitly requests another count.
- Slide 1 must use slide_role "objective".
- Slide 18 must use slide_role "practice_problem".
- Slide 19 must use slide_role "practice_answer".
- Slide 20 must use slide_role "summary".
- Use the provided content template slide as the main source_slide unless there is a strong reason not to.
- The template structure is simple: left-top bullet text area and right-bottom image area.
- Do not use tables unless the prompt strongly requires a comparison table.
- Every bullet must be a presentation-ready sentence of about 18 to 35 Korean characters when possible.
- Avoid essay-like long sentences.
- Avoid keyword-only short fragments.
- Prefer one idea per bullet and keep the wording easy to say aloud in class.
- Prefer concise nominal endings such as "~임", "~함", "~음", "~필요", "~중심".
- Avoid sentence endings like "~다", "~입니다", "~같다" in slide bullets.
- Put the image prompt in "image_prompt" when a visual aid is useful.
""".strip()


def load_env_file(path: Path) -> None:
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip())


def env_flag(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a lecture deck from an LLM prompt file and a PPT template."
    )
    parser.add_argument(
        "--prompt-file",
        default=str(DEFAULT_PROMPT_FILE),
        help="Prompt text file path. Default: ./prompts/lecture_prompt.txt",
    )
    parser.add_argument(
        "--template",
        help="Template .pptx path. Default: first .pptx in ./template",
    )
    parser.add_argument(
        "--curriculum-file",
        help="Curriculum text file path for multi-session generation. Default: ./prompts/curriculum.txt when present",
    )
    parser.add_argument(
        "--output",
        help="Output .pptx path. Default: ./output/generated_<template>.pptx",
    )
    parser.add_argument(
        "--json-output",
        help="Where to save the structured slide JSON. Default: ./output/generated_slide_plan.json",
    )
    parser.add_argument(
        "--analysis-output",
        help="Where to save the template analysis JSON. Default: ./output/template_analysis.json",
    )
    parser.add_argument(
        "--model",
        default=os.getenv("OPENAI_MODEL", "gpt-4.1"),
        help="LLM model name. Default: OPENAI_MODEL env or gpt-4.1",
    )
    parser.add_argument(
        "--mock",
        action="store_true",
        help="Skip API call and use built-in sample slide data.",
    )
    parser.add_argument(
        "--analyze-only",
        action="store_true",
        help="Only analyze the template and save template_analysis.json.",
    )
    parser.add_argument(
        "--skip-images",
        action="store_true",
        help="Skip image generation and keep the template image as-is. Overrides .env setting.",
    )
    return parser.parse_args()


def ensure_dirs() -> None:
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    PROMPTS_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    IMAGES_DIR.mkdir(parents=True, exist_ok=True)


def find_default_template() -> Path:
    candidates = [
        path for path in TEMPLATE_DIR.glob("*.pptx") if not path.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(
            f"No .pptx template found in {TEMPLATE_DIR}. Put your template PPT there."
        )
    if len(candidates) == 1:
        return candidates[0]

    preferred = [path for path in candidates if path.stem != "templates"]
    pool = preferred or candidates
    return max(pool, key=lambda path: path.stat().st_mtime)


def build_output_path(template_path: Path) -> Path:
    return OUTPUT_DIR / f"generated_{template_path.stem}.pptx"


def build_json_output_path() -> Path:
    return OUTPUT_DIR / "generated_slide_plan.json"


def build_analysis_output_path() -> Path:
    return OUTPUT_DIR / "template_analysis.json"


def build_raw_output_path() -> Path:
    return OUTPUT_DIR / "llm_raw_response.txt"


def build_image_output_dir() -> Path:
    IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    return IMAGES_DIR


def build_merged_output_path() -> Path:
    return OUTPUT_DIR / "final_merged_curriculum.pptx"


def build_error_log_path() -> Path:
    return OUTPUT_DIR / "error.log"


def append_error_log(message: str) -> None:
    path = build_error_log_path()
    with path.open("a", encoding="utf-8") as fp:
        fp.write(message.rstrip() + "\n")


def write_session_error_file(session_prefix: str, text: str) -> Path:
    path = OUTPUT_DIR / f"{session_prefix}_error.txt"
    path.write_text(text, encoding="utf-8")
    return path


def write_merge_error_file(text: str) -> Path:
    path = OUTPUT_DIR / "merge_error.log"
    path.write_text(text, encoding="utf-8")
    return path


def read_prompt_file(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"Prompt file not found: {path}")
    return path.read_text(encoding="utf-8")


def slugify_korean(text: str, max_len: int = 40) -> str:
    normalized = unicodedata.normalize("NFKC", text)
    normalized = re.sub(r"[^\w가-힣]+", "_", normalized, flags=re.UNICODE)
    normalized = re.sub(r"_+", "_", normalized).strip("_")
    return normalized[:max_len] or "session"


def parse_curriculum_file(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []

    lines = [line.strip() for line in path.read_text(encoding="utf-8").splitlines()]
    header_pattern = re.compile(r".*?(\d+)교시\.\s*(.+)")

    sessions: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    pending_label: str | None = None

    for line in lines:
        if not line:
            continue

        header_match = header_pattern.match(line)
        if header_match:
            if current:
                sessions.append(current)
            current = {
                "session_no": int(header_match.group(1)),
                "title": header_match.group(2).strip(),
                "core_points": [],
                "meta": {},
            }
            pending_label = None
            continue

        if current is None:
            continue

        if line.startswith("🟩 [") or line.startswith("🟦 ["):
            continue

        if line.startswith("👉"):
            pending_label = line.replace("👉", "").strip() or None
            continue

        if pending_label:
            current["meta"].setdefault(pending_label, []).append(line)
            pending_label = None
            continue

        current["core_points"].append(line)

    if current:
        sessions.append(current)

    return sessions


def render_session_prompt(prompt_template: str, session: dict[str, Any]) -> str:
    core_lines = "\n".join(f"- {item}" for item in session["core_points"])
    meta_lines: list[str] = []
    for label, values in session.get("meta", {}).items():
        for value in values:
            meta_lines.append(f"- {label}: {value}")

    rendered = prompt_template.replace(
        "주제: 생성형 AI 개요 및 업무 변화 이해",
        f"주제: {session['title']}",
    )

    default_core_block = (
        "[핵심 내용]\n"
        "- 생성형 AI란 무엇인가 (LLM 개념)\n"
        "- 기존 자동화 vs 생성형 AI 차이\n"
        "- 기업에서의 활용 사례 (보고서, 분석, 요약)\n"
        "- 생성형 AI의 한계 (환각, 통제 불가)"
    )
    replacement_core_block = "[핵심 내용]\n" + core_lines
    if meta_lines:
        replacement_core_block += "\n[추가 맥락]\n" + "\n".join(meta_lines)

    rendered = rendered.replace(default_core_block, replacement_core_block)
    rendered += (
        f"\n\n[교시 정보]\n- 교시: {session['session_no']}교시\n- 이번 교시 주제: {session['title']}\n"
    )
    return rendered


def build_image_output_dir_for_session(session_prefix: str | None = None) -> Path:
    if not session_prefix:
        return build_image_output_dir()
    path = build_image_output_dir() / session_prefix
    path.mkdir(parents=True, exist_ok=True)
    return path


def remove_all_shapes(slide) -> None:
    for shape in list(slide.shapes):
        shape.element.getparent().remove(shape.element)


def clone_slide(prs: Presentation, source_index: int):
    source_slide = prs.slides[source_index]
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    remove_all_shapes(new_slide)

    for shape in source_slide.shapes:
        new_slide.shapes._spTree.insert_element_before(
            deepcopy(shape.element), "p:extLst"
        )

    for rel in source_slide.part.rels.values():
        if any(
            token in rel.reltype
            for token in ("notesSlide", "slideLayout", "slideMaster", "theme")
        ):
            continue
        new_slide.part.rels._add_relationship(rel.reltype, rel._target, rel.is_external)

    return new_slide


def clone_slide_from_source(source_slide, target_prs: Presentation):
    raise NotImplementedError("Use clone_slide() within the template presentation instead.")


def remove_slides_by_indices(prs: Presentation, indices: list[int]) -> None:
    slide_id_list = prs.slides._sldIdLst
    for index in sorted(indices, reverse=True):
        slide_id = slide_id_list[index]
        rel_id = slide_id.rId
        prs.part.drop_rel(rel_id)
        slide_id_list.remove(slide_id)


def get_picture_shapes(slide) -> list[Any]:
    return [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]


def get_main_picture_spec(slide) -> dict[str, Any] | None:
    pictures = get_picture_shapes(slide)
    if not pictures:
        return None

    target = max(pictures, key=lambda shape: shape.width * shape.height)
    return {
        "left": target.left,
        "top": target.top,
        "width": target.width,
        "height": target.height,
    }


def remove_picture_shapes(slide) -> None:
    for shape in list(slide.shapes):
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        shape.element.getparent().remove(shape.element)


def set_text(shape, text: str) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    text_frame = shape.text_frame
    text_frame.clear()
    text_frame.paragraphs[0].text = text


def set_paragraphs(shape, lines: list[str]) -> None:
    if not getattr(shape, "has_text_frame", False):
        return

    cleaned = [line.strip() for line in lines if line and line.strip()]
    text_frame = shape.text_frame
    text_frame.clear()

    if not cleaned:
        text_frame.paragraphs[0].text = ""
        return

    for idx, line in enumerate(cleaned):
        paragraph = text_frame.paragraphs[0] if idx == 0 else text_frame.add_paragraph()
        paragraph.text = line
        paragraph.level = 0


def apply_font_style(shape, font_name: str, font_size: int, bold: bool | None = None) -> None:
    if not getattr(shape, "has_text_frame", False):
        return

    for paragraph in shape.text_frame.paragraphs:
        paragraph.line_spacing = 1.5
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            if bold is not None:
                run.font.bold = bold


def apply_default_text_style(shape, font_name: str, font_size: int = 20) -> None:
    apply_font_style(shape, font_name=font_name, font_size=font_size, bold=None)


def apply_title_style(shape, font_name: str) -> None:
    apply_font_style(shape, font_name=font_name, font_size=20, bold=True)


def iter_text_shapes(slide) -> list[Any]:
    return [shape for shape in slide.shapes if getattr(shape, "has_text_frame", False)]


def table_shape_count(slide) -> int:
    return sum(1 for shape in slide.shapes if getattr(shape, "has_table", False))


def picture_shape_count(slide) -> int:
    return sum(1 for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE)


def shape_text_preview(shape) -> str:
    if not getattr(shape, "has_text_frame", False):
        return ""
    parts = [p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip()]
    return " | ".join(parts)[:120]


def suggest_tags(slide_number: int, text_shapes: list[Any], pictures: int, tables: int) -> list[str]:
    tags: list[str] = []

    if slide_number == 1:
        tags.append("cover")
    if slide_number == 2:
        tags.append("agenda")
    if slide_number in {3, 26}:
        tags.append("section_or_closing")
    if tables:
        tags.append("table")
    if pictures:
        tags.append("image")
    if len(text_shapes) >= 5:
        tags.append("multi_text")
    if len(text_shapes) >= 3 and not tables:
        tags.append("bullet")
    if len(text_shapes) >= 4 and not tables:
        tags.append("practice_like")
    return tags


def analyze_template(template_path: Path) -> dict[str, Any]:
    prs = Presentation(str(template_path))
    slides_summary: list[dict[str, Any]] = []
    candidate_templates: list[dict[str, Any]] = []
    content_template_slide = 2 if len(prs.slides) >= 2 else 1

    for slide_number, slide in enumerate(prs.slides, start=1):
        text_shapes = iter_text_shapes(slide)
        text_previews = [shape_text_preview(shape) for shape in text_shapes if shape_text_preview(shape)]
        pictures = picture_shape_count(slide)
        tables = table_shape_count(slide)
        tags = suggest_tags(slide_number, text_shapes, pictures, tables)

        slide_info = {
            "slide_number": slide_number,
            "text_shape_count": len(text_shapes),
            "picture_count": pictures,
            "table_count": tables,
            "tags": tags,
            "text_previews": text_previews[:5],
        }
        slides_summary.append(slide_info)

        if slide_number == content_template_slide:
            candidate_templates.append(
                {
                    "source_slide": slide_number,
                    "text_shape_count": len(text_shapes),
                    "picture_count": pictures,
                    "table_count": tables,
                    "tags": tags,
                    "use_when": "Use this as the main content slide template with left-top bullets and right-bottom image.",
                    "text_previews": text_previews[:4],
                }
            )

    return {
        "template_file": str(template_path),
        "slide_count": len(prs.slides),
        "slide_width": prs.slide_width,
        "slide_height": prs.slide_height,
        "content_template_slide": content_template_slide,
        "candidate_templates": candidate_templates,
        "slides": slides_summary,
    }


def fill_table(shape, table_data: dict[str, Any]) -> None:
    if not getattr(shape, "has_table", False):
        return

    headers = table_data.get("headers") or []
    rows = table_data.get("rows") or []
    table = shape.table

    for col_idx, header in enumerate(headers[: len(table.columns)]):
        table.cell(0, col_idx).text = str(header)

    for row_idx, row in enumerate(rows[: len(table.rows) - 1], start=1):
        for col_idx, cell in enumerate(row[: len(table.columns)]):
            table.cell(row_idx, col_idx).text = str(cell)


def get_first_table_spec(slide) -> dict[str, Any] | None:
    for shape in slide.shapes:
        if getattr(shape, "has_table", False):
            return {
                "rows": len(shape.table.rows),
                "cols": len(shape.table.columns),
                "left": shape.left,
                "top": shape.top,
                "width": shape.width,
                "height": shape.height,
            }
    return None


def add_blank_table_from_source(target_slide, source_slide, table_data: dict[str, Any]) -> None:
    table_spec = get_first_table_spec(source_slide)
    if not table_spec:
        return

    rows = max(len(table_data.get("rows", [])) + 1, table_spec["rows"])
    cols = max(len(table_data.get("headers", [])), table_spec["cols"])
    table_shape = target_slide.shapes.add_table(
        rows,
        cols,
        table_spec["left"],
        table_spec["top"],
        table_spec["width"],
        table_spec["height"],
    )
    fill_table(table_shape, table_data)


def normalize_bullet_line(text: str) -> str:
    cleaned = (text or "").strip()
    if not cleaned:
        return ""

    cleaned = re.sub(r"^[\u2022\u25E6\u25AA\u25CF\-\*\·\▪\‣]+\s*", "", cleaned)
    cleaned = re.sub(r"^\d+[\.\)]\s*", "", cleaned)
    cleaned = re.sub(r"^[가-힣A-Za-z][\.\)]\s*", "", cleaned)
    cleaned = re.sub(r"^[\u2022\u25E6\u25AA\u25CF\-\*\·\▪\‣]+\s*", "", cleaned)
    return cleaned.strip()


def format_main_body(slide_data: dict[str, Any]) -> list[str]:
    bullets = slide_data.get("bullets") or []
    normalized: list[str] = []
    for bullet in bullets:
        line = normalize_bullet_line(str(bullet))
        if line:
            normalized.append(line)
    return normalized


def format_example_block(slide_data: dict[str, Any]) -> list[str]:
    lines: list[str] = []

    if slide_data.get("example"):
        lines.append("● 구체적 예시 또는 수치")
        lines.append(slide_data["example"])

    if slide_data.get("practice_prompt"):
        lines.append("● AI 프롬프트 예시")
        lines.append(slide_data["practice_prompt"])

    if slide_data.get("image_prompt"):
        lines.append("● 이미지 프롬프트")
        lines.append(slide_data["image_prompt"])

    return lines


def build_presenter_notes(slide_data: dict[str, Any]) -> list[str]:
    lines: list[str] = []

    if slide_data.get("why"):
        lines.append("[왜 이 내용이 필요한가]")
        lines.append(slide_data["why"])

    if slide_data.get("example"):
        lines.append("[구체적 예시 또는 수치]")
        lines.append(slide_data["example"])

    if slide_data.get("practice_prompt"):
        lines.append("[AI 프롬프트 예시]")
        lines.append(slide_data["practice_prompt"])

    if slide_data.get("image_prompt"):
        lines.append("[이미지 프롬프트]")
        lines.append(slide_data["image_prompt"])

    if slide_data.get("transition"):
        lines.append("[다음 단계 연결]")
        lines.append(slide_data["transition"])

    return lines


def set_presenter_notes(slide, slide_data: dict[str, Any], font_name: str) -> None:
    notes_slide = slide.notes_slide
    note_lines = build_presenter_notes(slide_data)
    if not note_lines:
        return

    for shape in notes_slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        if getattr(shape, "is_placeholder", False):
            placeholder_type = shape.placeholder_format.type
            if str(placeholder_type) != "BODY (2)":
                continue
            set_paragraphs(shape, note_lines)
            apply_default_text_style(shape, font_name=font_name, font_size=12)
            return


def fill_slide(slide, source_slide, slide_data: dict[str, Any], section_label: str) -> None:
    font_name = "Malgun Gothic"
    slide_role = slide_data["slide_role"]
    text_shapes = iter_text_shapes(slide)

    if len(text_shapes) < 3:
        raise ValueError(
            f"Source slide {slide_data['source_slide']} does not have enough text slots."
        )

    set_text(text_shapes[0], section_label)
    set_text(text_shapes[1], slide_data["title"])
    apply_title_style(text_shapes[0], font_name=font_name)
    apply_title_style(text_shapes[1], font_name=font_name)

    main_lines = format_main_body(slide_data)
    if slide_role in {"objective", "content", "summary"}:
        set_paragraphs(text_shapes[2], main_lines)
    elif slide_role == "practice_problem":
        set_paragraphs(text_shapes[2], main_lines)
    elif slide_role == "practice_answer":
        set_paragraphs(text_shapes[2], main_lines)
    else:
        set_paragraphs(text_shapes[2], main_lines)

    apply_default_text_style(text_shapes[2], font_name=font_name, font_size=16)
    set_presenter_notes(slide, slide_data, font_name=font_name)

    if slide_data.get("table"):
        add_blank_table_from_source(slide, source_slide, slide_data["table"])


def extract_json_block(text: str) -> dict[str, Any]:
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    fenced = re.search(r"```json\s*(\{.*\})\s*```", text, re.DOTALL)
    if fenced:
        return json.loads(fenced.group(1))

    bare = re.search(r"(\{.*\})", text, re.DOTALL)
    if bare:
        return json.loads(bare.group(1))

    raise ValueError("LLM response did not contain valid JSON.")


def build_llm_prompt(prompt_text: str, template_analysis: dict[str, Any]) -> str:
    analysis_payload = {
        "slide_count": template_analysis["slide_count"],
        "content_template_slide": template_analysis["content_template_slide"],
        "candidate_templates": template_analysis["candidate_templates"],
    }
    return (
        f"{prompt_text}\n\n"
        "[템플릿 분석 결과]\n"
        "아래 JSON은 템플릿 PPT를 로컬에서 분석한 결과다. "
        "슬라이드 구조를 설계할 때 반드시 이 후보 장표들 중에서 source_slide를 선택하라.\n\n"
        f"{json.dumps(analysis_payload, ensure_ascii=False, indent=2)}\n\n"
        f"{JSON_FORMAT_GUIDE}"
    )


def get_openai_client() -> OpenAI:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise EnvironmentError("OPENAI_API_KEY is not set.")

    return OpenAI(
        api_key=api_key,
        base_url=os.getenv("OPENAI_BASE_URL") or None,
        timeout=float(os.getenv("OPENAI_TIMEOUT_SECONDS", "300")),
    )


def generate_slide_plan_with_llm(
    prompt_text: str, template_analysis: dict[str, Any], model: str
) -> tuple[dict[str, Any], str]:
    client = get_openai_client()

    response = client.responses.create(
        model=model,
        input=[
            {
                "role": "system",
                "content": [
                    {
                        "type": "input_text",
                        "text": "You are a lecture deck planner. Output JSON only.",
                    }
                ],
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": build_llm_prompt(prompt_text, template_analysis),
                    }
                ],
            },
        ],
    )

    raw_text = response.output_text
    return extract_json_block(raw_text), raw_text


def generate_slide_image(
    client: OpenAI,
    image_prompt: str,
    output_dir: Path,
    slide_number: int,
) -> Path:
    image_model = os.getenv("OPENAI_IMAGE_MODEL", "gpt-image-1")
    result = client.images.generate(
        model=image_model,
        prompt=image_prompt,
        size="1024x1024",
    )

    image_path = output_dir / f"slide_{slide_number:02d}.png"
    image_base64 = None
    if getattr(result, "data", None):
        item = result.data[0]
        image_base64 = getattr(item, "b64_json", None)

    if not image_base64:
        raise ValueError("Image generation response did not include b64_json data.")

    image_path.write_bytes(base64.b64decode(image_base64))
    return image_path


def get_cached_image_path(output_dir: Path, slide_number: int) -> Path:
    return output_dir / f"slide_{slide_number:02d}.png"


def replace_slide_image(slide, source_slide, image_path: Path) -> None:
    picture_spec = get_main_picture_spec(source_slide)
    if not picture_spec:
        return

    remove_picture_shapes(slide)
    slide.shapes.add_picture(
        str(image_path),
        picture_spec["left"],
        picture_spec["top"],
        picture_spec["width"],
        picture_spec["height"],
    )


def clone_rendered_slide_to_prs(source_slide, target_prs: Presentation):
    blank_layout = target_prs.slide_layouts[6]
    new_slide = target_prs.slides.add_slide(blank_layout)

    for shape in source_slide.shapes:
        new_slide.shapes._spTree.insert_element_before(
            deepcopy(shape.element), "p:extLst"
        )

    for rel in source_slide.part.rels.values():
        if any(
            token in rel.reltype
            for token in ("notesSlide", "slideLayout", "slideMaster", "theme")
        ):
            continue
        new_slide.part.rels._add_relationship(rel.reltype, rel._target, rel.is_external)

    return new_slide


def decode_subprocess_output(data: bytes | None) -> str:
    if not data:
        return ""

    encodings = ["utf-8", "cp949"]
    preferred = locale.getpreferredencoding(False)
    if preferred and preferred not in encodings:
        encodings.append(preferred)

    for encoding in encodings:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue

    return data.decode("utf-8", errors="replace")


def is_valid_presentation(path: Path) -> bool:
    if not path.exists():
        return False
    try:
        prs = Presentation(str(path))
        return len(prs.slides) > 0
    except Exception:
        return False


def merge_presentations(ppt_paths: list[Path], output_path: Path) -> None:
    if not ppt_paths:
        return
    merge_script = BASE_DIR / "merge_ppt.ps1"
    if not merge_script.exists():
        raise FileNotFoundError(f"Merge script not found: {merge_script}")

    merge_list_path = OUTPUT_DIR / "merge_input_files.txt"
    merge_list_path.write_text(
        "\n".join(str(ppt_path) for ppt_path in ppt_paths),
        encoding="utf-8",
    )

    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(merge_script),
        "-OutputFile",
        str(output_path),
        "-InputListFile",
        str(merge_list_path),
    ]

    result = subprocess.run(command, capture_output=True, text=False)
    stdout_text = decode_subprocess_output(result.stdout)
    stderr_text = decode_subprocess_output(result.stderr)

    if result.returncode != 0 and is_valid_presentation(output_path):
        warning_text = (
            f"PowerPoint merge returned a non-zero exit code but the merged file is valid.\n"
            f"Exit code: {result.returncode}\n"
            f"STDOUT:\n{stdout_text}\n"
            f"STDERR:\n{stderr_text}"
        )
        write_merge_error_file(warning_text)
        return

    if result.returncode != 0:
        error_text = (
            f"PowerPoint merge failed.\n"
            f"Exit code: {result.returncode}\n"
            f"STDOUT:\n{stdout_text}\n"
            f"STDERR:\n{stderr_text}"
        )
        write_merge_error_file(error_text)
        raise RuntimeError(error_text)


def build_mock_plan() -> dict[str, Any]:
    slides: list[dict[str, Any]] = [
        {
            "slide_number": 1,
            "slide_role": "objective",
            "source_slide": 2,
            "title": "학습 목표",
            "why": "학습 목표를 먼저 이해해야 1시간 수업의 흐름을 따라가기 쉽습니다.",
            "bullets": [
                "생성형 AI와 LLM 개념 구분이 먼저임.",
                "기존 자동화와 차이 비교가 핵심임.",
                "업무 변화가 생기는 지점 확인 필요.",
                "환각과 통제 한계 이해도 필요함.",
            ],
            "example": "오늘 수업은 개념 이해, 업무 변화, 활용 사례, 한계, 실습 문제 설계 순서로 진행됩니다.",
            "practice_prompt": "",
            "transition": "먼저 생성형 AI가 무엇인지부터 정의해야 이후 사례를 같은 기준으로 해석할 수 있습니다.",
            "image_prompt": "교실 수업용 인포그래픽, 생성형 AI 학습 목표 4개를 아이콘으로 정리한 깔끔한 프레젠테이션 스타일",
        }
    ]

    content_titles = [
        "생성형 AI란 무엇인가",
        "LLM은 어떤 역할을 하는가",
        "생성형 AI가 기존 AI와 다른 이유",
        "기존 자동화는 어떤 방식으로 동작했는가",
        "생성형 AI 자동화는 무엇이 달라졌는가",
        "두 방식의 차이를 표로 비교",
        "보고서 작성 업무는 어떻게 바뀌는가",
        "데이터 분석 업무는 어떻게 바뀌는가",
        "문서 요약 업무는 어떻게 바뀌는가",
        "기업은 왜 생성형 AI를 도입하는가",
        "현장에서 자주 쓰는 활용 사례",
        "활용 효과를 수치로 보면 무엇이 보이는가",
        "생성형 AI의 가장 큰 한계는 무엇인가",
        "환각은 왜 발생하는가",
        "통제가 어려운 이유는 무엇인가",
        "안전하게 쓰기 위한 운영 원칙",
    ]

    for idx, title in enumerate(content_titles, start=2):
        source_slide = 2

        slide = {
            "slide_number": idx,
            "slide_role": "content",
            "source_slide": source_slide,
            "title": title,
            "why": f"{title}를 알아야 생성형 AI를 업무 변화 관점에서 정확히 설명할 수 있습니다.",
            "bullets": [
                f"{title}의 핵심 뜻 이해가 먼저임.",
                "중요한 특징을 짧게 정리한 내용임.",
                "실제 업무와 연결되는 지점 설명 중심.",
                "다음 내용으로 이어지게 구성한 흐름임.",
            ],
            "example": f"{title} 예시: 주간 업무 보고 초안을 AI가 먼저 만들고 담당자가 검토해 최종본을 확정하는 방식입니다.",
            "practice_prompt": "너는 실무 교육 강사야. 위 개념을 초보 직장인이 이해할 수 있도록 4문장으로 설명해줘.",
            "transition": "이 개념을 이해하면 다음 장표에서 실제 업무 변화와 연결해서 볼 수 있습니다.",
            "image_prompt": f"프레젠테이션용 교육 일러스트, {title}를 설명하는 깔끔한 업무 장면, 파란색 계열, 텍스트 없음",
        }

        slides.append(slide)

    slides.append(
        {
            "slide_number": 18,
            "slide_role": "practice_problem",
            "source_slide": 2,
            "title": "실습 문제",
            "why": "학습 내용을 문제로 바꿔야 개념 이해가 실제 설명 능력으로 전환됩니다.",
            "bullets": [
                "생성형 AI와 기존 자동화 차이 설명 문제임.",
                "정의와 차이점, 업무 예시 포함 필요.",
                "발표하듯 자연스럽게 답하는 것이 핵심임.",
            ],
            "example": "문제 정의: 기존 자동화와 생성형 AI 자동화의 차이를 교육생에게 설명하는 슬라이드 원고를 작성하시오.",
            "practice_prompt": "요구사항: 정의, 차이점, 실제 업무 사례를 포함하고 문장은 발표 가능한 길이로 작성하시오.",
            "transition": "다음 장표에서는 이 문제의 기대 답안 구조를 확인합니다.",
            "image_prompt": "실습 문제 슬라이드용 심플한 교육 아이콘, 문제 해결과 발표 준비를 상징하는 장면",
        }
    )
    slides.append(
        {
            "slide_number": 19,
            "slide_role": "practice_answer",
            "source_slide": 2,
            "title": "실습 문제 정답 예시",
            "why": "기대 답안 구조를 먼저 보면 어떤 수준으로 답해야 하는지 기준이 분명해집니다.",
            "bullets": [
                "정답은 정의부터 차근히 설명하는 구조임.",
                "차이점과 업무 예시를 함께 보여주는 답안임.",
                "마지막에는 검토 필요성까지 포함함.",
            ],
            "example": "정답 구조 예시: 1) 기존 자동화 정의 2) 생성형 AI 자동화 정의 3) 차이점 2가지 4) 보고서 작성 예시 5) 검토 필요성 정리",
            "practice_prompt": "예시 답안: 기존 자동화는 규칙에 따라 같은 결과를 반복하지만, 생성형 AI 자동화는 자연어 요청에 따라 문맥에 맞는 초안을 생성합니다. 따라서 결과 활용 전 검토가 반드시 필요합니다.",
            "transition": "이제 오늘 학습 내용을 마지막으로 한 번 더 정리합니다.",
            "image_prompt": "교육용 정답 예시 슬라이드, 발표 구조와 체크포인트를 보여주는 프레젠테이션 스타일",
        }
    )
    slides.append(
        {
            "slide_number": 20,
            "slide_role": "summary",
            "source_slide": 2,
            "title": "핵심 요약",
            "why": "마지막 정리를 통해 오늘 배운 내용을 한 번에 회상할 수 있습니다.",
            "bullets": [
                "생성형 AI는 새 내용을 만드는 기술임.",
                "LLM은 그 중심에서 언어를 처리함.",
                "업무 자동화 범위가 더 넓어지는 흐름임.",
                "결과는 반드시 검토가 필요함.",
                "좋은 프롬프트와 검증이 함께 필요함.",
            ],
            "example": "",
            "practice_prompt": "",
            "transition": "",
            "image_prompt": "강의 마무리 요약 인포그래픽, 핵심 포인트 5개를 정리한 프레젠테이션 스타일",
        }
    )

    return {
        "deck_title": "생성형 AI 개요 및 업무 변화 이해",
        "section_label": "1. 생성형 AI와 프롬프트",
        "slides": slides,
    }


def normalize_slide_plan(plan: dict[str, Any], template_analysis: dict[str, Any]) -> dict[str, Any]:
    valid_source_slides = {
        item["source_slide"] for item in template_analysis["candidate_templates"]
    }
    for slide in plan["slides"]:
        source_slide = slide.get("source_slide")
        if source_slide not in valid_source_slides:
            raise ValueError(
                f"LLM selected source_slide={source_slide}, which is not in template candidates."
            )
    return plan


def generate_slide_plan(
    prompt_text: str,
    template_analysis: dict[str, Any],
    model: str,
    use_mock: bool,
) -> tuple[dict[str, Any], str]:
    if use_mock:
        plan = normalize_slide_plan(build_mock_plan(), template_analysis)
        raw_text = json.dumps(plan, ensure_ascii=False, indent=2)
        return plan, raw_text
    plan, raw_text = generate_slide_plan_with_llm(prompt_text, template_analysis, model)
    return normalize_slide_plan(plan, template_analysis), raw_text


def render_presentation(template_path: Path, plan: dict[str, Any], output_path: Path) -> None:
    target_prs = Presentation(str(template_path))
    section_label = plan.get("section_label", plan.get("deck_title", "강의안"))

    original_slide_count = len(target_prs.slides)
    generated_slides: list[tuple[Any, Any, dict[str, Any]]] = []
    for slide_data in plan["slides"]:
        source_slide = target_prs.slides[slide_data["source_slide"] - 1]
        generated_slides.append(
            (clone_slide(target_prs, slide_data["source_slide"] - 1), source_slide, slide_data)
        )

    for slide, source_slide, slide_data in generated_slides:
        fill_slide(slide, source_slide, slide_data, section_label)

    remove_slides_by_indices(target_prs, list(range(original_slide_count)))
    target_prs.save(str(output_path))


def render_presentation_with_images(
    template_path: Path,
    plan: dict[str, Any],
    output_path: Path,
    skip_images: bool,
    image_output_dir: Path,
) -> None:
    target_prs = Presentation(str(template_path))
    section_label = plan.get("section_label", plan.get("deck_title", "강의안"))
    total_slides = len(plan["slides"])
    client = None if skip_images else get_openai_client()
    original_slide_count = len(target_prs.slides)

    for index, slide_data in enumerate(plan["slides"], start=1):
        print(f"[{index}/{total_slides}] 슬라이드 생성 중: {slide_data['title']}")
        source_slide = target_prs.slides[slide_data["source_slide"] - 1]
        slide = clone_slide(target_prs, slide_data["source_slide"] - 1)
        fill_slide(slide, source_slide, slide_data, section_label)

        image_prompt = (slide_data.get("image_prompt") or "").strip()
        cached_image_path = get_cached_image_path(image_output_dir, index)

        if cached_image_path.exists():
            replace_slide_image(slide, source_slide, cached_image_path)
            print(f"[{index}/{total_slides}] 기존 이미지 재사용")
            continue

        if image_prompt and client:
            print(f"[{index}/{total_slides}] 이미지 생성 중")
            try:
                image_path = generate_slide_image(
                    client=client,
                    image_prompt=image_prompt,
                    output_dir=image_output_dir,
                    slide_number=index,
                )
                replace_slide_image(slide, source_slide, image_path)
                print(f"[{index}/{total_slides}] 이미지 적용 완료")
            except Exception as exc:
                print(f"[{index}/{total_slides}] 이미지 생성 실패, 템플릿 이미지 유지: {exc}")
        else:
            print(f"[{index}/{total_slides}] 이미지 생성 건너뜀")

    remove_slides_by_indices(target_prs, list(range(original_slide_count)))
    target_prs.save(str(output_path))


def main() -> None:
    load_env_file(DEFAULT_ENV_FILE)
    ensure_dirs()
    args = parse_args()
    image_generation_enabled = env_flag("OPENAI_ENABLE_IMAGE_GENERATION", False)

    template_path = (
        Path(args.template).expanduser().resolve()
        if args.template
        else find_default_template()
    )
    curriculum_path = (
        Path(args.curriculum_file).expanduser().resolve()
        if args.curriculum_file
        else DEFAULT_CURRICULUM_FILE.resolve()
    )
    prompt_file = Path(args.prompt_file).expanduser().resolve()
    output_path = (
        Path(args.output).expanduser().resolve()
        if args.output
        else build_output_path(template_path)
    )
    json_output_path = (
        Path(args.json_output).expanduser().resolve()
        if args.json_output
        else build_json_output_path()
    )
    raw_output_path = build_raw_output_path()
    analysis_output_path = (
        Path(args.analysis_output).expanduser().resolve()
        if args.analysis_output
        else build_analysis_output_path()
    )

    template_analysis = analyze_template(template_path)
    analysis_output_path.write_text(
        json.dumps(template_analysis, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    if args.analyze_only:
        print(f"Template analysis saved: {analysis_output_path}")
        return

    prompt_template = read_prompt_file(prompt_file)
    sessions = parse_curriculum_file(curriculum_path)

    if not sessions:
        try:
            print("[단일 작업] LLM 호출 시작")
            slide_plan, raw_response_text = generate_slide_plan(
                prompt_template, template_analysis, args.model, args.mock
            )

            json_output_path.write_text(
                json.dumps(slide_plan, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            raw_output_path.write_text(raw_response_text, encoding="utf-8")
            render_presentation_with_images(
                template_path=template_path,
                plan=slide_plan,
                output_path=output_path,
                skip_images=args.mock or args.skip_images or (not image_generation_enabled),
                image_output_dir=build_image_output_dir_for_session(),
            )
        except Exception as exc:
            error_text = (
                "[단일 작업] 실패\n"
                f"오류 단계: LLM 호출 또는 PPT 생성\n"
                f"오류 타입: {type(exc).__name__}\n"
                f"오류 메시지: {exc}\n\n"
                f"{traceback.format_exc()}"
            )
            write_session_error_file("single_run", error_text)
            append_error_log(error_text)
            print("[단일 작업] 실패: output/single_run_error.txt 확인")
            raise

        print(f"Template analysis: {analysis_output_path}")
        print(f"Prompt file: {prompt_file}")
        print(f"LLM raw output: {raw_output_path}")
        print(f"Slide JSON: {json_output_path}")
        print(f"Created PPT: {output_path}")
        return

    total_sessions = len(sessions)
    print(f"Curriculum sessions detected: {total_sessions}")
    continue_on_error = env_flag("CONTINUE_ON_SESSION_ERROR", True)
    generated_ppt_paths: list[Path] = []

    for session_index, session in enumerate(sessions, start=1):
        session_slug = slugify_korean(session["title"])
        session_prefix = f"{session['session_no']:02d}_{session_slug}"
        session_prompt_text = render_session_prompt(prompt_template, session)
        session_prompt_path = OUTPUT_DIR / f"{session_prefix}_prompt.txt"
        session_json_path = OUTPUT_DIR / f"{session_prefix}_slide_plan.json"
        session_raw_path = OUTPUT_DIR / f"{session_prefix}_llm_raw_response.txt"
        session_ppt_path = OUTPUT_DIR / f"{session_prefix}.pptx"

        print(
            f"[교시 {session_index}/{total_sessions}] 생성 시작: {session['session_no']}교시 {session['title']}"
        )
        session_prompt_path.write_text(session_prompt_text, encoding="utf-8")

        try:
            print(f"[교시 {session_index}/{total_sessions}] LLM 호출 시작")
            slide_plan, raw_response_text = generate_slide_plan(
                session_prompt_text, template_analysis, args.model, args.mock
            )

            session_json_path.write_text(
                json.dumps(slide_plan, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            session_raw_path.write_text(raw_response_text, encoding="utf-8")

            print(f"[교시 {session_index}/{total_sessions}] PPT 렌더링 시작")
            render_presentation_with_images(
                template_path=template_path,
                plan=slide_plan,
                output_path=session_ppt_path,
                skip_images=args.mock or args.skip_images or (not image_generation_enabled),
                image_output_dir=build_image_output_dir_for_session(session_prefix),
            )

            generated_ppt_paths.append(session_ppt_path)
            print(f"[교시 {session_index}/{total_sessions}] 완료: {session_ppt_path}")
        except Exception as exc:
            error_text = (
                f"[교시 {session_index}/{total_sessions}] 실패\n"
                f"교시: {session['session_no']}교시 {session['title']}\n"
                f"오류 타입: {type(exc).__name__}\n"
                f"오류 메시지: {exc}\n\n"
                f"{traceback.format_exc()}"
            )
            error_file = write_session_error_file(session_prefix, error_text)
            append_error_log(error_text)
            print(f"[교시 {session_index}/{total_sessions}] 실패: {error_file}")
            if not continue_on_error:
                raise

    merged_output_path = build_merged_output_path()
    if generated_ppt_paths:
        print(f"[병합] 교시별 PPT {len(generated_ppt_paths)}개를 하나로 병합 중")
        merge_presentations(generated_ppt_paths, merged_output_path)
        print(f"[병합] 완료: {merged_output_path}")

    print(f"Template analysis: {analysis_output_path}")
    print(f"Curriculum file: {curriculum_path}")
    print(f"Output folder: {OUTPUT_DIR}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("[취소] Ctrl+C로 생성 작업이 중단되었습니다.")
        raise SystemExit(130)
