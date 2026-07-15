"""Content parser — parse content.json slides[] to PageContent-like dicts."""

from __future__ import annotations

import os
import re
from typing import Any


_GOAL_KEYWORDS: list[tuple[str, list[str]]] = [
    ("problem", ["problem", "痛点", "挑战", "问题", "pain point", "challenge"]),
    ("solution", ["solution", "方案", "解决", "approach", "how we"]),
    ("features", ["features", "功能", "特性", "能力", "capability", "what we offer"]),
    ("data", ["architecture", "架构", "技术", "系统", "system", "tech stack"]),
    ("code", ["quick start", "快速开始", "getting started", "installation", "安装", "usage", "使用方法", "example", "示例"]),
    ("cta", ["contact", "联系", "get in touch", "next step", "下一步", "开始"]),
    ("overview", ["overview", "概述", "目录", "agenda", "议程", "table of content"]),
]

_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".svg", ".gif"}

_COMPONENT_KEYWORDS: list[tuple[str, str, list[str]]] = [
    ("group", "swot", ["优势", "劣势", "机会", "威胁", "strength", "weakness", "opportunity", "threat", "swot", "对比", "比较"]),
    ("group", "hierarchy", ["组织", "架构", "层级", "汇报", "hierarchy", "org chart", "CEO", "CTO", "CFO", "VP", "director"]),
    ("group", "process", ["流程", "步骤", "阶段", "process", "step", "phase", "pipeline", "workflow"]),
    ("group", "timeline", ["时间线", "里程碑", "timeline", "milestone", "路线图", "roadmap"]),
]


def infer_component_category(bullets: list[str]) -> tuple[str | None, str | None]:
    n = len(bullets)
    if n < 2 or n > 9:
        return (None, None)

    combined = " ".join(bullets).lower()

    for comp_type, category, keywords in _COMPONENT_KEYWORDS:
        if any(kw.lower() in combined for kw in keywords):
            return (comp_type, category)

    if 3 <= n <= 8:
        return ("group", "process")

    return (None, None)


def _infer_goal(title: str, page_index: int, total_pages: int) -> str:
    if page_index == 0 and total_pages > 1:
        return "hook"
    if page_index == 1 and total_pages > 2:
        return "overview"
    title_lower = title.lower()
    for goal, keywords in _GOAL_KEYWORDS:
        if any(kw in title_lower for kw in keywords):
            return goal
    if page_index == 0:
        return "hook"
    if page_index == 1:
        return "overview"
    return "content"


def _clean_md_formatting(text: str) -> str:
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"\*(.+?)\*", r"\1", text)
    text = re.sub(r"`(.+?)`", r"\1", text)
    return text


def _is_html_line(line: str) -> bool:
    s = line.strip()
    return s.startswith("<") and (">" in s)


def _split_large_table(rows: list[list[str]], max_rows: int = 8) -> list[list[list[str]]]:
    if len(rows) <= max_rows:
        return [rows]
    header = rows[0] if rows else []
    data = rows[1:]
    chunks: list[list[list[str]]] = []
    for start in range(0, len(data), max_rows - 1):
        chunk = [header] + data[start:start + max_rows - 1]
        chunks.append(chunk)
    return chunks


def _parse_markdown_sections(text: str) -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    in_code_block = False
    code_lines: list[str] = []
    code_lang = ""

    raw_sections: list[dict[str, Any]] = []
    for line in text.split("\n"):
        stripped = line.strip()

        if stripped.startswith("```"):
            if in_code_block:
                in_code_block = False
                if current is not None:
                    code_text = "\n".join(code_lines)
                    current.setdefault("code_blocks", []).append({
                        "language": code_lang,
                        "code": code_text,
                    })
                code_lines = []
                code_lang = ""
            else:
                in_code_block = True
                code_lang = stripped[3:].strip()
                code_lines = []
            continue

        if in_code_block:
            code_lines.append(line)
            continue

        if stripped.startswith("# "):
            if current is not None:
                raw_sections.append(current)
            current = {"title": stripped[2:].strip(), "level": 1, "bullets": [], "text_lines": [], "code_blocks": [], "images": [], "tables": []}
            continue

        if current is None:
            current = {"title": "", "level": 1, "bullets": [], "text_lines": [], "code_blocks": [], "images": [], "tables": []}

        if stripped.startswith("## "):
            raw_sections.append(current)
            current = {"title": stripped[3:].strip(), "level": 2, "bullets": [], "text_lines": [], "code_blocks": [], "images": [], "tables": []}
            continue

        if stripped.startswith("|") and "|" in stripped[1:]:
            current["tables"].append(stripped)
            continue

        img_match = re.match(r"!\[.*?\]\((.+?)\)", stripped)
        if img_match:
            current["images"].append(img_match.group(1))
            continue

        if stripped.startswith("- ") or stripped.startswith("* "):
            current["bullets"].append(stripped[2:].strip())
            continue

        if stripped.startswith(tuple(f"{i}. " for i in range(1, 10))):
            current["bullets"].append(re.sub(r"^\d+\.\s*", "", stripped))
            continue

        if stripped and not _is_html_line(stripped):
            current["text_lines"].append(stripped)

    if current is not None:
        raw_sections.append(current)

    h1_count = sum(1 for s in raw_sections if s.get("level") == 1)
    if h1_count > 1:
        for s in raw_sections:
            s.pop("level", None)
            sections.append(s)
        return sections

    for s in raw_sections:
        s.pop("level", None)
        sections.append(s)
    return sections


def _resolve_image(image_ref: str, project_dir: str) -> str | None:
    if os.path.isabs(image_ref):
        return image_ref if os.path.isfile(image_ref) else None
    candidate = os.path.join(project_dir, image_ref)
    return candidate if os.path.isfile(candidate) else None


def parse_readme(readme_path: str, project_dir: str) -> list[dict[str, Any]]:
    if not os.path.isfile(readme_path):
        return []

    try:
        with open(readme_path, encoding="utf-8") as f:
            text = f.read()
    except OSError:
        return []

    if not text.strip():
        return []

    sections = _parse_markdown_sections(text)
    if not sections:
        return []

    merged: list[dict[str, Any]] = []
    for sec in sections:
        has_content = (sec.get("title") or sec.get("bullets") or sec.get("tables")
                       or sec.get("code_blocks") or sec.get("images") or sec.get("text_lines"))
        if not has_content:
            if merged:
                merged[-1].setdefault("text_lines", []).extend(sec.get("text_lines", []))
            continue
        if not sec.get("title") and merged:
            merged[-1].setdefault("bullets", []).extend(sec.get("bullets", []))
            merged[-1].setdefault("tables", []).extend(sec.get("tables", []))
            merged[-1].setdefault("code_blocks", []).extend(sec.get("code_blocks", []))
            merged[-1].setdefault("images", []).extend(sec.get("images", []))
            merged[-1].setdefault("text_lines", []).extend(sec.get("text_lines", []))
            continue
        merged.append(sec)

    pages: list[dict[str, Any]] = []
    total_pages = len(merged)

    for i, section in enumerate(merged):
        title = section.get("title", "")
        goal = _infer_goal(title, i, total_pages)

        bullets: list[str] = [_clean_md_formatting(b) for b in section.get("bullets", [])]

        code_blocks = section.get("code_blocks", [])
        code = None
        if code_blocks:
            first = code_blocks[0]
            if isinstance(first, dict):
                code = first

        tables = section.get("tables", [])
        table_rows: list[list[str]] = []
        for row_line in tables:
            if re.match(r"^\|[\s\-:|]+\|$", row_line):
                continue
            cells = [_clean_md_formatting(c.strip()) for c in row_line.split("|") if c.strip()]
            if cells:
                table_rows.append(cells)

        if table_rows:
            table_chunks = _split_large_table(table_rows)
            for ti, chunk in enumerate(table_chunks):
                page_title = title if ti == 0 else f"{title} ({ti + 1})"
                page_goal = goal if ti == 0 else "content"
                page: dict[str, Any] = {
                    "goal": page_goal,
                    "title": page_title,
                    "bullets": bullets if ti == 0 and bullets else None,
                    "image": None,
                    "code": code if ti == 0 else None,
                    "diagram_type": "table",
                    "diagram_data": {"type": "table", "rows": chunk},
                }
                pages.append(page)
            continue

        image = None
        for img_ref in section.get("images", []):
            resolved = _resolve_image(img_ref, project_dir)
            if resolved:
                image = resolved
                break

        page: dict[str, Any] = {
            "goal": goal,
            "title": title,
            "bullets": bullets if bullets else None,
            "image": image,
            "code": code,
            "diagram_type": None,
            "diagram_data": None,
        }

        notes_parts: list[str] = []
        for tl in section.get("text_lines", []):
            notes_parts.append(tl)
        if notes_parts:
            page["notes"] = "\n".join(notes_parts)

        pages.append(page)

    return pages


def load_enterprise_content(
    content_raw: dict[str, Any],
    project_dir: str,
) -> list[dict[str, Any]]:
    meta = content_raw.get("meta", {})
    slides = content_raw.get("slides", [])
    result: list[dict[str, Any]] = []

    for i, slide_data in enumerate(slides):
        image = slide_data.get("image")
        if image and not os.path.isabs(image):
            image = os.path.join(project_dir, image)

        cards = slide_data.get("cards")
        if cards:
            resolved_cards: list[dict[str, Any]] = []
            for card in cards:
                card_img = card.get("image")
                if card_img and not os.path.isabs(card_img):
                    card_img = os.path.join(project_dir, card_img)
                resolved = {**card}
                if card_img is not None:
                    resolved["image"] = card_img
                resolved_cards.append(resolved)
            cards = resolved_cards

        diagram = slide_data.get("diagram")
        diagram_type = diagram.get("type") if isinstance(diagram, dict) else None
        diagram_data = diagram if isinstance(diagram, dict) else None

        code = slide_data.get("code")
        exercise = slide_data.get("exercise")

        title = slide_data.get("title")
        if title is None and i == 0:
            title = meta.get("title")

        subtitle = slide_data.get("subtitle")
        if subtitle is None and i == 0:
            subtitle = meta.get("subtitle")

        result.append({
            "goal": slide_data.get("goal", "content"),
            "title": title or "",
            "subtitle": subtitle,
            "bullets": slide_data.get("bullets"),
            "image": image,
            "cards": cards,
            "diagram_type": diagram_type,
            "diagram_data": diagram_data,
            "code": code,
            "exercise": exercise,
            "notes": slide_data.get("notes"),
            "links": slide_data.get("links"),
            "chart": slide_data.get("chart"),
            "image_grid": slide_data.get("image_grid"),
            "icons": slide_data.get("icons"),
            "component_type": slide_data.get("component_type"),
            "component_category": slide_data.get("component_category"),
            "component_variant": slide_data.get("component_variant"),
            "blocks": slide_data.get("blocks"),
            "elements": slide_data.get("elements"),
        })

    return result
