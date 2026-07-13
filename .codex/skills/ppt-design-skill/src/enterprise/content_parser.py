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


def _parse_markdown_sections(text: str) -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    in_code_block = False
    code_lines: list[str] = []
    code_lang = ""

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
                sections.append(current)
            current = {"title": stripped[2:].strip(), "h2s": [], "bullets": [], "text_lines": [], "code_blocks": [], "images": [], "tables": []}
            continue

        if current is None:
            current = {"title": "", "h2s": [], "bullets": [], "text_lines": [], "code_blocks": [], "images": [], "tables": []}

        if stripped.startswith("## "):
            current["h2s"].append(stripped[3:].strip())
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

        if stripped:
            current["text_lines"].append(stripped)

    if current is not None:
        sections.append(current)

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

    pages: list[dict[str, Any]] = []
    total_pages = len(sections)

    for i, section in enumerate(sections):
        title = section.get("title", "")
        goal = _infer_goal(title, i, total_pages)

        bullets: list[str] = list(section.get("bullets", []))
        for h2 in section.get("h2s", []):
            bullets.append(h2)

        code_blocks = section.get("code_blocks", [])
        code = None
        if code_blocks:
            first = code_blocks[0]
            if isinstance(first, dict):
                code = first

        diagram_type = None
        diagram_data = None
        tables = section.get("tables", [])
        if tables:
            diagram_type = "table"
            rows: list[list[str]] = []
            for row_line in tables:
                if re.match(r"^\|[\s\-:|]+\|$", row_line):
                    continue
                cells = [c.strip() for c in row_line.split("|") if c.strip()]
                if cells:
                    rows.append(cells)
            if rows:
                diagram_data = {"type": "table", "rows": rows}

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
            "diagram_type": diagram_type,
            "diagram_data": diagram_data,
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
        })

    return result
