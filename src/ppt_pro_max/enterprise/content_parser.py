"""Content parser — parse content.json slides[] to PageContent-like dicts."""

from __future__ import annotations

import os
from typing import Any


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
