"""Layout Engine — content-adaptive positioning with Rect and ContentLayout."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Rect:
    width: float = 13.333
    height: float = 7.5
    margin_left: float = 0.9
    margin_right: float = 0.9
    margin_top: float = 0.6
    margin_bottom: float = 0.5

    @property
    def content_width(self) -> float:
        return self.width - self.margin_left - self.margin_right

    @property
    def content_height(self) -> float:
        return self.height - self.margin_top - self.margin_bottom


@dataclass
class ContentLayout:
    title_x: float = 0.9
    title_y: float = 0.5
    title_w: float = 11.533
    title_h: float = 0.6
    content_x: float = 0.9
    content_y: float = 1.6
    content_w: float = 7.0
    image_x: float = 8.3
    image_y: float = 1.2
    image_w: float = 4.2
    bullet_columns: int = 1


class LayoutEngine:

    def compute_content_layout(self, page: dict, slide_rect: Rect) -> ContentLayout:
        title = page.get("title", "")
        bullets = page.get("bullets", [])
        has_image = bool(page.get("image"))

        title_lines = self._estimate_lines(title, slide_rect.content_width, 28)
        title_h = max(0.5, title_lines * 0.4)

        content_y = slide_rect.margin_top + title_h + 0.4

        if len(bullets) <= 4:
            text_w = slide_rect.content_width * 0.65 if has_image else slide_rect.content_width
        elif len(bullets) <= 7:
            text_w = slide_rect.content_width * 0.7 if has_image else slide_rect.content_width
        else:
            text_w = slide_rect.content_width

        bullet_columns = 2 if len(bullets) >= 6 else 1

        if has_image:
            image_w = slide_rect.content_width - text_w - 0.4
            image_x = slide_rect.margin_left + text_w + 0.4
        else:
            image_w = 0
            image_x = 0

        return ContentLayout(
            title_x=slide_rect.margin_left,
            title_y=slide_rect.margin_top,
            title_w=slide_rect.content_width,
            title_h=title_h,
            content_x=slide_rect.margin_left,
            content_y=content_y,
            content_w=text_w,
            image_x=image_x,
            image_y=content_y,
            image_w=image_w,
            bullet_columns=bullet_columns,
        )

    @staticmethod
    def _estimate_lines(text: str, width_inches: float, font_size_pt: int) -> int:
        if not text:
            return 1
        cjk_count = sum(1 for ch in text if '\u4e00' <= ch <= '\u9fff' or '\u3000' <= ch <= '\u303f')
        latin_count = len(text) - cjk_count
        avg_char_w = (latin_count * font_size_pt * 0.012 + cjk_count * font_size_pt * 0.024) / max(1, len(text))
        chars_per_line = max(1, int(width_inches / avg_char_w))
        lines = (len(text) + chars_per_line - 1) // chars_per_line
        return max(1, lines)

    def compute_margins(self, page: dict, mode: str = "balanced") -> dict:
        bullets = page.get("bullets", [])
        cards = page.get("cards", [])
        has_image = bool(page.get("image"))
        content_count = len(bullets) + len(cards) * 3
        if has_image:
            content_count += 2
        if mode == "presenting" or content_count <= 3:
            return {"left": 1.5, "right": 1.5, "top": 1.2, "bottom": 0.8}
        elif mode == "reading" or content_count >= 8:
            return {"left": 0.8, "right": 0.8, "top": 0.6, "bottom": 0.5}
        else:
            return {"left": 0.9, "right": 0.9, "top": 0.6, "bottom": 0.5}
