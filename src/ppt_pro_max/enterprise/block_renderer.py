"""BlockRenderer — composable blocks layout system.

Renders individual content blocks within specified regions on a slide.
Each block type delegates to PrecisionRenderer primitives for brand compliance.
"""

from __future__ import annotations

import os
from typing import Any


class BlockRenderer:

    REGION_PRESETS: dict[str, dict[str, float]] = {
        "full":          {"x": 0.9,   "y": 1.6,  "w": 11.533, "h": 5.4},
        "left":          {"x": 0.9,   "y": 1.6,  "w": 5.566,  "h": 5.4},
        "right":         {"x": 6.866, "y": 1.6,  "w": 5.566,  "h": 5.4},
        "left-2-3":      {"x": 0.9,   "y": 1.6,  "w": 7.422,  "h": 5.4},
        "right-1-3":     {"x": 8.722, "y": 1.6,  "w": 3.711,  "h": 5.4},
        "left-1-3":      {"x": 0.9,   "y": 1.6,  "w": 3.711,  "h": 5.4},
        "right-2-3":     {"x": 5.011, "y": 1.6,  "w": 7.422,  "h": 5.4},
        "top":           {"x": 0.9,   "y": 1.6,  "w": 11.533, "h": 2.5},
        "bottom":        {"x": 0.9,   "y": 4.5,  "w": 11.533, "h": 2.5},
        "top-left":      {"x": 0.9,   "y": 1.6,  "w": 5.566,  "h": 2.5},
        "top-right":     {"x": 6.866, "y": 1.6,  "w": 5.566,  "h": 2.5},
        "bottom-left":   {"x": 0.9,   "y": 4.5,  "w": 5.566,  "h": 2.5},
        "bottom-right":  {"x": 6.866, "y": 4.5,  "w": 5.566,  "h": 2.5},
        "center":        {"x": 2.5,   "y": 2.0,  "w": 8.333,  "h": 4.6},
        "sidebar":       {"x": 0.0,   "y": 1.6,  "w": 2.5,   "h": 5.4},
    }

    def __init__(self, precision_renderer: Any) -> None:
        self._pr = precision_renderer

    def render(self, slide: Any, blocks: list[dict], is_hero: bool = False) -> None:
        for i, block in enumerate(blocks):
            region = self._resolve_region(block, i, len(blocks), is_hero)
            self._render_block(slide, block, region)

    def _resolve_region(self, block: dict, block_index: int,
                        total_blocks: int, is_hero: bool) -> dict:
        region_spec = block.get("region")
        if isinstance(region_spec, dict):
            return region_spec
        if isinstance(region_spec, str):
            return self.REGION_PRESETS.get(region_spec, self.REGION_PRESETS["full"])
        return self._auto_region(block_index, total_blocks, is_hero)

    def _auto_region(self, block_index: int, total_blocks: int,
                     is_hero: bool) -> dict:
        if total_blocks == 1:
            return self.REGION_PRESETS["full"]
        elif total_blocks == 2:
            return self.REGION_PRESETS["left" if block_index == 0 else "right"]
        elif total_blocks == 3:
            if block_index == 0:
                return self.REGION_PRESETS["left-2-3"]
            return self.REGION_PRESETS["top-right" if block_index == 1 else "bottom-right"]
        elif total_blocks == 4:
            quadrant = ["top-left", "top-right", "bottom-left", "bottom-right"]
            return self.REGION_PRESETS[quadrant[block_index]]
        return self.REGION_PRESETS["full"]

    def _render_block(self, slide: Any, block: dict, region: dict) -> None:
        block_type = block.get("type", "bullets")
        data = block.get("data", {})
        style = block.get("style", {})
        handler = getattr(self, f"_render_{block_type}_block", None)
        if handler:
            handler(slide, data, region, style)
        else:
            self._render_bullets_block(slide, data, region, style)

    # ── cards ──

    def _render_cards_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        items = data.get("items", [])
        layout = data.get("layout", "horizontal")
        featured_idx = data.get("featured_index", 0)
        if not items:
            return
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        n = len(items)
        if layout == "vertical":
            card_h = min(2.0, (h - 0.3 * (n - 1)) / n)
            card_w = w
            for i, item in enumerate(items):
                cy = y + i * (card_h + 0.3)
                title = item.get("title", "")
                body = item.get("text", item.get("body", ""))
                self._pr.add_card(slide, x, cy, card_w, card_h, title, body,
                                   featured=(i == featured_idx))
        elif layout == "grid-2x2":
            cols = 2
            rows = (n + 1) // 2
            gap_x = 0.3
            gap_y = 0.3
            card_w = (w - gap_x * (cols - 1)) / cols
            card_h = (h - gap_y * (rows - 1)) / rows
            for i, item in enumerate(items):
                col = i % cols
                row = i // cols
                cx = x + col * (card_w + gap_x)
                cy = y + row * (card_h + gap_y)
                title = item.get("title", "")
                body = item.get("text", item.get("body", ""))
                self._pr.add_card(slide, cx, cy, card_w, card_h, title, body,
                                   featured=(i == featured_idx))
        else:
            card_w = min(3.6, (w - 0.4 * (n - 1)) / n)
            for i, item in enumerate(items):
                cx = x + i * (card_w + 0.4)
                title = item.get("title", "")
                body = item.get("text", item.get("body", ""))
                self._pr.add_card(slide, cx, y, card_w, h, title, body,
                                   featured=(i == featured_idx))

    # ── bullets ──

    def _render_bullets_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        items = data.get("items", [])
        if not items:
            return
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        columns = data.get("columns")
        bullet_style = data.get("bullet_style", "dot")
        font_size = style.get("font_size", 14)
        color_role = style.get("color_role", "foreground")

        if columns is None:
            columns = 2 if len(items) >= 6 else 1

        prefix_map = {"dot": "\u2022  ", "dash": "\u2014  ", "number": None, "none": ""}
        prefix = prefix_map.get(bullet_style, "\u2022  ")

        if columns >= 2:
            col_w = (w - 0.3) / 2
            mid = (len(items) + 1) // 2
            left_lines = []
            right_lines = []
            for i, item in enumerate(items):
                if bullet_style == "number":
                    line = f"{i+1}. {item}"
                else:
                    line = f"{prefix}{item}"
                if i < mid:
                    left_lines.append(line)
                else:
                    right_lines.append(line)
            self._pr.add_multiline(slide, left_lines, x, y, col_w, h,
                                    size=font_size, color_role=color_role, spacing=6)
            self._pr.add_multiline(slide, right_lines, x + col_w + 0.3, y, col_w, h,
                                    size=font_size, color_role=color_role, spacing=6)
        else:
            lines = []
            for i, item in enumerate(items):
                if bullet_style == "number":
                    lines.append(f"{i+1}. {item}")
                else:
                    lines.append(f"{prefix}{item}")
            self._pr.add_multiline(slide, lines, x, y, w, h,
                                    size=font_size, color_role=color_role, spacing=6)

    # ── diagram ──

    def _render_diagram_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        diagram_type = data.get("diagram_type")
        diagram_data = data.get("diagram_data")
        if not diagram_type or not diagram_data:
            return
        from ppt_pro_max.renderer.diagram_engine import DiagramEngine
        from ppt_pro_max.renderer.diagram.diagram_style import DiagramStyle
        from ppt_pro_max.renderer.diagram.layout_engine import Region
        ds = DiagramStyle.from_brand_spec(self._pr.brand) if self._pr._brand else DiagramStyle()
        r = Region(left=region["x"], top=region["y"], width=region["w"], height=region["h"])
        engine = DiagramEngine()
        try:
            engine.render(slide, diagram_type, diagram_data, ds, r)
        except Exception:
            pass

    # ── code ──

    def _render_code_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        source = data.get("source", "")
        language = data.get("language", "")
        bg_color = style.get("bg_color", "#1E293B")
        font = style.get("font", "Consolas")
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        self._pr.add_rounded_rect(slide, x, y, w, h, fill_hex=bg_color, corner_radius="lg")
        if language:
            badge_text = language.upper()
            badge_w = len(badge_text) * 0.1 + 0.3
            self._pr.add_rounded_rect(slide, x + 0.1, y - 0.3, badge_w, 0.3,
                                       fill_role="primary", corner_radius="sm")
            self._pr.add_text(slide, badge_text, x + 0.2, y - 0.28, badge_w - 0.2, 0.26,
                              size=10, color_role="on-primary", bold=True)
        lines = source.split("\n")
        all_lines = [f"  {line}" for line in lines[:30]]
        self._pr.add_multiline(slide, all_lines, x + 0.3, y + 0.2, w - 0.6, h - 0.4,
                                font=font, size=11, color_role="muted-foreground", spacing=4)

    # ── exercise ──

    def _render_exercise_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        instructions = data.get("instructions", "")
        duration = data.get("duration", "")
        steps = data.get("steps", [])
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        badge_text = f"Exercise {duration}" if duration else "Exercise"
        self._pr.add_badge(slide, badge_text, x, y, variant="solid")
        if instructions:
            self._pr.add_text(slide, instructions, x, y + 0.6, w, 1.2,
                              font=self._pr._font_b(), size=13, color_role="muted-foreground")
        if steps:
            step_lines = [f"{i+1}. {s}" for i, s in enumerate(steps)]
            self._pr.add_multiline(slide, step_lines, x, y + 1.8, w, h - 2.0,
                                    size=13, color_role="foreground", spacing=6)

    # ── image ──

    def _render_image_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        path = data.get("path", "")
        mask = data.get("mask", True)
        cover = data.get("cover", False)
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        if not os.path.isfile(path):
            return
        if cover:
            self._pr.add_image(slide, path, x, y, w, h)
        elif mask:
            self._pr.add_masked_image(slide, path, x, y, w, h)
        else:
            self._pr.add_image(slide, path, x, y, w, h)

    # ── hexagons ──

    def _render_hexagons_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        items = data.get("items", [])
        layout = data.get("layout", "honeycomb")
        if not items:
            return
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        n = len(items)
        if layout == "honeycomb" and n <= 6:
            cols = 2 if n <= 4 else 3
            rows = (n + cols - 1) // cols
        elif layout == "grid-2x2":
            cols, rows = 2, 2
        else:
            cols = n
            rows = 1
        gap_x = 0.4
        gap_y = 0.4
        hex_w = (w - gap_x * max(0, cols - 1)) / cols
        hex_h = (h - gap_y * max(0, rows - 1)) / rows
        size = min(hex_w, hex_h * 1.15)
        for i, item in enumerate(items):
            col = i % cols
            row = i // cols
            cx = x + col * (hex_w + gap_x) + hex_w / 2 - size / 2
            cy = y + row * (hex_h + gap_y) + hex_h / 2 - size * 0.87 / 2
            if layout == "honeycomb" and row % 2 == 1:
                cx += hex_w / 2
            font_size = style.get("font_size", 16)
            color_role = item.get("color_role", "primary")
            self._pr.add_hexagon(slide, cx, cy, size,
                                  fill_role=color_role, gradient=True, shadow=True,
                                  label=item.get("label", ""), font_size=font_size)

    # ── ovals ──

    def _render_ovals_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        items = data.get("items", [])
        layout = data.get("layout", "horizontal")
        show_connectors = data.get("show_connectors", False)
        if not items:
            return
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        n = len(items)
        font_size = style.get("font_size", 16)
        subtitle_size = style.get("subtitle_size", 13)
        if layout == "vertical":
            size = min(w, (h - 0.3 * (n - 1)) / n)
            for i, item in enumerate(items):
                cy = y + i * (size + 0.3)
                cx = x + (w - size) / 2
                color_role = item.get("color_role", "primary")
                self._pr.add_oval(slide, cx, cy, size, size,
                                   fill_role=color_role, gradient=True, shadow=True,
                                   label=item.get("label", ""), font_size=font_size)
                if item.get("subtitle"):
                    self._pr.add_text(slide, item["subtitle"], cx, cy + size + 0.05, size, 0.3,
                                      size=subtitle_size, color_role="muted-foreground", align="center")
        else:
            size = min(h * 0.7, (w - 0.5 * (n - 1)) / n)
            for i, item in enumerate(items):
                cx = x + i * (size + 0.5) + (w / n - size) / 2
                cy = y + (h - size) / 2
                color_role = item.get("color_role", "primary")
                self._pr.add_oval(slide, cx, cy, size, size,
                                   fill_role=color_role, gradient=True, shadow=True,
                                   label=item.get("label", ""), font_size=font_size)
                if item.get("subtitle"):
                    self._pr.add_text(slide, item["subtitle"], cx, cy + size + 0.05, size, 0.3,
                                      size=subtitle_size, color_role="muted-foreground", align="center")

    # ── donuts ──

    def _render_donuts_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        items = data.get("items", [])
        layout = data.get("layout", "horizontal")
        if not items:
            return
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        n = len(items)
        font_size = style.get("font_size", 18)
        subtitle_size = style.get("subtitle_size", 13)
        if layout == "vertical":
            size = min(w, (h - 0.3 * (n - 1)) / n)
            for i, item in enumerate(items):
                cx = x + (w - size) / 2
                cy = y + i * (size + 0.3)
                self._pr.add_donut(slide, cx, cy, size,
                                    fill_role="primary", gradient=True, shadow=True,
                                    label=item.get("label", ""), font_size=font_size)
                if item.get("subtitle"):
                    self._pr.add_text(slide, item["subtitle"], cx, cy + size + 0.05, size, 0.3,
                                      size=subtitle_size, color_role="muted-foreground", align="center")
        else:
            size = min(h, (w - 0.5 * (n - 1)) / n)
            for i, item in enumerate(items):
                cx = x + i * (size + 0.5)
                cy = y + (h - size) / 2
                self._pr.add_donut(slide, cx, cy, size,
                                    fill_role="primary", gradient=True, shadow=True,
                                    label=item.get("label", ""), font_size=font_size)
                if item.get("subtitle"):
                    self._pr.add_text(slide, item["subtitle"], cx, cy + size + 0.05, size, 0.3,
                                      size=subtitle_size, color_role="muted-foreground", align="center")

    # ── metrics ──

    def _render_metrics_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        items = data.get("items", [])
        layout = data.get("layout", "horizontal")
        if not items:
            return
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        n = len(items)
        value_size = style.get("value_size", 40)
        label_size = style.get("label_size", 14)
        if layout == "vertical":
            col_w = w
            row_h = h / n
            for i, item in enumerate(items):
                cx = x
                cy = y + i * row_h
                color_role = item.get("color_role", "primary")
                self._pr.add_rounded_rect(slide, cx + 0.1, cy + 0.05, col_w - 0.2, row_h - 0.1,
                                           fill_role="muted", corner_radius="md", shadow=True)
                self._pr.add_text(slide, item.get("value", ""), cx, cy, col_w, row_h * 0.6,
                                  size=value_size, color_role=color_role, bold=True, align="center")
                self._pr.add_text(slide, item.get("label", ""), cx, cy + row_h * 0.6, col_w, row_h * 0.4,
                                  font=self._pr._font_b(), size=label_size, color_role="muted-foreground", align="center")
        else:
            col_w = w / n
            for i, item in enumerate(items):
                cx = x + i * col_w
                color_role = item.get("color_role", "primary")
                self._pr.add_rounded_rect(slide, cx + 0.15, y + 0.05, col_w - 0.3, h - 0.1,
                                           fill_role="muted", corner_radius="md", shadow=True)
                self._pr.add_text(slide, item.get("value", ""), cx, y, col_w, h * 0.6,
                                  size=value_size, color_role=color_role, bold=True, align="center")
                self._pr.add_text(slide, item.get("label", ""), cx, y + h * 0.6, col_w, h * 0.4,
                                  font=self._pr._font_b(), size=label_size, color_role="muted-foreground", align="center")

    # ── badge ──

    def _render_badge_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        text = data.get("text", "")
        variant = data.get("variant", "default")
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        self._pr.add_badge(slide, text, x, y, variant=variant)

    # ── gradient_line ──

    def _render_gradient_line_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        color_role = data.get("color_role", "accent")
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        color_hex = self._pr._c(color_role, self._pr._c("primary", "#2563EB"))
        self._pr.add_gradient_line(slide, x, y, w, h, color_hex)

    # ── table_chart ──

    def _render_table_chart_block(self, slide: Any, data: dict, region: dict, style: dict) -> None:
        headers = data.get("headers", [])
        rows = data.get("rows", [])
        if not headers and not rows:
            return
        x, y, w, h = region["x"], region["y"], region["w"], region["h"]
        n_cols = len(headers) if headers else (len(rows[0]) if rows else 0)
        if n_cols == 0:
            return
        col_w = w / n_cols
        row_h = min(0.5, h / (len(rows) + 1))
        muted_hex = self._pr._c("muted", "#F1F5F9")
        fg_hex = self._pr._c("foreground", "#1E293B")
        if headers:
            for j, header in enumerate(headers):
                self._pr.add_text(slide, header, x + j * col_w, y, col_w, row_h,
                                  size=12, color_role="foreground", bold=True)
        for i, row in enumerate(rows):
            ry = y + (i + 1) * row_h
            if i % 2 == 0:
                self._pr.add_rect(slide, x, ry, w, row_h, fill_hex=muted_hex)
            for j, cell in enumerate(row):
                if j < n_cols:
                    self._pr.add_text(slide, str(cell), x + j * col_w, ry, col_w, row_h,
                                      font=self._pr._font_b(), size=11, color_role="foreground")
