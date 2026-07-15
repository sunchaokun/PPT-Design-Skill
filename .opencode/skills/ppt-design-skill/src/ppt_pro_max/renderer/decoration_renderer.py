from __future__ import annotations

from typing import Callable


class DecorationRenderer:

    STYLES = [
        "accent_bar", "neon_lines", "gold_trim", "minimal_dots",
        "diamond_bullets", "gradient_bar", "circle_accent",
        "sidebar_nav", "no_decoration", "full_bleed_overlay",
    ]

    def apply_title_decoration(self, slide, title_x: float, title_y: float,
                                title_w: float, style: str,
                                colors: dict[str, str],
                                add_rect_fn: Callable | None = None,
                                add_oval_fn: Callable | None = None,
                                add_text_fn: Callable | None = None,
                                apply_glow_fn: Callable | None = None) -> None:
        method = getattr(self, f"_title_{style}", self._title_accent_bar)
        ctx = {
            "add_rect": add_rect_fn,
            "add_oval": add_oval_fn,
            "add_text": add_text_fn,
            "apply_glow": apply_glow_fn,
        }
        method(slide, title_x, title_y, title_w, colors, ctx)

    def _title_accent_bar(self, slide, x, y, w, colors, ctx):
        if ctx.get("add_rect"):
            ctx["add_rect"](slide, x, y + 0.55, w * 0.3, 0.04,
                            fill_hex=colors.get("primary", "#2563EB"))

    def _title_neon_lines(self, slide, x, y, w, colors, ctx):
        if ctx.get("add_rect"):
            shape = ctx["add_rect"](slide, x, y + 0.55, w * 0.3, 0.03,
                                     fill_hex=colors.get("primary", "#2563EB"))
            if ctx.get("apply_glow") and shape is not None:
                ctx["apply_glow"](shape, radius_pt=6,
                                  color=colors.get("primary", "#2563EB"),
                                  alpha_pct=30)

    def _title_gold_trim(self, slide, x, y, w, colors, ctx):
        if ctx.get("add_rect"):
            ctx["add_rect"](slide, x, y + 0.55, w * 0.3, 0.04,
                            fill_hex="#D4A853")

    def _title_minimal_dots(self, slide, x, y, w, colors, ctx):
        if ctx.get("add_oval"):
            ctx["add_oval"](slide, x - 0.15, y + 0.1, 0.08, 0.08,
                            fill_hex=colors.get("primary", "#2563EB"))

    def _title_diamond_bullets(self, slide, x, y, w, colors, ctx):
        if ctx.get("add_text"):
            ctx["add_text"](slide, "\u25C6", x - 0.2, y, 0.2, 0.4,
                            size=16, color_role="primary", bold=True)

    def _title_gradient_bar(self, slide, x, y, w, colors, ctx):
        if ctx.get("add_rect"):
            ctx["add_rect"](slide, x, y + 0.55, w * 0.3, 0.04,
                            fill_hex=colors.get("primary", "#2563EB"), gradient=True)

    def _title_circle_accent(self, slide, x, y, w, colors, ctx):
        if ctx.get("add_oval"):
            ctx["add_oval"](slide, x - 0.3, y - 0.05, 0.3, 0.3,
                            fill_hex=colors.get("primary", "#2563EB"))

    def _title_sidebar_nav(self, slide, x, y, w, colors, ctx):
        if ctx.get("add_rect"):
            ctx["add_rect"](slide, 0, 0, 0.06, 7.5,
                            fill_hex=colors.get("primary", "#2563EB"))

    def _title_no_decoration(self, slide, x, y, w, colors, ctx):
        pass

    def _title_full_bleed_overlay(self, slide, x, y, w, colors, ctx):
        pass
