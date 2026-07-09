"""Phase 4: PPT Renderer — python-pptx direct .pptx generation."""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from typing import Any

from ppt_pro_max.decider.design_decider import PageDesign
from ppt_pro_max.content.content_generator import PageContent
from ppt_pro_max.renderer.theme_mapper import ThemeMapper
from ppt_pro_max.renderer.layout_registry import LayoutRegistry, SLIDE_WIDTH, SLIDE_HEIGHT
from ppt_pro_max.renderer.chart_builder import ChartBuilder
from ppt_pro_max.renderer.effects import Effects
from ppt_pro_max.qa.qa_gates import QAGates

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.oxml.ns import qn
    from lxml import etree

    _PPTX_AVAILABLE = True
except ImportError:
    _PPTX_AVAILABLE = False


class PPTRenderer:
    def __init__(self):
        if not _PPTX_AVAILABLE:
            raise RuntimeError("python-pptx is required. Install with: pip install python-pptx")
        self.theme_mapper = ThemeMapper()
        self.layout_registry = LayoutRegistry()
        self.chart_builder = ChartBuilder()
        self.effects = Effects()
        self.qa = QAGates()

    def render(
        self,
        page_designs: list[PageDesign],
        page_contents: list[PageContent],
        output_path: str | None = None,
        fetch_images: bool = False,
        theme_name: str | None = None,
    ) -> dict[str, Any]:
        if not page_designs:
            raise ValueError("No page designs to render")

        design_system = self._extract_design_system(page_designs[0])
        ppt_theme = self.theme_mapper.map(design_system, theme_name=theme_name)

        prs = Presentation()
        prs.slide_width = Inches(SLIDE_WIDTH)
        prs.slide_height = Inches(SLIDE_HEIGHT)

        self.theme_mapper.apply_theme(prs, ppt_theme)

        for design, content in zip(page_designs, page_contents):
            self._render_slide(prs, design, content, ppt_theme)

        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"presentation_{timestamp}.pptx"

        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        prs.save(output_path)

        qa_result = self.qa.check(output_path, page_designs, page_contents)

        return {
            "output_path": os.path.abspath(output_path),
            "page_count": len(page_designs),
            "strategy": "generated",
            "theme": ppt_theme.get("name", "default"),
            "qa": qa_result,
        }

    def _extract_design_system(self, design: PageDesign) -> dict[str, Any]:
        return {
            "colors": {
                "primary": "#2563EB",
                "on-primary": "#FFFFFF",
                "secondary": "#64748B",
                "accent": "#F97316",
                "background": "#FFFFFF",
                "foreground": "#1E293B",
                "muted": "#F1F5F9",
                "muted-foreground": "#94A3B8",
                "border": "#E2E8F0",
                "destructive": "#EF4444",
            },
            "typography": {
                "heading": "Inter",
                "body": "Inter",
            },
        }

    def _render_slide(self, prs: Presentation, design: PageDesign, content: PageContent, theme: dict) -> None:
        layout_def = self.layout_registry.get_layout(design.layout)
        if layout_def is None:
            layout_def = self.layout_registry.get_layout("content-with-title")

        slide_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)

        self._apply_background(slide, design, theme)

        for ph_name, ph_def in layout_def.get("placeholders", {}).items():
            self._render_placeholder(slide, ph_name, ph_def, content, theme)

        self._apply_transition(slide, design.transition)

    def _apply_background(self, slide, design: PageDesign, theme: dict) -> None:
        bg_color = theme.get("colors", {}).get("background", "#FFFFFF")
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(bg_color.lstrip("#"))

    def _render_placeholder(self, slide, ph_name: str, ph_def: dict, content: PageContent, theme: dict) -> None:
        ph_type = ph_def.get("type", "text")
        if ph_type == "image":
            self._render_image_placeholder(slide, ph_def, content)
            return

        text = self._get_placeholder_text(ph_name, content)
        if text is None:
            return

        left = Inches(ph_def["x"])
        top = Inches(ph_def["y"])
        width = Inches(ph_def["width"])
        height = Inches(ph_def["height"])

        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = text

        font_size = ph_def.get("font_size", 16)
        p.font.size = Pt(font_size)
        p.font.bold = ph_def.get("font_weight") in ("bold", "semibold", "600", "700")

        alignment = ph_def.get("alignment", "left")
        align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
        p.alignment = align_map.get(alignment, PP_ALIGN.LEFT)

        color_role = ph_def.get("color_role", "foreground")
        color_hex = theme.get("colors", {}).get(color_role, "#1E293B")
        p.font.color.rgb = RGBColor.from_string(color_hex.lstrip("#"))

        font_name = ph_def.get("font_role", "body")
        font_family = theme.get("typography", {}).get(font_name, "Inter")
        p.font.name = font_family

    def _get_placeholder_text(self, ph_name: str, content: PageContent) -> str | None:
        if ph_name == "title":
            return content.title
        if ph_name in ("subtitle", "tagline"):
            return content.subtitle
        if ph_name == "body":
            return "\n".join(f"• {b}" for b in content.bullets) if content.bullets else None
        if ph_name == "section_number":
            return str(content.position)
        if ph_name == "section_title":
            return content.title
        if ph_name == "quote_text":
            return content.quote.get("text", "") if content.quote else None
        if ph_name == "quote_author":
            return content.quote.get("author", "") if content.quote else None
        return None

    def _render_image_placeholder(self, slide, ph_def: dict, content: PageContent) -> None:
        left = Inches(ph_def["x"])
        top = Inches(ph_def["y"])
        width = Inches(ph_def["width"])
        height = Inches(ph_def["height"])

        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left, top, width, height,
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0xF1, 0xF5, 0xF9)
        shape.line.color.rgb = RGBColor(0xE2, 0xE8, 0xF0)

        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = "[在此插入图片]"
        p.font.size = Pt(12)
        p.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
        p.alignment = PP_ALIGN.CENTER

    def _apply_transition(self, slide, transition_name: str) -> None:
        if not _PPTX_AVAILABLE:
            return
        try:
            transition_xml_map = {
                "fade": "<p:fade/>",
                "push": '<p:push dir="l"/>',
                "wipe": '<p:wipe dir="d"/>',
                "cut": "<p:cut/>",
            }
            child_xml = transition_xml_map.get(transition_name, "<p:fade/>")
            transition = etree.SubElement(slide._element, qn("p:transition"))
            transition.set("spd", "med")
            child = etree.fromstring(child_xml)
            transition.append(child)
        except Exception:
            pass
