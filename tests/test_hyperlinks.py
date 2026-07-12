"""Tests for Hyperlinks feature (v0.4 I-B)."""

from __future__ import annotations

import json
import os
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Inches


class TestHyperlinks:

    def test_url_link_on_bullet(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        pipeline._populate_slide(
            slide,
            {
                "title": "Links",
                "bullets": ["Visit us", "Contact us"],
                "links": [{"bullet_index": 0, "url": "https://example.com"}],
            },
            prs,
        )
        from pptx.enum.shapes import PP_PLACEHOLDER
        for ph in slide.placeholders:
            ph_type = ph.placeholder_format.type
            if ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT) or ph.placeholder_format.idx == 1:
                paras = list(ph.text_frame.paragraphs)
                assert len(paras) >= 1
                for run in paras[0].runs:
                    if run.hyperlink.address:
                        assert run.hyperlink.address == "https://example.com"
                        return
        pytest.fail("No hyperlink found on bullet")

    def test_mailto_link(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        pipeline._populate_slide(
            slide,
            {
                "title": "Contact",
                "bullets": ["Email us"],
                "links": [{"bullet_index": 0, "url": "mailto:sales@example.com"}],
            },
            prs,
        )
        from pptx.enum.shapes import PP_PLACEHOLDER
        for ph in slide.placeholders:
            ph_type = ph.placeholder_format.type
            if ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT) or ph.placeholder_format.idx == 1:
                for run in list(ph.text_frame.paragraphs)[0].runs:
                    if run.hyperlink.address:
                        assert run.hyperlink.address == "mailto:sales@example.com"
                        return
        pytest.fail("No mailto hyperlink found")

    def test_standalone_link(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[-1])
        pipeline._populate_slide(
            slide,
            {
                "title": "CTA",
                "links": [{"text": "Download PDF", "url": "https://example.com/report.pdf", "position": "bottom_right"}],
            },
            prs,
        )
        link_found = False
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.hyperlink.address == "https://example.com/report.pdf":
                            link_found = True
        assert link_found

    def test_no_links_no_side_effects(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        pipeline._populate_slide(slide, {"title": "No Links"}, prs)
        hyperlink_runs = []
        for sh in slide.shapes:
            if sh.has_text_frame:
                for para in sh.text_frame.paragraphs:
                    for run in para.runs:
                        if run.hyperlink and run.hyperlink.address:
                            hyperlink_runs.append(run)
        assert len(hyperlink_runs) == 0

    def test_link_style_accent_color(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        pipeline._populate_slide(
            slide,
            {
                "title": "Styled",
                "bullets": ["Click here"],
                "links": [{"bullet_index": 0, "url": "https://example.com"}],
            },
            prs,
        )
        from pptx.enum.shapes import PP_PLACEHOLDER
        for ph in slide.placeholders:
            ph_type = ph.placeholder_format.type
            if ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT) or ph.placeholder_format.idx == 1:
                for run in list(ph.text_frame.paragraphs)[0].runs:
                    if run.hyperlink.address:
                        assert run.font.underline is True
                        return
        pytest.fail("No styled hyperlink found")

    def test_content_json_links_passthrough(self, tmp_path):
        from ppt_pro_max.enterprise.content_parser import load_enterprise_content
        content = {
            "meta": {"title": "Test"},
            "slides": [
                {
                    "goal": "cta",
                    "title": "Get Started",
                    "bullets": ["Visit website"],
                    "links": [{"bullet_index": 0, "url": "https://example.com"}],
                },
            ],
        }
        result = load_enterprise_content(content, str(tmp_path))
        assert result[0].get("links") is not None
        assert result[0]["links"][0]["url"] == "https://example.com"
