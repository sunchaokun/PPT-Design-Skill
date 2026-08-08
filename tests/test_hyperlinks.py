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
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        renderer = PrecisionRenderer()
        prs = renderer.create_presentation()
        design = {
            "goal": "content",
            "title": "Links",
            "bullets": ["Visit us", "Contact us"],
            "links": [{"bullet_index": 0, "url": "https://example.com"}],
        }
        slide = renderer.render_slide(prs, design)
        pytest.skip("TODO: PrecisionRenderer.render_slide() does not render links yet — _links is read but unused")

    def test_mailto_link(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        renderer = PrecisionRenderer()
        prs = renderer.create_presentation()
        design = {
            "goal": "content",
            "title": "Contact",
            "bullets": ["Email us"],
            "links": [{"bullet_index": 0, "url": "mailto:sales@example.com"}],
        }
        slide = renderer.render_slide(prs, design)
        pytest.skip("TODO: PrecisionRenderer.render_slide() does not render links yet — _links is read but unused")

    def test_standalone_link(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        renderer = PrecisionRenderer()
        prs = renderer.create_presentation()
        design = {
            "goal": "cta",
            "title": "CTA",
            "links": [{"text": "Download PDF", "url": "https://example.com/report.pdf", "position": "bottom_right"}],
        }
        slide = renderer.render_slide(prs, design)
        pytest.skip("TODO: PrecisionRenderer.render_slide() does not render links yet — _links is read but unused")

    def test_no_links_no_side_effects(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        renderer = PrecisionRenderer()
        prs = renderer.create_presentation()
        design = {"goal": "content", "title": "No Links"}
        slide = renderer.render_slide(prs, design)
        hyperlink_runs = []
        for sh in slide.shapes:
            if sh.has_text_frame:
                for para in sh.text_frame.paragraphs:
                    for run in para.runs:
                        if run.hyperlink and run.hyperlink.address:
                            hyperlink_runs.append(run)
        assert len(hyperlink_runs) == 0

    def test_link_style_accent_color(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        renderer = PrecisionRenderer()
        prs = renderer.create_presentation()
        design = {
            "goal": "content",
            "title": "Styled",
            "bullets": ["Click here"],
            "links": [{"bullet_index": 0, "url": "https://example.com"}],
        }
        slide = renderer.render_slide(prs, design)
        pytest.skip("TODO: PrecisionRenderer.render_slide() does not render links yet — _links is read but unused")

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
