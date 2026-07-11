"""Tests for _populate_slide content rendering."""

from __future__ import annotations

import os
from pathlib import Path

import pytest
from PIL import Image
from pptx import Presentation
from pptx.util import Inches


def _make_png(path: Path, w: int = 200, h: int = 150) -> Path:
    img = Image.new("RGB", (w, h), color="green")
    img.save(str(path))
    return path


class TestPopulateSlide:

    def test_title_populated(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        pipeline._populate_slide(slide, {"title": "Hello World"}, prs)
        assert slide.shapes.title.text == "Hello World"

    def test_subtitle_populated(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        pipeline._populate_slide(slide, {"title": "Hi", "subtitle": "Sub text"}, prs)
        assert slide.shapes.title.text == "Hi"
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 2:
                assert ph.text == "Sub text"

    def test_bullets_populated(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        pipeline._populate_slide(slide, {"title": "Points", "bullets": ["First", "Second", "Third"]}, prs)
        assert slide.shapes.title.text == "Points"
        body_found = False
        for ph in slide.placeholders:
            ph_type = str(ph.placeholder_format.type)
            if "BODY" in ph_type or ph.placeholder_format.idx == 1:
                assert "First" in ph.text
                assert "Second" in ph.text
                body_found = True
                break
        assert body_found

    def test_bullets_with_existing_markers(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        pipeline._populate_slide(slide, {"title": "T", "bullets": ["• Already marked", "— Dash marked"]}, prs)
        for ph in slide.placeholders:
            if "BODY" in str(ph.placeholder_format.type) or ph.placeholder_format.idx == 1:
                assert "• Already marked" in ph.text
                assert "— Dash marked" in ph.text

    def test_image_inserted(self, tmp_path):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        img_path = _make_png(tmp_path / "content.png")
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[-1])
        pipeline._populate_slide(slide, {"title": "With Image", "image": str(img_path)}, prs)
        pics = [s for s in slide.shapes if s.shape_type == 13]
        assert len(pics) == 1

    def test_image_nonexistent_no_crash(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[-1])
        pipeline._populate_slide(slide, {"title": "No Image", "image": "/nonexistent/img.png"}, prs)
        pics = [s for s in slide.shapes if s.shape_type == 13]
        assert len(pics) == 0

    def test_empty_design_no_crash(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[-1])
        pipeline._populate_slide(slide, {}, prs)

    def test_full_pipeline_with_bullets_and_image(self, tmp_path):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        import json
        project = tmp_path / "fullproject"
        project.mkdir()
        prs = Presentation()
        for _ in range(2):
            prs.slides.add_slide(prs.slide_layouts[-1])
        prs.save(str(project / "template.pptx"))
        _make_png(project / "logo.png")
        _make_png(project / "hero.png")
        (project / "brand.json").write_text(json.dumps({"colors": {"primary": "#1A3C6E"}}), encoding="utf-8")
        content = {
            "meta": {"title": "Full Test"},
            "slides": [
                {"goal": "hook", "title": "Welcome", "image": "hero.png"},
                {"goal": "problem", "title": "Pain Points", "bullets": ["Slow", "Expensive", "Complex"]},
            ],
        }
        (project / "content.json").write_text(json.dumps(content), encoding="utf-8")
        pipeline = EnterprisePipeline()
        result = pipeline.run(query="Full Test", project_dir=str(project))
        assert result["num_slides"] == 2
        prs_out = Presentation(result["output_path"])
        assert prs_out.slides[0].shapes.title.text == "Welcome"
        assert prs_out.slides[1].shapes.title.text == "Pain Points"
