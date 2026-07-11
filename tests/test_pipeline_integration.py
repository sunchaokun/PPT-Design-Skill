"""Integration tests for EnterprisePipeline full flow."""

from __future__ import annotations

import json
import os
from pathlib import Path

import pytest
from PIL import Image
from pptx import Presentation


def _make_template(path: Path, num_slides: int = 3) -> Path:
    prs = Presentation()
    for i in range(num_slides):
        layout = prs.slide_layouts[-1]
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = f"Slide {i + 1}"
    prs.save(str(path))
    return path


def _make_png(path: Path) -> Path:
    img = Image.new("RGB", (100, 50), color="blue")
    img.save(str(path))
    return path


def _setup_project(tmp_path: Path, with_content: bool = False) -> Path:
    project = tmp_path / "myproject"
    project.mkdir()
    _make_template(project / "template.pptx", 3)
    _make_png(project / "logo.png")
    brand = {"colors": {"primary": "#1A3C6E", "accent": "#C9A227"}, "fonts": {"heading": "Georgia"}}
    (project / "brand.json").write_text(json.dumps(brand), encoding="utf-8")
    if with_content:
        content = {
            "meta": {"title": "My Pitch", "subtitle": "2026"},
            "slides": [
                {"goal": "hook", "title": "Welcome", "image": "hero.png"},
                {"goal": "problem", "title": "The Problem", "bullets": ["Pain 1", "Pain 2"]},
            ],
        }
        (project / "content.json").write_text(json.dumps(content), encoding="utf-8")
        _make_png(project / "hero.png")
    return project


class TestPipelineFullFlow:

    def test_full_render_with_content_json(self, tmp_path):
        """Full pipeline with content.json should produce valid PPTX."""
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        project = _setup_project(tmp_path, with_content=True)
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="My Pitch",
            project_dir=str(project),
            business_mode="pitch",
        )
        assert result["pipeline"] == "enterprise"
        assert result["num_slides"] == 2
        assert os.path.isfile(result["output_path"])
        prs = Presentation(result["output_path"])
        assert len(prs.slides) == 2

    def test_full_render_without_content_json(self, tmp_path):
        """Full pipeline without content.json uses auto story plan."""
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        project = _setup_project(tmp_path, with_content=False)
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="AI Startup",
            project_dir=str(project),
        )
        assert result["pipeline"] == "enterprise"
        assert os.path.isfile(result["output_path"])

    def test_full_render_with_review(self, tmp_path):
        """Review mode should produce proposal.json, not PPTX."""
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        project = _setup_project(tmp_path, with_content=True)
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="My Pitch",
            project_dir=str(project),
            review=True,
        )
        assert result["review"] is True
        assert os.path.isfile(result["proposal_path"])
        proposal = json.loads(Path(result["proposal_path"]).read_text(encoding="utf-8"))
        assert "pages" in proposal

    def test_full_render_with_page_revision(self, tmp_path):
        """Pipeline with --pages should use PageRevisionEngine."""
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        project = _setup_project(tmp_path, with_content=False)
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="Revised",
            project_dir=str(project),
            pages="-2",
        )
        assert result["mode"] == "page_revision"
        assert result["num_slides"] == 2
        assert os.path.isfile(result["output_path"])

    def test_full_render_version_meta(self, tmp_path):
        """Full render should write meta.json with slides[]."""
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        project = _setup_project(tmp_path, with_content=True)
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="My Pitch",
            project_dir=str(project),
        )
        vdir = os.path.join(str(project), "output", f"v{result['version']}")
        meta_path = os.path.join(vdir, "meta.json")
        assert os.path.isfile(meta_path)
        meta = json.loads(Path(meta_path).read_text(encoding="utf-8"))
        assert "slides" in meta
        assert len(meta["slides"]) == 2

    def test_full_render_with_logo(self, tmp_path):
        """Pipeline should insert logo on slides when logo exists."""
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        project = _setup_project(tmp_path, with_content=True)
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="My Pitch",
            project_dir=str(project),
        )
        prs = Presentation(result["output_path"])
        has_picture = any(
            s for s in prs.slides
            for sh in s.shapes
            if sh.shape_type == 13
        )
        assert has_picture

    def test_dry_run_still_works(self, tmp_path):
        """Dry run should not produce PPTX."""
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        project = _setup_project(tmp_path, with_content=True)
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="My Pitch",
            project_dir=str(project),
            dry_run=True,
        )
        assert result["dry_run"] is True
        assert "output_path" not in result
