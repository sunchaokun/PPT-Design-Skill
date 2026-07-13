"""Tests for P2: Pipeline unified PrecisionRenderer path.

Covers:
  1. Pipeline with template uses PrecisionRenderer (render_mode=precision)
  2. Pipeline without template uses PrecisionRenderer (render_mode=precision)
  3. Deleted paths (_populate_slide, _render_with_ppt_renderer, _populate_slide_with_layout) are gone
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Inches


def _make_project(tmp_path: Path, with_template: bool = True) -> str:
    project = tmp_path / "project"
    project.mkdir(exist_ok=True)
    if with_template:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        prs.save(str(project / "template.pptx"))
    content = {
        "slides": [
            {"goal": "hook", "title": "Cover"},
            {"goal": "content", "title": "Body", "bullets": ["A", "B"]},
            {"goal": "cta", "title": "End"},
        ]
    }
    (project / "content.json").write_text(json.dumps(content), encoding="utf-8")
    return str(project)


class TestPipelineUnifiedPrecision:
    """Pipeline must always use PrecisionRenderer."""

    def test_with_template_returns_precision_mode(self, tmp_path):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        project_dir = _make_project(tmp_path, with_template=True)
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="Test",
            project_dir=project_dir,
            density=4,
        )
        assert result.get("render_mode") == "precision", (
            f"Expected render_mode='precision', got '{result.get('render_mode')}'"
        )

    def test_without_template_returns_precision_mode(self, tmp_path):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        project_dir = _make_project(tmp_path, with_template=False)
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="Test",
            project_dir=project_dir,
            density=4,
        )
        assert result.get("render_mode") == "precision", (
            f"Expected render_mode='precision', got '{result.get('render_mode')}'"
        )

    def test_output_file_created(self, tmp_path):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        project_dir = _make_project(tmp_path, with_template=True)
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="Test",
            project_dir=project_dir,
            density=4,
        )
        output_path = result.get("output_path")
        assert output_path is not None
        assert Path(output_path).exists()


class TestDeletedPaths:
    """Old rendering paths must be removed from Pipeline."""

    def test_no_populate_slide_method(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        assert not hasattr(EnterprisePipeline, "_populate_slide"), (
            "_populate_slide should be deleted"
        )

    def test_no_render_with_ppt_renderer_method(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        assert not hasattr(EnterprisePipeline, "_render_with_ppt_renderer"), (
            "_render_with_ppt_renderer should be deleted"
        )

    def test_no_populate_slide_with_layout_method(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        assert not hasattr(EnterprisePipeline, "_populate_slide_with_layout"), (
            "_populate_slide_with_layout should be deleted"
        )
