"""Tests for PPT renderer — end-to-end .pptx generation."""

import os
import tempfile

from ppt_pro_max.renderer.ppt_renderer import PPTRenderer
from ppt_pro_max.renderer.layout_registry import MASTER_LAYOUTS
from ppt_pro_max.decider.design_decider import PageDesign
from ppt_pro_max.content.content_generator import PageContent


def _make_design(position=1, goal="hook", emotion="curiosity", layout="title-slide", chart_type=None) -> PageDesign:
    return PageDesign(
        position=position,
        goal=goal,
        emotion=emotion,
        layout=layout,
        copy_formula="AIDA",
        chart_type=chart_type,
    )


def _make_content(position=1, goal="hook", title="Test Title", bullets=None, metrics=None, chart_data=None) -> PageContent:
    return PageContent(
        position=position,
        goal=goal,
        title=title,
        subtitle="Test Subtitle",
        bullets=bullets or ["Point 1", "Point 2"],
        metrics=metrics,
        chart_data=chart_data,
    )


def test_render_minimal_pptx():
    with tempfile.TemporaryDirectory() as tmpdir:
        output = os.path.join(tmpdir, "test.pptx")
        renderer = PPTRenderer()

        designs = [_make_design(1, "hook", layout="title-slide"), _make_design(2, "cta", layout="cta-closing"), _make_design(3, "problem", layout="content-with-title")]
        contents = [_make_content(1, "hook", "Welcome"), _make_content(2, "cta", "Take Action"), _make_content(3, "problem", "The Problem")]

        result = renderer.render(designs, contents, output_path=output)
        assert os.path.exists(result["output_path"])
        assert result["page_count"] == 3


def test_render_with_theme():
    with tempfile.TemporaryDirectory() as tmpdir:
        output = os.path.join(tmpdir, "test_dark.pptx")
        renderer = PPTRenderer()

        designs = [_make_design(1, "hook"), _make_design(2, "cta", layout="cta-closing"), _make_design(3, "problem")]
        contents = [_make_content(1, "hook", "Welcome"), _make_content(2, "cta", "Go!"), _make_content(3, "problem", "Problem")]

        result = renderer.render(designs, contents, output_path=output, theme_name="dark-tech")
        assert result["theme"] == "Dark Tech"
        assert os.path.exists(result["output_path"])


def test_render_with_metrics():
    with tempfile.TemporaryDirectory() as tmpdir:
        output = os.path.join(tmpdir, "test_metrics.pptx")
        renderer = PPTRenderer()

        metrics = [
            {"label": "Users", "value": "10K+"},
            {"label": "Retention", "value": "95%"},
            {"label": "Growth", "value": "3x"},
            {"label": "ARR", "value": "$2M"},
        ]
        designs = [
            _make_design(1, "hook", layout="title-slide"),
            _make_design(2, "traction", layout="four-metrics", chart_type="Line Chart"),
            _make_design(3, "cta", layout="cta-closing"),
        ]
        contents = [
            _make_content(1, "hook", "Welcome"),
            _make_content(2, "traction", "Traction", metrics=metrics, chart_data={"type": "Line Chart", "data": {"labels": ["Q1", "Q2"], "values": [10, 25]}}),
            _make_content(3, "cta", "Go!"),
        ]

        result = renderer.render(designs, contents, output_path=output)
        assert os.path.exists(result["output_path"])


def test_render_with_cards():
    with tempfile.TemporaryDirectory() as tmpdir:
        output = os.path.join(tmpdir, "test_cards.pptx")
        renderer = PPTRenderer()

        designs = [
            _make_design(1, "hook", layout="title-slide"),
            _make_design(2, "features", layout="three-column-cards"),
            _make_design(3, "cta", layout="cta-closing"),
        ]
        contents = [
            _make_content(1, "hook", "Welcome"),
            _make_content(2, "features", "Features", bullets=["Fast", "Secure", "Easy"]),
            _make_content(3, "cta", "Go!"),
        ]

        result = renderer.render(designs, contents, output_path=output)
        assert os.path.exists(result["output_path"])


def test_render_all_themes():
    with tempfile.TemporaryDirectory() as tmpdir:
        renderer = PPTRenderer()
        designs = [_make_design(1, "hook"), _make_design(2, "cta", layout="cta-closing"), _make_design(3, "problem")]
        contents = [_make_content(1, "hook", "Hi"), _make_content(2, "cta", "Go"), _make_content(3, "problem", "Fix")]

        for theme in ["professional", "dark-tech", "warm-elegant", "vibrant-startup", "nature-calm"]:
            output = os.path.join(tmpdir, f"test_{theme}.pptx")
            result = renderer.render(designs, contents, output_path=output, theme_name=theme)
            assert os.path.exists(result["output_path"]), f"Failed for theme {theme}"
