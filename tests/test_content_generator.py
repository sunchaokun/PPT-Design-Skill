"""Tests for content generator."""

from ppt_pro_max.content.content_generator import ContentGenerator
from ppt_pro_max.decider.design_decider import PageDesign


def _make_design(goal="hook", copy_formula="AIDA", chart_type=None, position=1) -> PageDesign:
    return PageDesign(
        position=position,
        goal=goal,
        emotion="curiosity",
        layout="title-slide",
        copy_formula=copy_formula,
        chart_type=chart_type,
    )


def test_generate_placeholder_content():
    gen = ContentGenerator()
    designs = [_make_design(goal="hook", copy_formula="AIDA")]
    contents = gen.generate(designs)
    assert len(contents) == 1
    assert "{Attention hook}" in contents[0].title


def test_generate_pas_content():
    gen = ContentGenerator()
    designs = [_make_design(goal="problem", copy_formula="PAS")]
    contents = gen.generate(designs)
    assert "面临" in contents[0].title


def test_generate_bullets_fallback():
    gen = ContentGenerator()
    designs = [_make_design(goal="solution", copy_formula="FAB")]
    contents = gen.generate(designs)
    assert len(contents[0].bullets) == 3
    assert contents[0].bullets[0].startswith("[solution")


def test_generate_metrics():
    gen = ContentGenerator()
    designs = [_make_design(goal="traction", copy_formula="Proof Stack")]
    contents = gen.generate(designs)
    assert contents[0].metrics is not None
    assert len(contents[0].metrics) == 4
    assert contents[0].metrics[0]["label"] == "用户数"


def test_generate_chart_data():
    gen = ContentGenerator()
    designs = [_make_design(goal="traction", copy_formula="Proof Stack", chart_type="Line Chart")]
    contents = gen.generate(designs)
    assert contents[0].chart_data is not None
    assert contents[0].chart_data["type"] == "Line Chart"
    assert "values" in contents[0].chart_data["data"]


def test_generate_no_chart_when_no_chart_type():
    gen = ContentGenerator()
    designs = [_make_design(goal="hook", copy_formula="AIDA", chart_type=None)]
    contents = gen.generate(designs)
    assert contents[0].chart_data is None


def test_user_content_from_file(tmp_path):
    import json
    content_file = tmp_path / "test_content.json"
    data = {"company": "TestCorp", "product": "AI Tool"}
    content_file.write_text(json.dumps(data), encoding="utf-8")

    gen = ContentGenerator()
    designs = [_make_design(goal="hook", copy_formula="AIDA")]
    contents = gen.generate(designs, content_file=str(content_file))
    assert "TestCorp" not in contents[0].title or "{Attention hook}" in contents[0].title


def test_dot_notation_template():
    gen = ContentGenerator()
    user_content = {"pain_points": [{"title": "Slow"}, {"title": "Expensive"}]}
    result = gen._fill_template("{pain_points.0.title}", "problem", user_content)
    assert result == "Slow"


def test_image_keywords():
    gen = ContentGenerator()
    designs = [_make_design(goal="problem", copy_formula="PAS")]
    contents = gen.generate(designs)
    assert "frustration" in contents[0].image_keywords
