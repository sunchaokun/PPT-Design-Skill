"""Tests for QA gates."""

from ppt_pro_max.qa.qa_gates import QAGates
from ppt_pro_max.decider.design_decider import PageDesign
from ppt_pro_max.content.content_generator import PageContent


def test_qa_passes_good_ppt():
    designs = [
        PageDesign(position=1, goal="hook", emotion="curiosity", layout="title-slide"),
        PageDesign(position=2, goal="problem", emotion="frustration", layout="content-with-title"),
        PageDesign(position=3, goal="cta", emotion="urgency", layout="cta-closing"),
    ]
    contents = [
        PageContent(position=1, goal="hook", title="Welcome"),
        PageContent(position=2, goal="problem", title="The Problem"),
        PageContent(position=3, goal="cta", title="Take Action"),
    ]
    qa = QAGates()
    result = qa.check("test.pptx", designs, contents)
    assert result["passed"] is True


def test_qa_catches_missing_titles():
    designs = [PageDesign(position=1, goal="hook", emotion="curiosity", layout="title-slide")]
    contents = [PageContent(position=1, goal="hook", title="")]
    qa = QAGates()
    result = qa.check("test.pptx", designs, contents)
    title_check = next(c for c in result["checks"] if c["name"] == "title_per_page")
    assert title_check["passed"] is False


def test_qa_catches_too_few_pages():
    designs = [PageDesign(position=1, goal="hook", emotion="curiosity", layout="title-slide")]
    contents = [PageContent(position=1, goal="hook", title="Test")]
    qa = QAGates()
    result = qa.check("test.pptx", designs, contents)
    count_check = next(c for c in result["checks"] if c["name"] == "page_count")
    assert count_check["passed"] is False
