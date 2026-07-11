"""Tests for ReviewGate and content.json parser."""

from __future__ import annotations

import json
from pathlib import Path

import pytest


# ============================================================
# ReviewGate
# ============================================================

class TestReviewGate:

    def test_generate_proposal(self, tmp_path):
        """ReviewGate should generate a proposal dict."""
        from ppt_pro_max.enterprise.review_gate import ReviewGate
        gate = ReviewGate()
        proposal = gate.generate_proposal(
            project_dir=str(tmp_path),
            brand_spec={"source": "brand_json", "colors": {"primary": "#1A3C6E"}},
            story_plan={"strategy": "YC Seed Deck", "pages": [{"goal": "hook"}]},
            page_designs=[{"goal": "hook", "layout": "Title Slide", "title": "Hello"}],
            assets={"template": True, "logo": True},
        )
        assert "brand_spec" in proposal
        assert "story_plan" in proposal
        assert "pages" in proposal

    def test_write_proposal_json(self, tmp_path):
        """ReviewGate should write proposal to JSON file."""
        from ppt_pro_max.enterprise.review_gate import ReviewGate
        gate = ReviewGate()
        proposal = {"test": True}
        out_path = str(tmp_path / "proposal.json")
        gate.write_proposal(proposal, out_path)
        loaded = json.loads(Path(out_path).read_text(encoding="utf-8"))
        assert loaded["test"] is True

    def test_read_proposal_json(self, tmp_path):
        """ReviewGate should read proposal from JSON file."""
        from ppt_pro_max.enterprise.review_gate import ReviewGate
        gate = ReviewGate()
        proposal = {"version": 2, "pages": [{"goal": "hook"}]}
        out_path = str(tmp_path / "proposal.json")
        Path(out_path).write_text(json.dumps(proposal), encoding="utf-8")
        loaded = gate.read_proposal(out_path)
        assert loaded["version"] == 2

    def test_format_cli_display(self):
        """ReviewGate should format proposal for CLI display."""
        from ppt_pro_max.enterprise.review_gate import ReviewGate
        gate = ReviewGate()
        text = gate.format_cli({
            "brand_spec": {"source": "brand_json", "colors": {"primary": "#1A3C6E"}},
            "assets": {"template": True, "logo": False, "brand_json": True, "content_json": False, "image_pool_count": 3},
            "pages": [{"goal": "hook", "layout": "Title Slide", "title": "Hello"}],
        })
        assert "primary" in text
        assert "hook" in text


# ============================================================
# Content.json parser
# ============================================================

class TestEnterpriseContentParser:

    def test_parse_content_json(self, tmp_path):
        """_load_enterprise_content should parse slides[] to list of dicts."""
        from ppt_pro_max.enterprise.content_parser import load_enterprise_content
        data = {
            "meta": {"title": "Test", "subtitle": "Sub", "author": "Me"},
            "slides": [
                {"goal": "hook", "title": "Hello", "image": "hero.png"},
                {"goal": "problem", "title": "Pain", "bullets": ["a", "b"]},
            ],
        }
        result = load_enterprise_content(data, str(tmp_path))
        assert len(result) == 2
        assert result[0]["goal"] == "hook"
        assert result[0]["image"] == str(tmp_path / "hero.png")
        assert result[1]["bullets"] == ["a", "b"]

    def test_parse_content_json_with_diagram(self, tmp_path):
        """_load_enterprise_content should extract diagram_type and diagram_data."""
        from ppt_pro_max.enterprise.content_parser import load_enterprise_content
        data = {
            "slides": [
                {
                    "goal": "process",
                    "title": "Flow",
                    "diagram": {
                        "type": "flowchart",
                        "nodes": [{"id": 1, "label": "Start"}],
                        "edges": [],
                    },
                },
            ],
        }
        result = load_enterprise_content(data, str(tmp_path))
        assert result[0]["diagram_type"] == "flowchart"
        assert result[0]["diagram_data"]["type"] == "flowchart"

    def test_parse_content_json_with_cards(self, tmp_path):
        """_load_enterprise_content should parse cards with image paths."""
        from ppt_pro_max.enterprise.content_parser import load_enterprise_content
        data = {
            "slides": [
                {
                    "goal": "features",
                    "title": "Features",
                    "cards": [
                        {"title": "A", "body": "desc", "image": "dash.png"},
                    ],
                },
            ],
        }
        result = load_enterprise_content(data, str(tmp_path))
        assert len(result[0]["cards"]) == 1
        assert result[0]["cards"][0]["image"] == str(tmp_path / "dash.png")

    def test_parse_empty_content_json(self):
        """Empty slides[] should return empty list."""
        from ppt_pro_max.enterprise.content_parser import load_enterprise_content
        result = load_enterprise_content({"slides": []}, "/tmp")
        assert result == []

    def test_parse_content_json_meta_fallback(self, tmp_path):
        """First slide without title should use meta.title as fallback."""
        from ppt_pro_max.enterprise.content_parser import load_enterprise_content
        data = {
            "meta": {"title": "Main Title"},
            "slides": [
                {"goal": "hook"},
            ],
        }
        result = load_enterprise_content(data, str(tmp_path))
        assert result[0]["title"] == "Main Title"
