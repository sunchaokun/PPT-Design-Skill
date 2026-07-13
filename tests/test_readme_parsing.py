"""Tests for P4: ContentParser README.md parsing."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from ppt_pro_max.enterprise.content_parser import parse_readme


def _write_readme(path: Path, content: str) -> Path:
    path.write_text(content, encoding="utf-8")
    return path


class TestReadmeBasicParsing:

    def test_single_h1_becomes_hook_page(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# My Project\n\nA great project.")
        pages = parse_readme(str(readme), str(tmp_path))
        assert len(pages) >= 1
        assert pages[0]["goal"] == "hook"
        assert "My Project" in pages[0]["title"]

    def test_multiple_h1s_become_multiple_pages(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Page One\n\nText1\n\n# Page Two\n\nText2")
        pages = parse_readme(str(readme), str(tmp_path))
        assert len(pages) == 2

    def test_h2_becomes_bullets(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Features\n\n## Fast\nSpeed details\n\n## Secure\nSecurity details")
        pages = parse_readme(str(readme), str(tmp_path))
        assert len(pages) == 1
        bullets = pages[0].get("bullets") or []
        assert any("Fast" in b for b in bullets)

    def test_list_items_become_bullets(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Features\n\n- Feature A\n- Feature B\n- Feature C")
        pages = parse_readme(str(readme), str(tmp_path))
        bullets = pages[0].get("bullets") or []
        assert "Feature A" in bullets
        assert "Feature B" in bullets

    def test_code_block_becomes_code(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Quick Start\n\n```python\nprint('hello')\n```")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0].get("code") is not None

    def test_table_becomes_diagram(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Comparison\n\n| A | B |\n|---|---|\n| 1 | 2 |")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0].get("diagram_type") == "table" or pages[0].get("diagram_data") is not None

    def test_empty_readme_returns_empty(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages == []

    def test_no_headings_creates_single_content_page(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "Just some text without headings.")
        pages = parse_readme(str(readme), str(tmp_path))
        assert len(pages) == 1
        assert pages[0]["goal"] in ("content", "hook")


class TestReadmeGoalInference:

    def test_problem_keyword_infers_problem_goal(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Problem Statement\n\nThe current challenges")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0]["goal"] == "problem"

    def test_solution_keyword_infers_solution_goal(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Solution\n\nHow we solve it")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0]["goal"] == "solution"

    def test_features_keyword_infers_features_goal(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Features\n\nWhat we offer")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0]["goal"] == "features"

    def test_market_keyword_infers_content_goal(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Intro\n\nText\n\n# Overview\n\nAgenda\n\n# Market Size\n\nTAM analysis")
        pages = parse_readme(str(readme), str(tmp_path))
        if len(pages) >= 3:
            assert pages[2]["goal"] == "content"

    def test_architecture_keyword_infers_data_goal(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Architecture\n\nSystem design")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0]["goal"] == "data"

    def test_code_keyword_infers_code_goal(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Quick Start\n\nGetting started")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0]["goal"] == "code"

    def test_contact_keyword_infers_cta_goal(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Contact Us\n\nGet in touch")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0]["goal"] == "cta"

    def test_first_page_always_hook(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Introduction\n\n## Overview\n- Point 1\n\n# Features\n\n- Feature A")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0]["goal"] == "hook"

    def test_second_page_overview(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Intro\n\nText\n\n# Overview\n\nAgenda items")
        pages = parse_readme(str(readme), str(tmp_path))
        if len(pages) >= 2:
            assert pages[1]["goal"] == "overview"


class TestReadmeChineseKeywords:

    def test_chinese_problem_keyword(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# 痛点分析\n\n当前面临的挑战")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0]["goal"] == "problem"

    def test_chinese_solution_keyword(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# 解决方案\n\n我们的方法")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0]["goal"] == "solution"

    def test_chinese_features_keyword(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# 功能特性\n\n核心能力")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0]["goal"] == "features"

    def test_chinese_contact_keyword(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# 联系我们\n\n联系方式")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0]["goal"] == "cta"


class TestReadmeImageReferences:

    def test_image_ref_assigned_to_page(self, tmp_path):
        img = tmp_path / "diagram.png"
        img.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 100)
        readme = _write_readme(tmp_path / "README.md", f"# Architecture\n\n![Diagram]({img.name})")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0].get("image") is not None

    def test_nonexistent_image_ignored(self, tmp_path):
        readme = _write_readme(tmp_path / "README.md", "# Architecture\n\n![Missing](nonexistent.png)")
        pages = parse_readme(str(readme), str(tmp_path))
        assert pages[0].get("image") is None


class TestReadmeFileNotFound:

    def test_nonexistent_file_returns_empty(self, tmp_path):
        pages = parse_readme(str(tmp_path / "nonexistent.md"), str(tmp_path))
        assert pages == []
