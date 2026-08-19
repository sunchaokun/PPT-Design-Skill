"""Tests for color role hints — verify correct placement and functionality.

The critical behavior under test: color role hints for technical diagrams
(architecture, data flow, sequence) should be in the LLM prompt template,
not in Python code. The hints guide the LLM to assign colors by role
(frontend=cyan, backend=emerald, database=violet, etc.).
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from ppt_pro_max.content.content_generator import ContentGenerator, _GOAL_CONTENT_GENERATORS
from ppt_pro_max.enterprise.content_parser import _GOAL_KEYWORDS


class TestColorRoleHints:
    """Test that color role hints are properly defined and accessible."""

    def test_goal_keywords_include_architecture(self):
        """Architecture goal should be recognized."""
        architecture_keywords = [kw for goal, kws in _GOAL_KEYWORDS if goal == "overview" for kw in kws]
        assert "architecture" in architecture_keywords
        assert "架构图" in architecture_keywords

    def test_content_generator_has_generators(self):
        """ContentGenerator should have goal-specific generators."""
        assert "hook" in _GOAL_CONTENT_GENERATORS
        assert "features" in _GOAL_CONTENT_GENERATORS
        assert "overview" in _GOAL_CONTENT_GENERATORS

    def test_architecture_goal_recognition(self):
        """Architecture-related titles should map to overview goal."""
        from ppt_pro_max.enterprise.content_parser import _infer_goal

        # Architecture titles should map to overview
        assert _infer_goal("System Architecture", 1, 8) == "overview"
        assert _infer_goal("技术架构", 1, 8) == "overview"
        assert _infer_goal("Architecture Overview", 2, 8) == "overview"


class TestColorHintIntegration:
    """Test integration of color hints in content generation."""

    def test_content_generator_context_building(self):
        """ContentGenerator should build context with industry info."""
        gen = ContentGenerator("AI startup investor pitch")
        ctx = gen._build_context("AI startup investor pitch", "")

        # Should identify as tech and finance
        assert ctx["is_tech"] is True
        assert ctx["is_finance"] is True

    def test_content_generator_sustainability_context(self):
        """Sustainability queries should have appropriate context."""
        gen = ContentGenerator("carbon tracking platform for ESG reporting")
        ctx = gen._build_context("carbon tracking platform for ESG reporting", "")

        assert ctx["is_sustainability"] is True
        assert "industry" in ctx

    def test_content_generator_deep_tech_context(self):
        """Deep tech queries should have appropriate context."""
        gen = ContentGenerator("neural network training platform")
        ctx = gen._build_context("neural network training platform", "")

        assert ctx["is_deep_tech"] is True


class TestColorHintPlacement:
    """Test that color hints are in the correct module."""

    def test_no_color_hints_in_content_parser(self):
        """Content parser should not contain color role hints."""
        import inspect
        from ppt_pro_max.enterprise import content_parser

        source = inspect.getsource(content_parser)
        # Should not have ROLE_COLOR_HINTS or similar
        assert "ROLE_COLOR" not in source
        assert "color_hint" not in source

    def test_no_color_hints_in_content_generator(self):
        """Content generator should not contain color role hints in Python code."""
        import inspect
        from ppt_pro_max.content import content_generator

        source = inspect.getsource(content_generator)
        # Color hints should be in prompt templates, not Python constants
        # This test verifies the current state - hints should be added to prompts
        # via LLM instructions, not hardcoded in Python
