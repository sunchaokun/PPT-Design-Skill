"""Contract tests for ui_ux_adapter — the API surface consumed by the 4 caller modules.

Verifies every function consumers depend on returns the expected structure:
  theme_composer:  search_color, search_typography, search_style
  design_decider:  get_design_system, search_style, is_available
  story_planner:   search_design, search_landing, is_available
  slide_search:    search_reasoning, is_available
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from ppt_pro_max.adapters import ui_ux_adapter as ux


class TestAvailability:
    def test_is_available_true(self):
        # Bundled data => always available, no external install
        assert ux.is_available() is True

    def test_no_ux_env_var_required(self, monkeypatch):
        monkeypatch.delenv("UX_PRO_MAX_DIR", raising=False)
        assert ux.is_available() is True


class TestThemeComposerContract:
    """theme_composer calls: search_color(q,3), search_typography(q,3), search_style(q,3)."""

    def test_search_color_structure(self):
        r = ux.search_color("fintech", 2)
        assert isinstance(r, list)
        if r:
            row = r[0]
            # theme_composer reads these keys via _KEY_MAP
            assert "Primary" in row
            assert "Accent" in row
            assert "Background" in row

    def test_search_style_structure(self):
        r = ux.search_style("minimal", 3)
        assert isinstance(r, list)
        if r:
            row = r[0]
            # theme_composer._ux_find_style reads these
            assert "Style Category" in row
            assert "Effects & Animation" in row
            assert "Keywords" in row

    def test_search_typography_structure(self):
        r = ux.search_typography("professional", 3)
        assert isinstance(r, list)
        if r:
            row = r[0]
            # theme_composer._ux_find_typography reads these
            assert "Heading Font" in row
            assert "Body Font" in row

    def test_max_results_respected(self):
        assert len(ux.search_style("modern", 1)) <= 1


class TestDesignDeciderContract:
    """design_decider calls: get_design_system(q, variance, motion, density), search_style, is_available."""

    def test_get_design_system_all_consumer_keys(self):
        ds = ux.get_design_system("AI startup pitch", variance=5, motion=3, density=5)
        # design_decider._extract_ux_intelligence reads these
        for key in ["colors", "typography", "style_name", "style_effects",
                    "anti_patterns", "decision_rules", "pattern_name",
                    "pattern_sections", "pattern_cta_placement", "pattern_conversion"]:
            assert key in ds, f"design_decider needs {key}"

    def test_colors_accessible_by_role(self):
        ds = ux.get_design_system("fintech")
        c = ds["colors"]
        # design_decider checks 'primary', 'background', 'foreground'
        assert "primary" in c

    def test_typography_heading_body(self):
        ds = ux.get_design_system("fintech")
        t = ds["typography"]
        assert "heading" in t and "body" in t


class TestStoryPlannerContract:
    """story_planner calls: search_design(q,'product',1), search_landing(q,2)."""

    def test_search_design_product_type(self):
        r = ux.search_design("fintech", "product", 1)
        if r:
            # story_planner._detect_product_type reads Product Type + Primary Style Recommendation
            assert "Product Type" in r[0]

    def test_search_landing_section_order(self):
        r = ux.search_landing("hero", 2)
        if r:
            # story_planner._find_landing_pattern reads Section Order
            assert "Section Order" in r[0]

    def test_search_design_auto_domain(self):
        r = ux.search_design("fintech startup")
        assert isinstance(r, list)


class TestSlideSearchContract:
    """slide_search_adapter calls: search_reasoning(goal) — expects dict with 'pattern'."""

    def test_search_reasoning_has_pattern(self):
        r = ux.search_reasoning("dashboard")
        assert "pattern" in r
        assert isinstance(r["pattern"], str)

    def test_search_reasoning_decision_rules(self):
        r = ux.search_reasoning("dashboard")
        assert "decision_rules" in r
        assert isinstance(r["decision_rules"], dict)

    def test_search_reasoning_unknown_goal_safe(self):
        r = ux.search_reasoning("totally_unknown_goal_xyz")
        assert "pattern" in r  # default fallback, never crashes


class TestRobustness:
    def test_functions_never_raise(self):
        # All public functions must degrade gracefully, not raise
        ux.search_design("")
        ux.search_style("")
        ux.search_color("")
        ux.search_typography("")
        ux.search_landing("")
        ux.search_reasoning("")
        ds = ux.get_design_system("")
        assert isinstance(ds, dict)

    def test_returns_lists_not_none(self):
        for fn in [ux.search_design, ux.search_style, ux.search_color,
                   ux.search_typography, ux.search_landing]:
            assert isinstance(fn("anything"), list)

    def test_design_system_never_none(self):
        assert isinstance(ux.get_design_system("anything"), dict)

    def test_repeated_calls_consistent(self):
        a = ux.get_design_system("AI pitch", variance=4, motion=4, density=4)
        b = ux.get_design_system("AI pitch", variance=4, motion=4, density=4)
        assert a["colors"] == b["colors"]
