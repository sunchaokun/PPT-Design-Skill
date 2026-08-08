"""Phase 2: Design Decider — context-aware per-page design decisions.

Now actually consumes the design intelligence from ui-ux-pro-max:
  - Product-specific color palettes (192 palettes from colors.csv)
  - Style effects and specifications (84 styles from styles.csv)
  - Typography pairings (74 pairs from typography.csv)
  - UI reasoning rules (161 rules from ui-reasoning.csv)
  - Anti-patterns per industry
  - Decision rules (conditional design logic)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from ppt_pro_max.adapters.ui_ux_adapter import (
    get_design_system,
    search_style,
    is_available,
)
from ppt_pro_max.adapters.slide_search_adapter import (
    layout_for_goal,
    background_config,
    full_bleed,
    pattern_break,
)


@dataclass
class PageDesign:
    position: int
    goal: str
    emotion: str
    layout: str = ""
    typography: dict[str, Any] = field(default_factory=dict)
    color_treatment: dict[str, Any] = field(default_factory=dict)
    copy_formula: str = ""
    chart_type: str | None = None
    background: dict[str, Any] = field(default_factory=dict)
    break_pattern: bool = False
    full_bleed: bool = False
    animation: str = "fade-in"
    transition: str = "fade"
    style_effects: str = ""
    anti_patterns: str = ""
    decision_rules: dict[str, Any] = field(default_factory=dict)


_GOAL_COPY_FORMULA: dict[str, str] = {
    "hook": "AIDA",
    "problem": "PAS",
    "agitation": "Cost of Inaction",
    "solution": "FAB",
    "proof": "Social Proof",
    "traction": "Proof Stack",
    "cta": "AIDA + Urgency",
    "features": "Feature-Benefit",
    "testimonials": "Social Proof",
    "pricing": "Anchor-Compare",
}

_GOAL_CHART: dict[str, str | None] = {
    "traction": "Line Chart",
    "metrics": "Bar Chart Vertical",
    "market": "Pie Chart",
    "financial": "Bar Chart Vertical",
    "comparison": "Bar Chart Horizontal",
}

_STYLE_EFFECTS_MAP: dict[str, str] = {
    "Glassmorphism": "backdrop-blur translucent-overlay",
    "Neumorphism": "soft-shadow press-effect",
    "Brutalism": "no-radius bold-contrast",
    "Cyberpunk UI": "neon-glow scanlines",
    "Aurora UI": "flowing-gradient blend-mode",
    "Claymorphism": "inner-shadow soft-press",
    "Flat Design": "no-shadow solid-fill",
    "Soft UI": "soft-shadow subtle-gradient",
    "Minimalism": "no-decoration whitespace-heavy",
    "Dark Mode": "dark-bg glow-accent",
    "Gradient Design": "gradient-fill mesh-gradient",
    "Neomorphism": "dual-shadow emboss-effect",
    "Retro UI": "vintage-colors rounded-bold",
    "Organic Design": "curved-shapes natural-colors",
    "Bento Grid": "grid-cards rounded-corners",
}


class DesignDecider:
    def __init__(self):
        self._design_system: dict[str, Any] = {}
        self._ux_colors: dict[str, str] = {}
        self._ux_typography: dict[str, str] = {}
        self._ux_style_name: str = ""
        self._ux_style_effects: str = ""
        self._ux_anti_patterns: str = ""
        self._ux_decision_rules: dict[str, Any] = {}
        self._ux_pattern_name: str = ""
        self._ux_pattern_sections: str = ""
        self._ux_pattern_cta: str = ""
        self._ux_pattern_conversion: str = ""

    def decide(
        self,
        story_plan: Any,
        theme: str | None = None,
        variance: int | None = None,
        motion: int | None = None,
        density: int | None = None,
    ) -> list[PageDesign]:
        query = getattr(story_plan, "product_type", "general") or "general"
        style_rec = getattr(story_plan, "style_recommendation", "") or ""

        self._design_system = get_design_system(
            query, variance=variance, motion=motion, density=density
        )
        self._extract_ux_intelligence(self._design_system)

        if is_available() and style_rec:
            self._enrich_with_style_search(style_rec)

        page_designs = []
        for page in story_plan.pages:
            design = self._decide_page(page, story_plan.total_slides, motion)
            page_designs.append(design)

        return page_designs

    @property
    def design_system(self) -> dict[str, Any]:
        return self._design_system

    @property
    def ux_colors(self) -> dict[str, str]:
        return self._ux_colors

    @property
    def ux_typography(self) -> dict[str, str]:
        return self._ux_typography

    @property
    def ux_style_effects(self) -> str:
        return self._ux_style_effects

    @property
    def ux_anti_patterns(self) -> str:
        return self._ux_anti_patterns

    def _extract_ux_intelligence(self, ds: dict[str, Any]) -> None:
        colors = ds.get("colors", {})
        if colors and any(k in colors for k in ("primary", "background", "foreground")):
            self._ux_colors = colors

        typo = ds.get("typography", {})
        if typo and any(k in typo for k in ("heading", "body")):
            self._ux_typography = {
                "heading": typo.get("heading", "Inter"),
                "body": typo.get("body", "Inter"),
                "mood": typo.get("mood", ""),
                "best_for": typo.get("best_for", ""),
            }

        self._ux_style_name = ds.get("style_name", "")
        self._ux_style_effects = ds.get("style_effects", "")
        self._ux_anti_patterns = ds.get("anti_patterns", "")
        self._ux_decision_rules = ds.get("decision_rules", {})
        self._ux_pattern_name = ds.get("pattern_name", "")
        self._ux_pattern_sections = ds.get("pattern_sections", "")
        self._ux_pattern_cta = ds.get("pattern_cta_placement", "")
        self._ux_pattern_conversion = ds.get("pattern_conversion", "")

    def _enrich_with_style_search(self, style_rec: str) -> None:
        styles = search_style(style_rec, 2)
        if styles:
            best = styles[0]
            style_name = best.get("Style Category", "")
            if style_name and not self._ux_style_name:
                self._ux_style_name = style_name
            effects = best.get("Effects & Animation", "")
            if effects and not self._ux_style_effects:
                self._ux_style_effects = effects
            if style_name in _STYLE_EFFECTS_MAP and not self._ux_style_effects:
                self._ux_style_effects = _STYLE_EFFECTS_MAP[style_name]

    def _decide_page(self, page: Any, total: int, motion: int | None) -> PageDesign:
        layout = layout_for_goal(page.goal)
        typo = self._resolve_typography(page.goal)
        color = self._resolve_color_treatment(page.emotion, page.goal)
        bg = background_config(page.goal)
        is_bleed = full_bleed(page.position, page.emotion)
        is_break = pattern_break(page.position, total)
        copy_formula = _GOAL_COPY_FORMULA.get(page.goal, "AIDA")
        chart_type = _GOAL_CHART.get(page.goal)
        animation = self._select_animation(page.goal, motion)
        transition = self._select_transition(page.emotion)

        return PageDesign(
            position=page.position,
            goal=page.goal,
            emotion=page.emotion,
            layout=layout,
            typography=typo,
            color_treatment=color,
            copy_formula=copy_formula,
            chart_type=chart_type,
            background=bg,
            break_pattern=is_break,
            full_bleed=is_bleed,
            animation=animation,
            transition=transition,
            style_effects=self._ux_style_effects,
            anti_patterns=self._ux_anti_patterns,
            decision_rules=self._ux_decision_rules,
        )

    def _resolve_typography(self, goal: str) -> dict[str, Any]:
        base = {
            "primary_size": 28,
            "secondary_size": 16,
            "weight_contrast": "600-400",
        }
        if self._ux_typography:
            base["heading_font"] = self._ux_typography.get("heading", "Inter")
            base["body_font"] = self._ux_typography.get("body", "Inter")
            base["mood"] = self._ux_typography.get("mood", "")
        return base

    def _resolve_color_treatment(self, emotion: str, goal: str) -> dict[str, Any]:
        base = {
            "background": "default",
            "text_color": "foreground",
            "accent_usage": "moderate",
        }
        if self._ux_colors:
            base["ux_primary"] = self._ux_colors.get("primary", "")
            base["ux_accent"] = self._ux_colors.get("accent", "")
            base["ux_background"] = self._ux_colors.get("background", "")
            base["ux_foreground"] = self._ux_colors.get("foreground", "")
            base["ux_muted"] = self._ux_colors.get("muted", "")
            base["ux_border"] = self._ux_colors.get("border", "")
            base["ux_palette"] = self._ux_colors

        if self._ux_style_effects:
            base["style_effects"] = self._ux_style_effects

        if self._ux_anti_patterns:
            base["anti_patterns"] = self._ux_anti_patterns

        if self._ux_decision_rules:
            base["decision_rules"] = self._ux_decision_rules

        return base

    def _select_animation(self, goal: str, motion: int | None) -> str:
        if motion is not None and motion <= 2:
            return "appear"
        _ANIMATION_MAP: dict[str, str] = {
            "hook": "fly-in-bottom",
            "cta": "grow",
            "features": "fade-in-stagger",
            "metrics": "fade-in-stagger",
        }
        return _ANIMATION_MAP.get(goal, "fade-in")

    def _select_transition(self, emotion: str) -> str:
        _TRANSITION_MAP: dict[str, str] = {
            "frustration": "push-left",
            "hope": "wipe-down",
            "urgency": "fade-slow",
            "fear": "cut",
        }
        return _TRANSITION_MAP.get(emotion, "fade")
