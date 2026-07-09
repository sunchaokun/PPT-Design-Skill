"""Phase 2: Design Decider — context-aware per-page design decisions."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from ppt_pro_max.planner.story_planner import StoryPlan, PagePlan
from ppt_pro_max.adapters.ui_ux_adapter import get_design_system
from ppt_pro_max.adapters.slide_search_adapter import (
    layout_for_goal,
    typography_for_slide,
    color_for_emotion,
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


class DesignDecider:
    def decide(
        self,
        story_plan: StoryPlan,
        theme: str | None = None,
        variance: int | None = None,
        motion: int | None = None,
        density: int | None = None,
    ) -> list[PageDesign]:
        design_system = get_design_system(story_plan.product_type or "general")

        page_designs = []
        for page in story_plan.pages:
            design = self._decide_page(page, story_plan.total_slides, design_system, motion)
            page_designs.append(design)

        return page_designs

    def _decide_page(self, page: PagePlan, total: int, design_system: dict, motion: int | None) -> PageDesign:
        layout = layout_for_goal(page.goal)
        typo = typography_for_slide(page.goal)
        color = color_for_emotion(page.emotion)
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
        )

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
