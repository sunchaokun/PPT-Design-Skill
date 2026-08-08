"""Phase 1: Story Planner — narrative-driven page structure.

Uses ui-ux-pro-max for product type detection and landing page patterns
when available, falling back to hardcoded strategies otherwise.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from ppt_pro_max.adapters.ui_ux_adapter import (
    search_design,
    search_landing,
    is_available,
)


@dataclass
class PagePlan:
    position: int
    goal: str
    emotion: str
    sparkline: str = ""


@dataclass
class StoryPlan:
    strategy: str
    total_slides: int
    pages: list[PagePlan] = field(default_factory=list)
    emotion_arc: str = ""
    product_type: str = ""
    landing_pattern: str = ""
    style_recommendation: str = ""


_STRATEGY_STRUCTURES: dict[str, dict[str, Any]] = {
    "YC Seed Deck": {
        "total": 10,
        "structure": ["hook", "problem", "solution", "product", "traction", "market", "business", "team", "financial", "cta"],
        "emotions": ["curiosity", "frustration", "hope", "trust", "confidence", "confidence", "trust", "warmth", "confidence", "urgency"],
        "sparkline": ["hook", "what-is", "what-could-be", "what-could-be", "proof", "proof", "trust", "trust", "what-could-be", "action"],
    },
    "Product Demo": {
        "total": 8,
        "structure": ["hook", "problem", "solution", "features", "demo", "testimonials", "pricing", "cta"],
        "emotions": ["curiosity", "frustration", "hope", "confidence", "trust", "warmth", "confidence", "urgency"],
        "sparkline": ["hook", "what-is", "what-could-be", "proof", "proof", "trust", "proof", "action"],
    },
    "Sales Pitch": {
        "total": 7,
        "structure": ["hook", "problem", "agitation", "solution", "proof", "offer", "cta"],
        "emotions": ["curiosity", "frustration", "fear", "hope", "trust", "confidence", "urgency"],
        "sparkline": ["hook", "what-is", "what-is", "what-could-be", "proof", "what-could-be", "action"],
    },
    "Education Course": {
        "total": 12,
        "structure": ["hook", "overview", "concept", "example", "practice", "concept", "example", "practice", "review", "extension", "summary", "cta"],
        "emotions": ["curiosity", "confidence", "understanding", "clarity", "engagement", "understanding", "clarity", "engagement", "confidence", "wonder", "satisfaction", "motivation"],
        "sparkline": ["hook", "what-is", "what-could-be", "proof", "proof", "what-could-be", "proof", "proof", "proof", "what-could-be", "proof", "action"],
    },
    "Training Workshop": {
        "total": 10,
        "structure": ["hook", "objectives", "concept", "demo", "exercise", "concept", "demo", "exercise", "review", "summary"],
        "emotions": ["curiosity", "confidence", "understanding", "trust", "engagement", "understanding", "trust", "engagement", "confidence", "satisfaction"],
        "sparkline": ["hook", "what-is", "what-could-be", "proof", "proof", "what-could-be", "proof", "proof", "proof", "proof"],
    },
    "Business Report": {
        "total": 8,
        "structure": ["hook", "overview", "metrics", "analysis", "findings", "recommendations", "action_plan", "summary"],
        "emotions": ["curiosity", "confidence", "trust", "understanding", "clarity", "confidence", "determination", "satisfaction"],
        "sparkline": ["hook", "what-is", "proof", "what-could-be", "proof", "what-could-be", "action", "proof"],
    },
}

_LANDING_SECTION_TO_GOAL: dict[str, str] = {
    "hero": "hook",
    "headline": "hook",
    "problem": "problem",
    "pain": "problem",
    "value": "solution",
    "solution": "solution",
    "features": "features",
    "feature": "features",
    "benefit": "features",
    "benefits": "features",
    "proof": "proof",
    "testimonial": "testimonials",
    "testimonials": "testimonials",
    "social": "testimonials",
    "pricing": "pricing",
    "comparison": "comparison",
    "demo": "demo",
    "product": "product",
    "cta": "cta",
    "call-to-action": "cta",
    "footer": "cta",
    "about": "content",
    "team": "content",
    "faq": "content",
    "metrics": "metrics",
    "stats": "metrics",
    "data": "data",
    "chart": "data",
    "overview": "overview",
    "summary": "content",
    "review": "content",
    "case": "proof",
    "process": "content",
    "how-it-works": "solution",
    "specs": "features",
    "detail": "features",
    "bento": "features",
    "tech": "features",
}

_GOAL_DEFAULT_EMOTION: dict[str, str] = {
    "hook": "curiosity",
    "problem": "frustration",
    "solution": "hope",
    "features": "confidence",
    "demo": "trust",
    "testimonials": "warmth",
    "pricing": "confidence",
    "cta": "urgency",
    "proof": "trust",
    "metrics": "confidence",
    "content": "confidence",
    "overview": "confidence",
    "comparison": "confidence",
    "data": "trust",
    "product": "trust",
}


class StoryPlanner:
    def plan(self, query: str, strategy_override: str | None = None, slide_count_override: int | None = None) -> StoryPlan:
        product_type, style_rec = self._detect_product_type(query)
        strategy_name = strategy_override or self._select_strategy(query, product_type)
        strategy = _STRATEGY_STRUCTURES.get(strategy_name, _STRATEGY_STRUCTURES["YC Seed Deck"])

        landing_pattern = ""
        if is_available() and not strategy_override:
            landing_pattern = self._find_landing_pattern(query)

        if landing_pattern:
            pages = self._build_from_landing(landing_pattern, slide_count_override)
        else:
            total = slide_count_override or strategy["total"]
            structure = strategy["structure"][:total]
            emotions = strategy["emotions"][:total]
            sparklines = strategy["sparkline"][:total]

            while len(structure) < total:
                structure.append("content")
                emotions.append("confidence")
                sparklines.append("proof")

            pages = [
                PagePlan(position=i + 1, goal=structure[i], emotion=emotions[i], sparkline=sparklines[i])
                for i in range(len(structure))
            ]

        return StoryPlan(
            strategy=strategy_name,
            total_slides=len(pages),
            pages=pages,
            emotion_arc="→".join(p.emotion for p in pages),
            product_type=product_type,
            landing_pattern=landing_pattern,
            style_recommendation=style_rec,
        )

    def _detect_product_type(self, query: str) -> tuple[str, str]:
        results = search_design(query, "product", 1)
        if results:
            pt = results[0].get("Product Type", results[0].get("product_type", "general"))
            style_rec = results[0].get("Primary Style Recommendation", "")
            return pt, style_rec
        return "general", ""

    def _find_landing_pattern(self, query: str) -> str:
        results = search_landing(query, 2)
        if results:
            sections = results[0].get("Section Order", "")
            if sections:
                return sections
        return ""

    def _build_from_landing(self, section_order: str, slide_count_override: int | None = None) -> list[PagePlan]:
        raw_sections = [s.strip() for s in section_order.split(",") if s.strip()]
        if not raw_sections:
            raw_sections = [s.strip() for s in section_order.split(">") if s.strip()]

        goals = []
        for section in raw_sections:
            section_lower = section.lower()
            matched = False
            for key, goal in _LANDING_SECTION_TO_GOAL.items():
                if key in section_lower:
                    goals.append(goal)
                    matched = True
                    break
            if not matched:
                goals.append("content")

        if slide_count_override and slide_count_override < len(goals):
            goals = goals[:slide_count_override]
            if goals[-1] != "cta":
                goals[-1] = "cta"
        elif slide_count_override and slide_count_override > len(goals):
            while len(goals) < slide_count_override:
                goals.append("content")
            if goals[-1] != "cta":
                goals[-1] = "cta"

        if not goals:
            goals = ["hook", "problem", "solution", "features", "cta"]

        return [
            PagePlan(
                position=i + 1,
                goal=goals[i],
                emotion=_GOAL_DEFAULT_EMOTION.get(goals[i], "confidence"),
                sparkline="hook" if i == 0 else ("action" if i == len(goals) - 1 else "proof"),
            )
            for i in range(len(goals))
        ]

    def _select_strategy(self, query: str, product_type: str) -> str:
        q = query.lower()
        if any(kw in q for kw in ["融资", "路演", "investor", "pitch", "seed", "series"]):
            return "YC Seed Deck"
        if any(kw in q for kw in ["产品", "demo", "product", "展示", "介绍"]):
            return "Product Demo"
        if any(kw in q for kw in ["销售", "sales", "报价", "offer"]):
            return "Sales Pitch"
        if any(kw in q for kw in ["教育", "课程", "education", "course", "lesson", "teaching"]):
            return "Education Course"
        if any(kw in q for kw in ["培训", "训练", "training", "workshop", "exercise"]):
            return "Training Workshop"
        if any(kw in q for kw in ["报告", "汇报", "report", "annual", "quarterly", "review"]):
            return "Business Report"
        return "YC Seed Deck"
