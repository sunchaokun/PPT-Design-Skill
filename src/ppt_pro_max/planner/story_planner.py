"""Phase 1: Story Planner — narrative-driven page structure."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from ppt_pro_max.adapters.ui_ux_adapter import search_design, load_csv


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
}


class StoryPlanner:
    def plan(self, query: str, strategy_override: str | None = None, slide_count_override: int | None = None) -> StoryPlan:
        product_type = self._detect_product_type(query)
        strategy_name = strategy_override or self._select_strategy(query, product_type)
        strategy = _STRATEGY_STRUCTURES.get(strategy_name, _STRATEGY_STRUCTURES["YC Seed Deck"])

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
            for i in range(total)
        ]

        return StoryPlan(
            strategy=strategy_name,
            total_slides=total,
            pages=pages,
            emotion_arc="→".join(emotions),
            product_type=product_type,
        )

    def _detect_product_type(self, query: str) -> str:
        results = search_design(query, "product", 1)
        if results:
            return results[0].get("product_type", "general")
        return "general"

    def _select_strategy(self, query: str, product_type: str) -> str:
        q = query.lower()
        if any(kw in q for kw in ["融资", "路演", "investor", "pitch", "seed", "series"]):
            return "YC Seed Deck"
        if any(kw in q for kw in ["产品", "demo", "product", "展示", "介绍"]):
            return "Product Demo"
        if any(kw in q for kw in ["销售", "sales", "报价", "报价", "offer"]):
            return "Sales Pitch"
        return "YC Seed Deck"
