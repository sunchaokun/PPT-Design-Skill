"""Phase 3: Content Generator — copy formulas and placeholder content."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any

from ppt_pro_max.decider.design_decider import PageDesign


@dataclass
class PageContent:
    position: int
    goal: str
    title: str = ""
    subtitle: str | None = None
    bullets: list[str] = field(default_factory=list)
    metrics: list[dict[str, str]] | None = None
    quote: dict[str, str] | None = None
    chart_data: dict[str, Any] | None = None
    image_keywords: str = ""


_COPY_TEMPLATES: dict[str, dict[str, str]] = {
    "AIDA": {
        "title": "{Attention hook}",
        "subtitle": "{Interesting detail}. {Desirable outcome}.",
    },
    "PAS": {
        "title": "你正面临{问题}",
        "subtitle": "它正在{恶化}. {解决方案}修复这一切。",
    },
    "Cost of Inaction": {
        "title": "没有{方案}，你每天损失{金额}",
        "subtitle": None,
    },
    "FAB": {
        "title": "{Feature}让你{Advantage}",
        "subtitle": "所以你能{Benefit}。",
    },
    "Social Proof": {
        "title": "{N}+用户信赖{产品}",
        "subtitle": None,
    },
    "Proof Stack": {
        "title": "{metric1} | {metric2} | {metric3}",
        "subtitle": None,
    },
    "AIDA + Urgency": {
        "title": "立即{行动}",
        "subtitle": "限时{优惠}。",
    },
    "Feature-Benefit": {
        "title": "{Feature}",
        "subtitle": "{Benefit}",
    },
    "Anchor-Compare": {
        "title": "选择适合你的方案",
        "subtitle": None,
    },
}


_GOAL_IMAGE_KEYWORDS: dict[str, str] = {
    "hook": "abstract technology innovation",
    "problem": "frustration stress problem dark",
    "agitation": "urgency alarm warning",
    "solution": "hope light solution bright",
    "features": "product features grid clean",
    "traction": "growth chart upward success",
    "market": "market globe world opportunity",
    "team": "team collaboration people",
    "cta": "action button start begin",
}


class ContentGenerator:
    def generate(self, page_designs: list[PageDesign], content_file: str | None = None) -> list[PageContent]:
        user_content = self._load_user_content(content_file)

        contents = []
        for design in page_designs:
            content = self._generate_page(design, user_content)
            contents.append(content)

        return contents

    def _generate_page(self, design: PageDesign, user_content: dict[str, Any]) -> PageContent:
        template = _COPY_TEMPLATES.get(design.copy_formula, _COPY_TEMPLATES["AIDA"])
        goal = design.goal

        title = self._fill_template(template["title"], goal, user_content)
        subtitle = self._fill_template(template.get("subtitle", ""), goal, user_content) if template.get("subtitle") else None

        bullets = self._generate_bullets(goal, user_content)
        metrics = self._generate_metrics(goal, user_content)
        chart_data = self._generate_chart_data(goal, design.chart_type, user_content)
        image_keywords = _GOAL_IMAGE_KEYWORDS.get(goal, "abstract")

        return PageContent(
            position=design.position,
            goal=goal,
            title=title,
            subtitle=subtitle,
            bullets=bullets,
            metrics=metrics,
            chart_data=chart_data,
            image_keywords=image_keywords,
        )

    def _fill_template(self, template: str, goal: str, user_content: dict[str, Any]) -> str:
        if not template:
            return f"[{goal.upper()}]"
        for key, value in user_content.items():
            if isinstance(value, str):
                template = template.replace(f"{{{key}}}", value)
        return template

    def _generate_bullets(self, goal: str, user_content: dict[str, Any]) -> list[str]:
        if goal == "problem" and "pain_points" in user_content:
            return [p.get("title", p.get("desc", "")) for p in user_content["pain_points"][:3]]
        if goal == "features" and "features" in user_content:
            return [f.get("title", f.get("name", "")) for f in user_content["features"][:3]]
        return [f"[{goal} 要点1]", f"[{goal} 要点2]", f"[{goal} 要点3]"]

    def _generate_metrics(self, goal: str, user_content: dict[str, Any]) -> list[dict[str, str]] | None:
        if goal in ("traction", "metrics") and "metrics" in user_content:
            return [{"label": k, "value": str(v)} for k, v in user_content["metrics"].items()[:4]]
        if goal in ("traction", "metrics"):
            return [
                {"label": "用户数", "value": "[N]+"},
                {"label": "留存率", "value": "[X]%"},
                {"label": "增长率", "value": "[X]x"},
                {"label": "ARR", "value": "$[X]M"},
            ]
        return None

    def _generate_chart_data(self, goal: str, chart_type: str | None, user_content: dict[str, Any]) -> dict[str, Any] | None:
        if chart_type is None:
            return None
        if "chart_data" in user_content:
            for key, data in user_content["chart_data"].items():
                return {"type": chart_type, "data": data}
        return {
            "type": chart_type,
            "data": {
                "labels": ["Q1", "Q2", "Q3", "Q4"],
                "values": [10, 25, 45, 80],
            },
        }

    def _load_user_content(self, content_file: str | None) -> dict[str, Any]:
        if content_file is None:
            return {}
        try:
            with open(content_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {}
