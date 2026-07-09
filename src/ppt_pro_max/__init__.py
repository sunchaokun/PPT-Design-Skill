"""PPT Design Skill — AI-powered PPT generation."""

from __future__ import annotations

from pathlib import Path
from typing import Any

__version__ = "0.1.0"

from ppt_pro_max.renderer.ppt_renderer import PPTRenderer
from ppt_pro_max.planner.story_planner import StoryPlanner
from ppt_pro_max.decider.design_decider import DesignDecider
from ppt_pro_max.content.content_generator import ContentGenerator


def generate_ppt(
    query: str,
    strategy: str | None = None,
    theme: str | None = None,
    slides: int | None = None,
    content_file: str | None = None,
    variance: int | None = None,
    motion: int | None = None,
    density: int | None = None,
    fetch_images: bool = False,
    image_mode: str = "placeholder",
    image_config: dict[str, Any] | None = None,
    llm_provider: str | None = None,
    llm_api_key: str | None = None,
    llm_base_url: str | None = None,
    llm_model: str | None = None,
    persist: bool = False,
    dry_run: bool = False,
    output: str | None = None,
) -> dict:
    planner = StoryPlanner()
    story_plan = planner.plan(query, strategy_override=strategy, slide_count_override=slides)

    decider = DesignDecider()
    page_designs = decider.decide(story_plan, theme=theme, variance=variance, motion=motion, density=density)

    generator = ContentGenerator(query=query)
    page_contents = generator.generate(page_designs, content_file=content_file)

    if dry_run:
        return {
            "dry_run": True,
            "strategy": story_plan.strategy,
            "page_count": story_plan.total_slides,
            "pages": [
                {
                    "position": d.position,
                    "goal": d.goal,
                    "emotion": d.emotion,
                    "layout": d.layout,
                }
                for d in page_designs
            ],
        }

    effective_image_mode = image_mode
    if fetch_images and image_mode == "placeholder":
        if llm_provider:
            effective_image_mode = "generate"
        else:
            effective_image_mode = "auto"

    if llm_provider or llm_api_key or llm_base_url or llm_model:
        if image_config is None:
            image_config = {}
        if llm_provider:
            image_config["llm_provider"] = llm_provider
        if llm_api_key:
            image_config["llm_api_key"] = llm_api_key
        if llm_base_url:
            image_config["llm_base_url"] = llm_base_url
        if llm_model:
            image_config["llm_model"] = llm_model

    renderer = PPTRenderer(image_mode=effective_image_mode, image_config=image_config)
    result = renderer.render(
        page_designs, page_contents,
        output_path=output, fetch_images=fetch_images,
        theme_name=theme, design_system=decider.design_system,
    )

    if persist:
        _persist_design_system(decider.design_system, result.get("output_path", ""))

    return result


def _persist_design_system(design_system: dict, pptx_path: str) -> None:
    from ppt_pro_max.renderer.theme_mapper import ThemeMapper

    mapper = ThemeMapper()
    theme = mapper.map(design_system)
    master_path = Path(pptx_path).parent / "design-system" / "MASTER.md"
    master_path.parent.mkdir(parents=True, exist_ok=True)

    colors = theme.get("colors", {})
    typo = theme.get("typography", {})

    lines = [
        "# Design System — MASTER.md",
        "",
        "## Colors",
        "",
    ]
    for role, hex_val in colors.items():
        lines.append(f"- **{role}**: `{hex_val}`")
    lines.extend([
        "",
        "## Typography",
        "",
        f"- **Heading**: {typo.get('heading', 'Inter')}",
        f"- **Body**: {typo.get('body', 'Inter')}",
        "",
        f"## Dark Mode: {'Yes' if theme.get('dark_mode') else 'No'}",
        "",
    ])

    master_path.write_text("\n".join(lines), encoding="utf-8")
