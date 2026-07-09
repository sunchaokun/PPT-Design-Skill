"""PPT Design Skill — AI-powered PPT generation."""

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
    persist: bool = False,
    dry_run: bool = False,
    output: str | None = None,
) -> dict:
    planner = StoryPlanner()
    story_plan = planner.plan(query, strategy_override=strategy, slide_count_override=slides)

    decider = DesignDecider()
    page_designs = decider.decide(story_plan, theme=theme, variance=variance, motion=motion, density=density)

    generator = ContentGenerator()
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

    renderer = PPTRenderer()
    result = renderer.render(page_designs, page_contents, output_path=output, fetch_images=fetch_images)

    if persist:
        pass

    return result
