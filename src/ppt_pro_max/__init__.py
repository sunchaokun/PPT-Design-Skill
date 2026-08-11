"""PPT Design Skill — AI-powered PPT generation."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

__version__ = "0.17.0"

from ppt_pro_max.planner.story_planner import StoryPlanner
from ppt_pro_max.decider.design_decider import DesignDecider
from ppt_pro_max.content.content_generator import ContentGenerator
from ppt_pro_max.renderer.theme_composer import ThemeComposer


def _ensure_dotenv():
    from ppt_pro_max.build_helpers import _load_dotenv
    _load_dotenv()


def fetch_image(
    keywords: str,
    *,
    mode: str = "auto",
    emotion: str = "",
    goal: str = "",
    width: int = 1920,
    height: int = 1080,
    llm_provider: str | None = None,
    llm_api_key: str | None = None,
    llm_base_url: str | None = None,
    llm_model: str | None = None,
    unsplash_access_key: str | None = None,
    pexels_api_key: str | None = None,
    image_cache_dir: str | None = None,
    auto_detect: bool = True,
) -> dict[str, Any]:
    """Standalone image generation/fetch API — no PPT required.

    Returns dict with keys: path (str|None), mode, provider, keywords, width, height.
    """
    _ensure_dotenv()
    from ppt_pro_max.renderer.image_fetcher import ImageFetcher

    fetcher = ImageFetcher(
        mode=mode,
        llm_provider=llm_provider,
        llm_api_key=llm_api_key,
        llm_base_url=llm_base_url,
        llm_model=llm_model,
        unsplash_access_key=unsplash_access_key,
        pexels_api_key=pexels_api_key,
        image_cache_dir=image_cache_dir,
        auto_detect=auto_detect,
    )
    path = fetcher.fetch(keywords, emotion=emotion, goal=goal, width=width, height=height)
    return {
        "path": path,
        "mode": mode,
        "provider": fetcher.llm_provider or "",
        "keywords": keywords,
        "width": width,
        "height": height,
    }


def extract_design_dna(pptx_path: str) -> dict[str, Any]:
    from ppt_pro_max.enterprise.design_dna_extractor import DesignDNAExtractor

    extractor = DesignDNAExtractor()
    dna = extractor.extract(pptx_path)

    return {
        "source_path": dna.source_path,
        "slide_width_emu": dna.slide_width_emu,
        "slide_height_emu": dna.slide_height_emu,
        "num_slides": len(dna.slides),
        "slides": [
            {
                "slide_index": s.slide_index,
                "page_type": s.page_type,
                "text_zones": [
                    {
                        "zone_id": z.zone_id,
                        "role": z.role,
                        "text": z.text,
                        "font_name": z.font_name,
                        "font_size_pt": z.font_size_pt,
                        "bold": z.bold,
                        "color_hex": z.color_hex,
                        "bounds": list(z.bounds),
                    }
                    for z in s.text_zones
                ],
                "image_refs": [{"shape_id": r.get("shape_id"), "is_background": r.get("is_background")} for r in s.image_refs],
                "layout_name": s.layout_name,
                "background_colors": s.background_colors,
                "notes": s.notes_text,
            }
            for s in dna.slides
        ],
        "color_palette": dna.color_palette,
        "font_scheme": dna.font_scheme,
        "cjk_font_scheme": dna.cjk_font_scheme,
        "actual_colors": dna.actual_colors,
        "actual_fonts": dna.actual_fonts,
        "actual_font_sizes": {str(k): v for k, v in dna.actual_font_sizes.items()},
        "has_logo": dna.logo_blob is not None,
        "decorative_groups_count": len(dna.decorative_groups),
        "brand_spec": {
            "source": dna.brand_spec.source if dna.brand_spec else "none",
            "colors": dna.brand_spec.colors if dna.brand_spec else None,
            "fonts": dna.brand_spec.fonts if dna.brand_spec else None,
            "dark_mode": dna.brand_spec.dark_mode if dna.brand_spec else False,
        },
    }


def generate_ppt(
    query: str = "",
    strategy: str | None = None,
    theme: str | None = None,
    style: str | None = None,
    palette: str | None = None,
    fonts: str | None = None,
    decoration: str | None = None,
    layout_variant: str | None = None,
    mood: str | None = None,
    style_seed: int | None = None,
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
    proposal: bool = False,
    confirmed_proposal: str | None = None,
    materials_dir: str | None = None,
    content: dict[str, Any] | None = None,
    auto_detect: bool = True,
) -> dict:
    _ensure_dotenv()
    if content is not None and content_file is None:
        import tempfile as _tf
        _content_dir = _tf.mkdtemp(prefix="ppt_content_")
        _content_file = os.path.join(_content_dir, "content.json")
        with open(_content_file, "w", encoding="utf-8") as _f:
            json.dump(content, _f, ensure_ascii=False)
        content_file = _content_file
        import atexit
        import shutil as _shutil
        atexit.register(lambda: _shutil.rmtree(_content_dir, ignore_errors=True))

    if content_file and _has_slides(content_file):
        if proposal:
            import warnings as _warnings
            _warnings.warn(
                "content_file with slides is ignored when proposal=True",
                UserWarning,
                stacklevel=2,
            )

    if proposal:
        return _generate_proposals(
            query=query, style=style or theme,
            output_dir=output, style_seed=style_seed,
            materials_dir=materials_dir,
        )
    return _generate_ppt_freestyle(
        query=query, strategy=strategy, theme=theme, style=style,
        palette=palette, fonts=fonts, decoration=decoration,
        layout_variant=layout_variant, mood=mood, style_seed=style_seed,
        slides=slides, content_file=content_file, variance=variance,
        motion=motion, density=density, fetch_images=fetch_images,
        image_mode=image_mode, image_config=image_config,
        llm_provider=llm_provider, llm_api_key=llm_api_key,
        llm_base_url=llm_base_url, llm_model=llm_model,
        persist=persist, dry_run=dry_run, output=output,
        auto_detect=auto_detect,
    )


def _generate_ppt_freestyle(
    query: str = "",
    strategy: str | None = None,
    theme: str | None = None,
    style: str | None = None,
    palette: str | None = None,
    fonts: str | None = None,
    decoration: str | None = None,
    layout_variant: str | None = None,
    mood: str | None = None,
    style_seed: int | None = None,
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
    auto_detect: bool = True,
) -> dict:
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

    if image_config is None:
        image_config = {}
    if not auto_detect:
        image_config["auto_detect"] = False

    composed_theme = None
    if style or palette or fonts or decoration or layout_variant or mood:
        composer = ThemeComposer()
        composed_theme = composer.compose(
            style=style or theme,
            palette=palette,
            fonts=fonts,
            decoration=decoration,
            layout=layout_variant,
            mood=mood,
            seed=style_seed,
        )

    image_fetcher = None
    if effective_image_mode and effective_image_mode != "placeholder":
        from ppt_pro_max.renderer.image_fetcher import ImageFetcher
        image_fetcher = ImageFetcher(mode=effective_image_mode, **dict(image_config))

    page_dicts = None
    fallback_theme = None
    if content_file and _has_slides(content_file):
        page_dicts = _load_content_json_to_pages(content_file)
        if page_dicts is not None:
            if dry_run:
                return {
                    "dry_run": True,
                    "strategy": "content.json",
                    "page_count": len(page_dicts),
                    "pages": [
                        {"position": i, "goal": p.get("goal", "content"),
                         "emotion": None, "layout": None}
                        for i, p in enumerate(page_dicts)
                    ],
                }
            _fetch_missing_images(page_dicts, image_fetcher)

    if page_dicts is None:
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
        fallback_theme = decider.design_system
        page_dicts = _build_freestyle_page_dicts(page_designs, page_contents, image_fetcher)

    result = _render_page_dicts(
        page_dicts, composed_theme, output,
        fallback_theme=fallback_theme,
    )

    if composed_theme:
        result["theme_atoms"] = composed_theme.get("atoms", {})

    if persist and fallback_theme:
        _persist_design_system(fallback_theme, result.get("output_path", ""))

    return result


def _build_freestyle_page_dicts(page_designs, page_contents, image_fetcher) -> list[dict]:
    page_dicts = []
    for design, content in zip(page_designs, page_contents):
        page = {
            "goal": design.goal,
            "title": content.title,
            "subtitle": content.subtitle,
            "bullets": content.bullets,
            "metrics": content.metrics,
            "chart": content.chart_data,
            "quote": content.quote,
            "image_keywords": content.image_keywords,
        }
        page_dicts.append(page)
    _fetch_missing_images(page_dicts, image_fetcher)
    return page_dicts


def _fetch_missing_images(page_dicts, image_fetcher) -> None:
    if image_fetcher is None:
        return
    for page in page_dicts:
        if page.get("image") and os.path.isfile(page["image"]):
            continue
        keywords = page.get("image_keywords") or page.get("goal") or ""
        if not keywords:
            continue
        try:
            fetched = image_fetcher.fetch(
                keywords=keywords,
                goal=page.get("goal", ""),
                width=1920,
                height=1080,
            )
            if fetched and os.path.isfile(fetched):
                page["image"] = fetched
        except Exception:
            pass


def _render_page_dicts(
    page_dicts,
    composed_theme,
    output,
    fallback_theme=None,
) -> dict:
    from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
    from ppt_pro_max.enterprise.brand_spec import BrandSpec

    effective_theme = composed_theme or fallback_theme
    if not effective_theme:
        composer = ThemeComposer()
        effective_theme = composer.compose(style="professional")

    colors = effective_theme.get("colors", {})
    typography = effective_theme.get("typography", {})
    brand_spec = BrandSpec(
        source="freestyle_composed",
        colors=colors,
        fonts=typography,
        dark_mode=effective_theme.get("dark_mode", False),
    )

    layout_variant_out = dict(effective_theme.get("layout_variant", {}))
    deco_atom = effective_theme.get("atoms", {}).get("decoration", "accent-bar")
    layout_variant_out["decoration_style"] = deco_atom
    layout_variant_out["decoration"] = effective_theme.get("decoration", {})

    precision = PrecisionRenderer(brand_spec=brand_spec)
    prs = precision.create_presentation()

    total = len(page_dicts)
    for i, page in enumerate(page_dicts):
        precision.render_slide(prs, page,
                               layout_variant=layout_variant_out,
                               page_index=i, total_pages=total)

    if output is None:
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output = f"presentation_{timestamp}.pptx"

    output_dir = os.path.dirname(output)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    precision.save(prs, output)

    return {
        "output_path": os.path.abspath(output),
        "page_count": len(page_dicts),
        "strategy": "generated",
        "render_mode": "precision",
    }


def _has_slides(content_file: str) -> bool:
    try:
        with open(content_file, "r", encoding="utf-8") as f:
            raw = json.load(f)
        return isinstance(raw.get("slides"), list) and len(raw["slides"]) > 0
    except (FileNotFoundError, json.JSONDecodeError):
        return False


def _load_content_json_to_pages(content_file: str):
    from ppt_pro_max.enterprise.content_parser import load_enterprise_content
    from ppt_pro_max.content.content_generator import _GOAL_IMAGE_KEYWORDS

    try:
        with open(content_file, "r", encoding="utf-8") as f:
            raw = json.load(f)

        project_dir = os.path.dirname(os.path.abspath(content_file))
        pages = load_enterprise_content(raw, project_dir)

        if not pages:
            return None

        for page in pages:
            if not page.get("image_keywords"):
                page["image_keywords"] = _GOAL_IMAGE_KEYWORDS.get(page.get("goal", ""), "abstract")

        return pages
    except (json.JSONDecodeError, TypeError, OSError, KeyError):
        return None


def _generate_proposals(
    query: str,
    style: str | None = None,
    output_dir: str | None = None,
    style_seed: int | None = None,
    materials_dir: str | None = None,
) -> dict:
    from ppt_pro_max.enterprise.proposal_generator import ProposalGenerator

    if output_dir is None:
        output_dir = "output"

    gen = ProposalGenerator()
    proposals = gen.generate(
        query=query,
        style=style,
        output_dir=output_dir,
        project_dir=materials_dir,
        seed=style_seed,
    )
    return {"proposals": proposals}


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
