"""PPT Design Skill — AI-powered PPT generation."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

__version__ = "0.8.0"

from ppt_pro_max.renderer.ppt_renderer import PPTRenderer
from ppt_pro_max.planner.story_planner import StoryPlanner
from ppt_pro_max.decider.design_decider import DesignDecider
from ppt_pro_max.content.content_generator import ContentGenerator
from ppt_pro_max.renderer.theme_composer import ThemeComposer


def query_component_library(
    type: str | None = None,
    category: str | None = None,
    node_count: int | None = None,
    component_library: str | None = None,
    project_dir: str | None = None,
) -> dict | list:
    from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path

    db_path = find_db_path(component_library, project_dir)
    if db_path is None:
        return {} if type is None else []

    lib = ComponentLibrary(db_path)
    try:
        if type is None:
            return lib.catalog()
        results = lib.search(type=type, category=category or "", node_count=node_count, limit=50)
        return results
    finally:
        lib.close()


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
    query: str,
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
    project: str | None = None,
    business_mode: str | None = None,
    review: bool = False,
    review_file: str | None = None,
    output_version: int | None = None,
    from_version: int | None = None,
    pages: str | None = None,
    history: bool = False,
    proposal: bool = False,
    confirmed_proposal: str | None = None,
    materials_dir: str | None = None,
    beautify: str | None = None,
    beautify_mode: str | None = None,
    component_library: str | None = None,
    content: dict[str, Any] | None = None,
    auto_detect: bool = True,
    patches: list[dict[str, Any]] | None = None,
) -> dict:
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

    if beautify:
        return _generate_ppt_beautify(
            input_pptx=beautify,
            style=style or theme,
            output=output,
            density=density,
            motion=motion,
            beautify_mode=beautify_mode,
            component_library=component_library,
            patches=patches,
        )
    if proposal:
        return _generate_proposals(
            query=query, style=style or theme, project=project,
            output_dir=output, style_seed=style_seed,
            materials_dir=materials_dir,
        )
    if project:
        return _generate_ppt_enterprise(
            query=query, project=project, dry_run=dry_run,
            business_mode=business_mode, density=density,
            review=review, review_file=review_file,
            output_version=output_version, from_version=from_version,
            pages=pages, history=history, output=output,
            motion=motion, content_file=content_file,
            fetch_images=fetch_images,
            llm_provider=llm_provider, llm_api_key=llm_api_key,
            llm_base_url=llm_base_url, llm_model=llm_model,
            image_mode=image_mode, image_config=image_config,
            materials_dir=materials_dir,
            confirmed_proposal=confirmed_proposal,
            component_library=component_library,
            auto_detect=auto_detect,
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
        component_library=component_library,
    )


def _generate_ppt_freestyle(
    query: str,
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
    component_library: str | None = None,
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

    from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
    db_path = find_db_path(component_library)
    component_lib = None
    if db_path:
        try:
            component_lib = ComponentLibrary(db_path)
        except Exception:
            component_lib = None

    try:
        if component_lib is not None:
            result = _render_freestyle_with_components(
                page_designs, page_contents, composed_theme,
                component_lib, output, fetch_images, effective_image_mode,
                image_config, decider.design_system, query,
            )
        else:
            renderer = PPTRenderer(image_mode=effective_image_mode, image_config=image_config)
            result = renderer.render(
                page_designs, page_contents,
                output_path=output, fetch_images=fetch_images,
                theme_name=theme, design_system=decider.design_system,
                composed_theme=composed_theme,
            )
    finally:
        if component_lib:
            try:
                component_lib.close()
            except Exception:
                pass

    if composed_theme:
        result["theme_atoms"] = composed_theme.get("atoms", {})

    if persist:
        _persist_design_system(decider.design_system, result.get("output_path", ""))

    return result


def _render_freestyle_with_components(
    page_designs,
    page_contents,
    composed_theme,
    component_lib,
    output,
    fetch_images,
    image_mode,
    image_config,
    design_system,
    query,
) -> dict:
    from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
    from ppt_pro_max.enterprise.brand_spec import BrandSpec
    from ppt_pro_max.enterprise.content_parser import infer_component_category

    effective_theme = composed_theme or design_system or {}
    colors = effective_theme.get("colors", {})
    typography = effective_theme.get("typography", {})
    brand_spec = BrandSpec(
        source="freestyle_composed",
        colors=colors,
        fonts=typography,
        dark_mode=effective_theme.get("dark_mode", False),
    )

    page_dicts = []
    for design, content in zip(page_designs, page_contents):
        page = {
            "goal": design.goal,
            "title": content.title,
            "subtitle": content.subtitle,
            "bullets": content.bullets,
            "metrics": content.metrics,
            "chart_data": content.chart_data,
            "quote": content.quote,
            "image_keywords": content.image_keywords,
        }
        if content.bullets and not page.get("component_type"):
            comp_type, comp_cat = infer_component_category(content.bullets)
            if comp_type:
                page["component_type"] = comp_type
                page["component_category"] = comp_cat
        page_dicts.append(page)

    precision = PrecisionRenderer(brand_spec=brand_spec)
    prs = precision.create_presentation()

    total = len(page_dicts)
    for i, page in enumerate(page_dicts):
        precision.render_slide(prs, page, component_lib=component_lib,
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
        "render_mode": "precision_with_components",
    }


def _generate_proposals(
    query: str,
    style: str | None = None,
    project: str | None = None,
    output_dir: str | None = None,
    style_seed: int | None = None,
    materials_dir: str | None = None,
) -> dict:
    from ppt_pro_max.enterprise.proposal_generator import ProposalGenerator

    if output_dir is None:
        if project:
            output_dir = os.path.join(project, "output")
        else:
            output_dir = "output"

    gen = ProposalGenerator()
    proposals = gen.generate(
        query=query,
        style=style,
        output_dir=output_dir,
        project_dir=project or materials_dir,
        seed=style_seed,
    )
    return {"proposals": proposals}


def _generate_ppt_enterprise(
    query: str,
    project: str,
    dry_run: bool = False,
    business_mode: str | None = None,
    density: int | None = None,
    motion: int | None = None,
    review: bool = False,
    review_file: str | None = None,
    output_version: int | None = None,
    from_version: int | None = None,
    pages: str | None = None,
    history: bool = False,
    output: str | None = None,
    content_file: str | None = None,
    fetch_images: bool = False,
    llm_provider: str | None = None,
    llm_api_key: str | None = None,
    llm_base_url: str | None = None,
    llm_model: str | None = None,
    image_mode: str = "placeholder",
    image_config: dict[str, Any] | None = None,
    materials_dir: str | None = None,
    confirmed_proposal: str | None = None,
    component_library: str | None = None,
    auto_detect: bool = True,
) -> dict:
    import warnings
    warnings.warn(
        "Enterprise pipeline (project=) is deprecated due to known quality issues. "
        "Use FreeStyle or Build mode instead.",
        DeprecationWarning,
        stacklevel=2,
    )
    from ppt_pro_max.enterprise.pipeline import EnterprisePipeline

    effective_image_mode = image_mode
    if fetch_images and effective_image_mode == "placeholder":
        if llm_provider:
            effective_image_mode = "generate"
        else:
            effective_image_mode = "auto"

    pipeline = EnterprisePipeline()
    if image_config is None:
        image_config = {}
    if not auto_detect:
        image_config["auto_detect"] = False
    return pipeline.run(
        query=query,
        project_dir=project,
        dry_run=dry_run,
        business_mode=business_mode,
        density=density,
        motion=motion,
        review=review,
        review_file=review_file,
        output_version=output_version,
        from_version=from_version,
        pages=pages,
        history=history,
        output=output,
        content_file=content_file,
        llm_provider=llm_provider,
        llm_api_key=llm_api_key,
        llm_base_url=llm_base_url,
        llm_model=llm_model,
        image_mode=effective_image_mode,
        image_config=image_config,
        component_library=component_library,
    )


def _generate_ppt_beautify(
    input_pptx: str,
    style: str | None = None,
    output: str | None = None,
    density: int | None = None,
    motion: int | None = None,
    brand_spec=None,
    beautify_mode: str | None = None,
    component_library: str | None = None,
    patches: list[dict[str, Any]] | None = None,
) -> dict:
    import warnings
    warnings.warn(
        "Beautify pipeline is deprecated due to known quality issues. "
        "Use FreeStyle or Build mode instead.",
        DeprecationWarning,
        stacklevel=2,
    )
    from ppt_pro_max.enterprise.pipeline import EnterprisePipeline

    pipeline = EnterprisePipeline()
    return pipeline.run_beautify(
        input_pptx=input_pptx,
        output_path=output,
        style=style,
        brand_spec=brand_spec,
        density=density,
        motion=motion,
        beautify_mode=beautify_mode,
        component_library=component_library,
        patches=patches,
    )


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
