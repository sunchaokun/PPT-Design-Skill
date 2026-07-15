"""CLI entry point for PPT Design Skill."""

import argparse
import json
import sys

for _stream_name in ("stdout", "stderr"):
    _stream = getattr(sys, _stream_name)
    if _stream and hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

from ppt_pro_max import generate_ppt  # noqa: E402


def _load_dotenv():
    from pathlib import Path

    try:
        from dotenv import load_dotenv
    except ImportError:
        return

    candidates = []
    cwd = Path.cwd()
    candidates.append(cwd / ".env")
    pkg_dir = Path(__file__).resolve().parent
    for p in (pkg_dir, pkg_dir.parent, pkg_dir.parent.parent):
        candidates.append(p / ".env")
    home = Path.home()
    candidates.append(home / ".ppt-pro-max" / ".env")

    for env_path in candidates:
        if env_path.is_file():
            load_dotenv(env_path, override=False)
            return


def main():
    _load_dotenv()

    parser = argparse.ArgumentParser(
        prog="ppt-design",
        description="AI-powered PPT generation — narrative-driven, design-intelligent, fully editable .pptx",
    )
    parser.add_argument("query", nargs="?", default=None, help='Presentation topic, e.g. "AI产品融资路演" (--history 模式下可省略)')
    parser.add_argument("--strategy", help="Override presentation strategy")
    parser.add_argument("--theme", help="Theme preset name (backward compatible)")
    parser.add_argument("--style", help='Natural language style, e.g. "warm fintech pitch" or "dark cyberpunk"')
    parser.add_argument("--palette", help="Color palette name (25+ options: ocean-blue, cyber-neon, golden-luxury, ...)")
    parser.add_argument("--fonts", help="Font pair name (20+ options: modern-sans, serif-editorial, tech-mono, ...)")
    parser.add_argument("--decoration", help="Decoration style (10+ options: accent-bar, neon-lines, gold-trim, ...)")
    parser.add_argument("--layout-variant", help="Layout variant (8+ options: sidebar-left, centered, grid-2x2, ...)")
    parser.add_argument("--mood", help="Mood hint (professional, tech, warm, elegant, vibrant, nature, ...)")
    parser.add_argument("--style-seed", type=int, help="Random seed for reproducible style combinations")
    parser.add_argument("--slides", type=int, help="Override slide count")
    parser.add_argument("--content", help="JSON file with real content")
    parser.add_argument("--variance", type=int, choices=range(1, 11), help="Design variance 1-10")
    parser.add_argument("--motion", type=int, choices=range(1, 11), help="Animation intensity 1-10")
    parser.add_argument("--density", type=int, choices=range(1, 11), help="Content density 1-10")

    img_group = parser.add_argument_group("image options")
    img_group.add_argument(
        "--image-mode",
        choices=["placeholder", "search", "generate", "enhance"],
        default="placeholder",
        help="Image mode: placeholder (default), search (Unsplash/Pexels), generate (DALL-E/Wanx), enhance (Kimi K2.6 enhance + search)",
    )
    img_group.add_argument("--fetch-images", action="store_true", help="Shortcut for --image-mode search")
    img_group.add_argument("--unsplash-key", help="Unsplash API access key (or set UNSPLASH_ACCESS_KEY)")
    img_group.add_argument("--pexels-key", help="Pexels API key (or set PEXELS_API_KEY)")
    img_group.add_argument("--llm-provider", choices=["seedream", "doubao", "volcengine", "gpt-image", "dalle", "openai", "wanx", "tongyi", "aliyun", "kimi", "moonshot"], help="LLM image provider")
    img_group.add_argument("--llm-api-key", help="LLM API key (or set PPT_IMAGE_LLM_API_KEY)")
    img_group.add_argument("--llm-base-url", help="LLM API base URL override")
    img_group.add_argument("--llm-model", help="LLM model name override")

    parser.add_argument("--persist", action="store_true", help="Persist design system as MASTER.md")
    parser.add_argument("--dry-run", action="store_true", help="Output design decisions only")
    parser.add_argument("-o", "--output", help="Output .pptx file path")

    ent_group = parser.add_argument_group("enterprise options")
    ent_group.add_argument("--project", help="Project directory (triggers Enterprise Pipeline)")
    ent_group.add_argument("--business-mode", choices=["pitch", "education", "training", "report"], help="Business mode for Enterprise Pipeline")
    ent_group.add_argument("--review", action="store_true", help="Enable review gate (show plan before generating)")
    ent_group.add_argument("--review-file", help="Output review plan to JSON file (requires --review)")
    version_group = ent_group.add_mutually_exclusive_group()
    version_group.add_argument("--output-version", type=int, metavar="N", help="Specify output version number (overwrite existing)")
    version_group.add_argument("--from-version", type=int, metavar="N", help="Revise based on specified version's meta.json context")
    ent_group.add_argument("--pages", help="Page-level operations (e.g. 3,5 +6 -8 3>5 3<>7)")
    ent_group.add_argument("--history", action="store_true", help="Show version history (query optional)")
    ent_group.add_argument("--beautify", help="Beautify an existing PPT file (path to .pptx)")
    ent_group.add_argument("--beautify-mode", choices=["light", "full"], help="Beautify mode: light (color/font swap) or full (rebuild with components)")
    ent_group.add_argument("--component-library", help="Path to component library index.db")
    ent_group.add_argument("--component-catalog", action="store_true", help="Print component library catalog and exit")
    ent_group.add_argument("--component-search-type", help="Search components by type (group/smartart)")
    ent_group.add_argument("--component-search-category", help="Search components by category (process/swot/...)")
    ent_group.add_argument("--component-search-nodes", type=int, help="Search components by node count")

    args = parser.parse_args()

    if not args.history and not args.beautify and not args.component_catalog and args.query is None:
        parser.error("query is required unless --history is specified")
    if args.pages and not args.project and not args.history:
        parser.error("--pages requires --project")
    if args.review_file and not args.review:
        parser.error("--review-file requires --review")

    image_config = {}
    if args.unsplash_key:
        image_config["unsplash_access_key"] = args.unsplash_key
    if args.pexels_key:
        image_config["pexels_api_key"] = args.pexels_key
    if args.llm_provider:
        image_config["llm_provider"] = args.llm_provider
    if args.llm_api_key:
        image_config["llm_api_key"] = args.llm_api_key
    if args.llm_base_url:
        image_config["llm_base_url"] = args.llm_base_url
    if args.llm_model:
        image_config["llm_model"] = args.llm_model

    image_mode = args.image_mode

    if args.component_catalog or args.component_search_type:
        from ppt_pro_max import query_component_library
        if args.component_search_type:
            results = query_component_library(
                type=args.component_search_type,
                category=args.component_search_category,
                node_count=args.component_search_nodes,
                component_library=args.component_library,
            )
            if isinstance(results, list):
                for r in results:
                    print(f"  id={r.get('id')} type={r.get('type')} category={r.get('category')} "
                          f"variant={r.get('variant','')} nodes={r.get('node_count')}")
            else:
                print("No results")
        else:
            catalog = query_component_library(component_library=args.component_library)
            if isinstance(catalog, dict):
                for comp_type, categories in catalog.items():
                    print(f"\n{comp_type}:")
                    for cat, info in categories.items():
                        print(f"  {cat}: count={info['count']}, nodes={info['min_nodes']}-{info['max_nodes']}")
            else:
                print("No component library found")
        return

    try:
        result = generate_ppt(
            query=args.query or "",
            strategy=args.strategy,
            theme=args.theme,
            style=args.style,
            palette=args.palette,
            fonts=args.fonts,
            decoration=args.decoration,
            layout_variant=args.layout_variant,
            mood=args.mood,
            style_seed=args.style_seed,
            slides=args.slides,
            content_file=args.content,
            variance=args.variance,
            motion=args.motion,
            density=args.density,
            fetch_images=args.fetch_images,
            image_mode=image_mode,
            image_config=image_config if image_config else None,
            persist=args.persist,
            dry_run=args.dry_run,
            output=args.output,
            project=args.project,
            business_mode=args.business_mode,
            review=args.review,
            review_file=args.review_file,
            output_version=args.output_version,
            from_version=args.from_version,
            pages=args.pages,
            history=args.history,
            beautify=args.beautify,
            beautify_mode=args.beautify_mode,
            component_library=args.component_library,
        )
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    if result.get("dry_run"):
        print(json.dumps(result, indent=2, ensure_ascii=False))
    elif result.get("history"):
        for v in result.get("versions", []):
            meta = v.get("meta") or {}
            print(f"v{v['version']} | {meta.get('created_at', 'N/A')} | {meta.get('page_count', meta.get('num_slides', '?'))}页 | {meta.get('business_mode', '?')} | {meta.get('strategy', '?')}")
    elif result.get("review"):
        proposal = result.get("proposal", {})
        from ppt_pro_max.enterprise.review_gate import ReviewGate
        gate = ReviewGate()
        display = gate.format_cli(proposal)
        print(display)
        try:
            answer = input().strip().lower()
        except (EOFError, KeyboardInterrupt):
            answer = "n"
        if answer in ("y", "yes"):
            print("Proceeding with generation...")
            result = generate_ppt(
                query=args.query or "",
                strategy=args.strategy,
                theme=args.theme,
                style=args.style,
                palette=args.palette,
                fonts=args.fonts,
                decoration=args.decoration,
                layout_variant=args.layout_variant,
                mood=args.mood,
                style_seed=args.style_seed,
                slides=args.slides,
                content_file=args.content,
                variance=args.variance,
                motion=args.motion,
                density=args.density,
                fetch_images=args.fetch_images,
                image_mode=image_mode,
                image_config=image_config if image_config else None,
                persist=args.persist,
                dry_run=False,
                output=args.output,
                project=args.project,
                business_mode=args.business_mode,
                review=False,
                output_version=args.output_version,
                from_version=args.from_version,
                pages=args.pages,
                history=False,
            )
            if result.get("error"):
                print(f"Error: {result['error']}", file=sys.stderr)
                sys.exit(1)
            print(f"Generated: {result['output_path']}")
            print(f"Pages: {result.get('page_count', result.get('num_slides', '?'))}")
            if result.get("version"):
                print(f"Version: v{result['version']}")
        else:
            print(f"Review proposal saved: {result['proposal_path']}")
            print("Cancelled. Edit proposal and re-run without --review to apply.")
    elif result.get("error"):
        print(f"Error: {result['error']}", file=sys.stderr)
        sys.exit(1)
    else:
        print(f"Generated: {result['output_path']}")
        print(f"Pages: {result.get('page_count', result.get('num_slides', '?'))}")
        if "strategy" in result:
            print(f"Strategy: {result['strategy']}")
        if result.get("theme"):
            print(f"Theme: {result['theme']}")
        if result.get("theme_atoms"):
            atoms = result["theme_atoms"]
            print(f"Style: palette={atoms.get('palette')}, fonts={atoms.get('fonts')}, decoration={atoms.get('decoration')}, layout={atoms.get('layout')}")
        if result.get("mode"):
            print(f"Mode: {result['mode']}")
        if result.get("version"):
            print(f"Version: v{result['version']}")


if __name__ == "__main__":
    main()
