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

from ppt_pro_max import generate_ppt, fetch_image, extract_design_dna  # noqa: E402


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


def _add_image_options(parser):
    img = parser.add_argument_group("image options")
    img.add_argument(
        "--image-mode",
        choices=["placeholder", "search", "generate", "enhance", "auto"],
        default="auto",
        help="Image mode: placeholder, search (Unsplash/Pexels), generate (AI), enhance (Kimi+search), auto (generate→search)",
    )
    img.add_argument("--fetch-images", action="store_true", help="Shortcut for --image-mode search")
    img.add_argument("--unsplash-key", help="Unsplash API access key (or set UNSPLASH_ACCESS_KEY)")
    img.add_argument("--pexels-key", help="Pexels API key (or set PEXELS_API_KEY)")
    img.add_argument("--llm-provider", choices=["seedream", "doubao", "volcengine", "gpt-image", "dalle", "openai", "wanx", "tongyi", "aliyun", "kimi", "moonshot", "gemini", "google"], help="LLM image provider")
    img.add_argument("--llm-api-key", help="LLM API key (or set PPT_IMAGE_LLM_API_KEY)")
    img.add_argument("--llm-base-url", help="LLM API base URL override")
    img.add_argument("--llm-model", help="LLM model name override")
    img.add_argument("--no-auto-detect", action="store_true", help="Disable auto-detection of LLM config from host tools")


def _build_image_config(args):
    cfg = {}
    if getattr(args, "unsplash_key", None):
        cfg["unsplash_access_key"] = args.unsplash_key
    if getattr(args, "pexels_key", None):
        cfg["pexels_api_key"] = args.pexels_key
    if getattr(args, "llm_provider", None):
        cfg["llm_provider"] = args.llm_provider
    if getattr(args, "llm_api_key", None):
        cfg["llm_api_key"] = args.llm_api_key
    if getattr(args, "llm_base_url", None):
        cfg["llm_base_url"] = args.llm_base_url
    if getattr(args, "llm_model", None):
        cfg["llm_model"] = args.llm_model
    if getattr(args, "no_auto_detect", False):
        cfg["auto_detect"] = False
    return cfg


def _cmd_image(args):
    mode = args.image_mode
    if args.fetch_images:
        mode = "search"

    result = fetch_image(
        keywords=args.keywords,
        mode=mode,
        emotion=args.emotion or "",
        goal=args.goal or "",
        width=args.width,
        height=args.height,
        llm_provider=args.llm_provider,
        llm_api_key=args.llm_api_key,
        llm_base_url=args.llm_base_url,
        llm_model=args.llm_model,
        unsplash_access_key=args.unsplash_key,
        pexels_api_key=args.pexels_key,
        auto_detect=not args.no_auto_detect,
    )

    if result["path"]:
        print(result["path"])
    else:
        print("No image found", file=sys.stderr)
        sys.exit(1)

    if args.verbose:
        for k, v in result.items():
            if k != "path":
                print(f"  {k}: {v}")


def _run_image_command(argv):
    parser = argparse.ArgumentParser(
        prog="ppt-design image",
        description="Generate or fetch images independently using Seedream / GPT Image / DALL-E / Gemini / Wanx / Unsplash / Pexels",
    )
    parser.add_argument("keywords", help='Image keywords, e.g. "AI startup futuristic city"')
    parser.add_argument("--emotion", help="Emotion hint: curiosity, hope, confidence, warmth, urgency, ...")
    parser.add_argument("--goal", help="Slide goal: hook, problem, solution, features, cta, ...")
    parser.add_argument("--width", type=int, default=1920, help="Image width in pixels (default: 1920)")
    parser.add_argument("--height", type=int, default=1080, help="Image height in pixels (default: 1080)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Print extra info (mode, provider, etc.)")
    _add_image_options(parser)
    args = parser.parse_args(argv)
    _cmd_image(args)


def _cmd_analyze(args):
    try:
        dna = extract_design_dna(args.pptx_path)
        print(json.dumps(dna, indent=2, ensure_ascii=False))
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


def _run_analyze_command(argv):
    parser = argparse.ArgumentParser(
        prog="ppt-design analyze",
        description="Extract design DNA from an existing .pptx file",
    )
    parser.add_argument("pptx_path", help="Path to .pptx file to analyze")
    args = parser.parse_args(argv)
    _cmd_analyze(args)


def main():
    _load_dotenv()

    if len(sys.argv) > 1 and sys.argv[1] == "image":
        _run_image_command(sys.argv[2:])
        return

    if len(sys.argv) > 1 and sys.argv[1] == "analyze":
        _run_analyze_command(sys.argv[2:])
        return

    parser = argparse.ArgumentParser(
        prog="ppt-design",
        description="AI-powered PPT generation — narrative-driven, design-intelligent, fully editable .pptx",
    )
    parser.add_argument("query", nargs="?", default=None, help='Presentation topic, e.g. "AI产品融资路演"')
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
    parser.add_argument("--content", dest="content_file", help="JSON file with real content")
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
    img_group.add_argument("--llm-provider", choices=["seedream", "doubao", "volcengine", "gpt-image", "dalle", "openai", "wanx", "tongyi", "aliyun", "kimi", "moonshot", "gemini", "google"], help="LLM image provider")
    img_group.add_argument("--llm-api-key", help="LLM API key (or set PPT_IMAGE_LLM_API_KEY)")
    img_group.add_argument("--llm-base-url", help="LLM API base URL override")
    img_group.add_argument("--llm-model", help="LLM model name override")
    img_group.add_argument("--no-auto-detect", action="store_true", help="Disable auto-detection of LLM config from host tools")

    proposal_group = parser.add_argument_group("proposal options")
    proposal_group.add_argument("--proposal", action="store_true", help="Generate 2-3 style preview PPTs instead of full PPT")
    proposal_group.add_argument("--confirmed-proposal", help="Resume from a confirmed proposal (e.g. 'A', 'B', 'C')")
    proposal_group.add_argument("--materials-dir", help="Project materials directory for proposals")

    parser.add_argument("--persist", action="store_true", help="Persist design system as MASTER.md")
    parser.add_argument("--dry-run", action="store_true", help="Output design decisions only")
    parser.add_argument("-o", "--output", help="Output .pptx file path")

    args = parser.parse_args()

    if args.query is None and not args.proposal:
        parser.error("query is required unless --proposal is specified")

    image_config = _build_image_config(args)

    image_mode = args.image_mode
    if args.fetch_images:
        image_mode = "search"

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
            content_file=args.content_file,
            variance=args.variance,
            motion=args.motion,
            density=args.density,
            fetch_images=args.fetch_images,
            image_mode=image_mode,
            image_config=image_config if image_config else None,
            llm_provider=args.llm_provider,
            llm_api_key=args.llm_api_key,
            llm_base_url=args.llm_base_url,
            llm_model=args.llm_model,
            persist=args.persist,
            dry_run=args.dry_run,
            output=args.output,
            proposal=args.proposal,
            confirmed_proposal=args.confirmed_proposal,
            materials_dir=args.materials_dir,
            auto_detect=not args.no_auto_detect,
        )
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    if result.get("dry_run"):
        print(json.dumps(result, indent=2, ensure_ascii=False))
    elif result.get("proposals"):
        proposals = result["proposals"]
        for p in proposals:
            label = p.get("label", "?")
            path = p.get("path", "")
            print(f"Proposal {label}: {path}")
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


if __name__ == "__main__":
    main()
