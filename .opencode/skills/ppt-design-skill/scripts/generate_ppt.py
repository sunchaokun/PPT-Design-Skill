#!/usr/bin/env python3
"""PPT Design Skill — CLI entry point for AI coding assistants.

Usage:
  python generate_ppt.py "AI startup investor pitch" [--style "dark cyberpunk"] [--fetch-images]
  python generate_ppt.py "路演" --project ./my-project --density 6
  python generate_ppt.py "" --project ./my-project --pages "-3,2<>5"
  python generate_ppt.py "" --project ./my-project --history
"""

import sys
import os

pkg_src = os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "src")
pkg_src = os.path.normpath(pkg_src)
if os.path.isdir(pkg_src):
    sys.path.insert(0, pkg_src)

from ppt_pro_max import generate_ppt

def main():
    if len(sys.argv) < 2:
        print('Usage: python generate_ppt.py "topic" [--style "warm fintech"] [--fetch-images] [--project DIR]')
        sys.exit(1)

    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("query", nargs="?", default=None)
    parser.add_argument("--style")
    parser.add_argument("--palette")
    parser.add_argument("--fonts")
    parser.add_argument("--decoration")
    parser.add_argument("--layout-variant")
    parser.add_argument("--mood")
    parser.add_argument("--theme")
    parser.add_argument("--strategy")
    parser.add_argument("--style-seed", type=int)
    parser.add_argument("--slides", type=int)
    parser.add_argument("--content")
    parser.add_argument("--variance", type=int)
    parser.add_argument("--motion", type=int)
    parser.add_argument("--density", type=int)

    parser.add_argument("--image-mode", choices=["placeholder", "search", "generate", "enhance"])
    parser.add_argument("--fetch-images", action="store_true")
    parser.add_argument("--unsplash-key")
    parser.add_argument("--pexels-key")
    parser.add_argument("--llm-provider")
    parser.add_argument("--llm-api-key")
    parser.add_argument("--llm-base-url")
    parser.add_argument("--llm-model")

    parser.add_argument("--persist", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("-o", "--output")

    parser.add_argument("--project")
    parser.add_argument("--business-mode", choices=["pitch", "education", "training", "report"])
    parser.add_argument("--review", action="store_true")
    parser.add_argument("--review-file")
    parser.add_argument("--output-version", type=int, metavar="N")
    parser.add_argument("--from-version", type=int, metavar="N")
    parser.add_argument("--pages")
    parser.add_argument("--history", action="store_true")

    args = parser.parse_args()

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

    result = generate_ppt(
        query=args.query or "",
        style=args.style,
        palette=args.palette,
        fonts=args.fonts,
        decoration=args.decoration,
        layout_variant=args.layout_variant,
        mood=args.mood,
        theme=args.theme,
        strategy=args.strategy,
        style_seed=args.style_seed,
        slides=args.slides,
        content_file=args.content,
        variance=args.variance,
        motion=args.motion,
        density=args.density,
        fetch_images=args.fetch_images,
        image_mode=args.image_mode,
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
    )

    if result.get("dry_run"):
        import json
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
    elif result.get("error"):
        print(f"Error: {result['error']}", file=sys.stderr)
        sys.exit(1)
    else:
        print(f"Generated: {result['output_path']}")
        print(f"Pages: {result.get('page_count', result.get('num_slides', '?'))}")
        if result.get("theme_atoms"):
            a = result["theme_atoms"]
            print(f"Style: palette={a.get('palette')}, fonts={a.get('fonts')}, deco={a.get('decoration')}, layout={a.get('layout')}")
        if result.get("version"):
            print(f"Version: v{result['version']}")

if __name__ == "__main__":
    main()
