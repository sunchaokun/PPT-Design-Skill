#!/usr/bin/env python3
"""PPT Design Skill — CLI entry point for AI coding assistants.

Usage:
  python generate_ppt.py "AI startup investor pitch" [--style "dark cyberpunk"] [--fetch-images]
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
        print('Usage: python generate_ppt.py "topic" [--style "warm fintech"] [--fetch-images]')
        sys.exit(1)

    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("query")
    parser.add_argument("--style")
    parser.add_argument("--palette")
    parser.add_argument("--fonts")
    parser.add_argument("--decoration")
    parser.add_argument("--layout-variant")
    parser.add_argument("--mood")
    parser.add_argument("--theme")
    parser.add_argument("--strategy")
    parser.add_argument("--slides", type=int)
    parser.add_argument("--content")
    parser.add_argument("--fetch-images", action="store_true")
    parser.add_argument("--llm-provider")
    parser.add_argument("--llm-api-key")
    parser.add_argument("--llm-model")
    parser.add_argument("--variance", type=int)
    parser.add_argument("--motion", type=int)
    parser.add_argument("--density", type=int)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("-o", "--output")

    args = parser.parse_args()

    result = generate_ppt(
        query=args.query,
        style=args.style,
        palette=args.palette,
        fonts=args.fonts,
        decoration=args.decoration,
        layout_variant=args.layout_variant,
        mood=args.mood,
        theme=args.theme,
        strategy=args.strategy,
        slides=args.slides,
        content_file=args.content,
        fetch_images=args.fetch_images,
        llm_provider=args.llm_provider,
        llm_api_key=args.llm_api_key,
        llm_model=args.llm_model,
        variance=args.variance,
        motion=args.motion,
        density=args.density,
        dry_run=args.dry_run,
        output=args.output,
    )

    if result.get("dry_run"):
        import json
        print(json.dumps(result, indent=2, ensure_ascii=False))
    else:
        print(f"Generated: {result['output_path']}")
        print(f"Pages: {result['page_count']}")
        if result.get("theme_atoms"):
            a = result["theme_atoms"]
            print(f"Style: palette={a.get('palette')}, fonts={a.get('fonts')}, deco={a.get('decoration')}, layout={a.get('layout')}")

if __name__ == "__main__":
    main()
