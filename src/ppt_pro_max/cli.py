"""CLI entry point for PPT Design Skill."""

import argparse
import json
import sys

from ppt_pro_max import generate_ppt


def main():
    parser = argparse.ArgumentParser(
        prog="ppt-design",
        description="AI-powered PPT generation — narrative-driven, design-intelligent, fully editable .pptx",
    )
    parser.add_argument("query", help='Presentation topic, e.g. "AI产品融资路演"')
    parser.add_argument("--strategy", help="Override presentation strategy")
    parser.add_argument("--theme", help="Theme preset name")
    parser.add_argument("--slides", type=int, help="Override slide count")
    parser.add_argument("--content", help="JSON file with real content")
    parser.add_argument("--variance", type=int, choices=range(1, 11), help="Design variance 1-10")
    parser.add_argument("--motion", type=int, choices=range(1, 11), help="Animation intensity 1-10")
    parser.add_argument("--density", type=int, choices=range(1, 11), help="Content density 1-10")
    parser.add_argument("--fetch-images", action="store_true", help="Auto-download images from Unsplash/Pexels")
    parser.add_argument("--persist", action="store_true", help="Persist design system as MASTER.md")
    parser.add_argument("--dry-run", action="store_true", help="Output design decisions only")
    parser.add_argument("-o", "--output", help="Output .pptx file path")

    args = parser.parse_args()

    try:
        result = generate_ppt(
            query=args.query,
            strategy=args.strategy,
            theme=args.theme,
            slides=args.slides,
            content_file=args.content,
            variance=args.variance,
            motion=args.motion,
            density=args.density,
            fetch_images=args.fetch_images,
            persist=args.persist,
            dry_run=args.dry_run,
            output=args.output,
        )
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    if result.get("dry_run"):
        print(json.dumps(result, indent=2, ensure_ascii=False))
    else:
        print(f"Generated: {result['output_path']}")
        print(f"Pages: {result['page_count']}")
        print(f"Strategy: {result['strategy']}")
        print(f"Theme: {result.get('theme', 'default')}")


if __name__ == "__main__":
    main()
