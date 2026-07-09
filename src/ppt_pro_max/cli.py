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
    img_group.add_argument("--llm-provider", choices=["dalle", "openai", "wanx", "tongyi", "aliyun", "kimi", "moonshot"], help="LLM image provider")
    img_group.add_argument("--llm-api-key", help="LLM API key (or set PPT_IMAGE_LLM_API_KEY)")
    img_group.add_argument("--llm-base-url", help="LLM API base URL override")
    img_group.add_argument("--llm-model", help="LLM model name override")

    parser.add_argument("--persist", action="store_true", help="Persist design system as MASTER.md")
    parser.add_argument("--dry-run", action="store_true", help="Output design decisions only")
    parser.add_argument("-o", "--output", help="Output .pptx file path")

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

    image_mode = args.image_mode
    if args.fetch_images and image_mode == "placeholder":
        image_mode = "search"

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
            image_mode=image_mode,
            image_config=image_config if image_config else None,
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
