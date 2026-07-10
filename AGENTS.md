# AGENTS.md

## Project: PPT Design Skill

AI-powered PPT generation — 40,000+ style combinations, narrative-driven, fully editable .pptx.

## Commands

### Generate PPT
```bash
python -m ppt_pro_max "AI startup investor pitch" --style "dark cyberpunk" --fetch-images --llm-provider seedream --llm-api-key $ARK_API_KEY
```

### Run Tests
```bash
python -m pytest tests/ -q
```

### Lint
```bash
python -m ruff check src/
```

## Architecture

4-phase pipeline: StoryPlanner → DesignDecider → ContentGenerator → PPTRenderer

- `src/ppt_pro_max/__init__.py` — `generate_ppt()` API (entry point)
- `src/ppt_pro_max/planner/story_planner.py` — Phase 1: narrative planning
- `src/ppt_pro_max/decider/design_decider.py` — Phase 2: design decisions
- `src/ppt_pro_max/content/content_generator.py` — Phase 3: content generation
- `src/ppt_pro_max/renderer/ppt_renderer.py` — Phase 4: PPT rendering
- `src/ppt_pro_max/renderer/theme_composer.py` — 40,000+ style combinations
- `src/ppt_pro_max/renderer/image_fetcher.py` — 4 image generation engines (Seedream/GPT Image/DALL-E/Wanx) + 1 enhancer (Kimi)

## Key Constraints

- **python-pptx 1.0.2**: `PP_TRANSITION_TYPE` does NOT exist, must use XML for transitions
- **Cover-fit images**: Use `_add_picture_cover()` which Pillow pre-crops — never use `add_picture` with stretch
- **Cache-first**: All image engines check cache before API call to prevent duplicate charges
- **Source of truth**: `src/ppt_pro_max/` — never modify ui-ux-pro-max
- **Windows**: Use `python` not `python3`

## Style System

- 25 color palettes × 20 font pairs × 10 decorations × 8 layout variants = 40,000 combos
- Natural language: `--style "warm fintech"` auto-selects matching atoms
- Exact control: `--palette wine-burgundy --fonts elegant-serif`
- Presets backward compatible: professional, dark-tech, warm-elegant, vibrant-startup, nature-calm
