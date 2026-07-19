# AGENTS.md

## Project: PPT Design Skill

AI-powered PPT generation — 3 modes: FreeStyle (one-liner), Enterprise (brand compliance), Build (pixel-perfect). 40,000+ style combos, fully editable .pptx.

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

### Enterprise Pipeline (P1-P8 unified)

- `src/ppt_pro_max/enterprise/pipeline.py` — Orchestration: template-clone path (when visual DNA exists) or PrecisionRenderer (no template); `run_beautify()` for beautify mode; `_run_template_clone()` for clone+expand+DeliveryGate
- `src/ppt_pro_max/enterprise/precision_renderer.py` — Unified renderer: `render_slide()` dispatches by goal (hook→hero, content→bullets, features→cards, data→chart, code→code-block, exercise→exercise, overview→sidebar)
- `src/ppt_pro_max/enterprise/scanner.py` — Scans project dir for template.pptx, brand.json, content.json, README.md, images
- `src/ppt_pro_max/enterprise/content_parser.py` — Parses content.json AND README.md (P4: H1→pages, H2→bullets, code blocks, tables, images, goal inference with English+Chinese keywords)
- `src/ppt_pro_max/enterprise/image_matcher.py` — Image assignment: keyword-based `match_images()` + size-aware `assign_images_by_size()` (P5: >1500px→background, 800-1500→scene, <800→icon) + `auto_generate_image_prompts()` for AI image fetcher
- `src/ppt_pro_max/enterprise/proposal_generator.py` — P6: Generate 2-3 style preview PPTs (4 slides each: hook/problem/features/cta) with differentiated palettes/moods
- `src/ppt_pro_max/enterprise/slide_extractor.py` — P9: Extract content + layout from existing PPT (for beautify mode)
- `src/ppt_pro_max/enterprise/smartart_extractor.py` — P12: SmartArt XML parsing (data/layout/colors/quickStyle, no drawing needed)
- `src/ppt_pro_max/enterprise/group_extractor.py` — P12: GroupShape recursive extraction (texts/images/structure)
- `src/ppt_pro_max/enterprise/ole_extractor.py` — P12: OLE/embedded object metadata extraction
- `src/ppt_pro_max/enterprise/component_library.py` — P13: SQLite-indexed component library (search/match/add/bulk_import/checksum dedup)
- `src/ppt_pro_max/enterprise/component_renderer.py` — P14: Component rendering bridge (match→fill data→apply brand colors→inject to slide)
- `src/ppt_pro_max/enterprise/design_dna_extractor.py` — Template DNA extraction + clone_and_patch + expand_and_patch + plan_pages + content zone rerender
- `src/ppt_pro_max/enterprise/delivery_gate.py` — Quality check (13 items) + auto-fix + report formatting
- `src/ppt_pro_max/renderer/theme_composer.py` — 35 moods (P3: added international/cream/frosted/mckinsey/consulting/pastel/retro/government/legal/pharma/realestate/automotive/aviation/energy/telecom/logistics)

## Prerequisites

- Python 3.x
- **ui-ux-pro-max skill** is a **required** dependency (not optional):
  ```bash
  npx ui-ux-pro-max-cli init --ai opencode
  ```
  If not found, the system raises `UiUxProMaxNotFoundError` at runtime.
  Install location is auto-detected (see `adapters/ui_ux_adapter.py` search order).
  You can also set `UX_PRO_MAX_DIR` env var to explicitly specify the path.
- **Pillow >= 10.0** required for cover-fit image cropping

## Key Constraints

- **python-pptx 1.0.2**: `PP_TRANSITION_TYPE` does NOT exist, must use XML for transitions
- **Cover-fit images**: Use `_add_picture_cover()` which Pillow pre-crops — never use `add_picture` with stretch
- **Cache-first**: All image engines check cache before API call to prevent duplicate charges
- **Source of truth**: `src/ppt_pro_max/` — never modify ui-ux-pro-max
- **Windows**: Use `python` not `python3`
- **Pipeline unified (P2)**: Pipeline renders via PrecisionRenderer.render_slide() when no template; uses clone_and_patch when template has visual DNA
- **Content priority**: content.json > README.md > StoryPlanner (P4 integration)
- **Template-clone path**: When template.pptx has visual DNA → extract DNA → plan_pages → expand_and_patch → DeliveryGate → deliver (never re-render from scratch)
- **Composite titles**: Multi-textbox titles (e.g., "乡/村/振/兴") — replace first zone, clear rest
- **Content zone rerender**: Zone mismatch → keep shell (background+decorations+title), clear content area, PrecisionRenderer rerenders with DNA brand_spec
- **layout field**: bullets/cards/process/hierarchy/timeline/swot/table/code/chart/auto — controls content zone rendering
- **DeliveryGate**: 13 checks, fatal items must be zero before delivery, auto-fix up to 2 rounds
- **Image assignment flow**: match_images() → assign_images_by_size() → auto_generate_image_prompts() → ImageFetcher (P5 integration)
- **Proposal flow**: `generate_ppt(proposal=True)` → 3 preview PPTs → user picks → `generate_ppt(confirmed_proposal="B")` (P6-P7)
- **Beautify flow**: `generate_ppt(beautify="client.pptx", style="professional")` → SlideExtractor → PrecisionRenderer (P9-P10)
- **SmartArt storage**: 4 XML parts only (data/layout/colors/quickStyle), no drawing (PowerPoint auto-rebuilds), colors.xml must store original (P12)
- **Component library**: SQLite index + filesystem storage, checksum dedup, match engine (P13)

## Style System

- 25 color palettes × 20 font pairs × 10 decorations × 8 layout variants = 40,000 combos
- Natural language: `--style "warm fintech"` auto-selects matching atoms
- Exact control: `--palette wine-burgundy --fonts elegant-serif`
- Presets backward compatible: professional, dark-tech, warm-elegant, vibrant-startup, nature-calm
- 35 mood categories (P3): professional, tech, dark, warm, elegant, luxury, vibrant, startup, nature, calm, minimal, bold, fresh, industrial, fintech, health, education, sustainability, creative, international, cream, frosted, mckinsey, consulting, pastel, retro, government, legal, pharma, realestate, automotive, aviation, energy, telecom, logistics

## P1-P14 Implementation Status

| Phase | Feature | Status | Tests |
|-------|---------|--------|-------|
| P1 | PrecisionRenderer.render_slide() | Done | 13 in test_render_slide.py |
| P2 | Pipeline unified PrecisionRenderer | Done | 6 in test_pipeline_unified.py + 34 migrated tests |
| P3 | mood_words expansion (35 moods) | Done | 27 in test_mood_words_expansion.py |
| P4 | ContentParser README.md parsing | Done | 24 in test_readme_parsing.py + 3 integration |
| P5 | Image size classification + image_prompt | Done | 22 in test_image_size_prompt.py + 4 integration |
| P6 | ProposalGenerator (2-3 style previews) | Done | 12 in test_proposal_generator.py |
| P7 | generate_ppt() API (proposal/confirmed_proposal/materials_dir) | Done | 10 in test_generate_ppt_api.py |
| P8 | End-to-end tests (P1-P7) | Done | 17 in test_e2e_p1_p7.py |
| P9 | SlideExtractor + beautify rendering | Done | 17 in test_slide_extractor.py |
| P10 | beautify API + CLI integration | Done | 5 in test_beautify_e2e.py (API tests) |
| P11 | Beautify mode end-to-end tests | Done | 6 in test_beautify_e2e.py (E2E tests) |
| P12 | SmartArt/GroupShape/OLE XML extractors | Done | 8 in test_p9_p14.py |
| P13 | ComponentLibrary (Schema + Index + CRUD + match) | Done | 10 in test_p9_p14.py |
| P14 | ComponentRenderer (match→fill→brand→inject) | Done | 3 in test_p9_p14.py |
| P15 | Component Library Integration (catalog+API+content fields+renderer+beautify) | Done | 38 in test_component_integration.py |
| DQ | Design Quality Upgrades (Tier 1+2+3, 28 upgrades) | Done | 95 in test_design_quality.py + 25 in test_design_integration.py |
| TC | Template Clone Pipeline (expand+plan+delivery+rerender) | Done | 29 in test_template_clone_pipeline.py |

## End-to-End Evaluation

**824 tests passed, 5 skipped, 0 failures. Lint clean on all modified files.**

### Test Scenarios & Results

| ID | Scenario | Slides | Shapes | Images | Min Font | Max Font | File Size |
|----|----------|--------|--------|--------|----------|----------|-----------|
| E1 | Freestyle dark cyberpunk | 6 | 33 | 4 | 16pt | 44pt | 52.8 KB |
| E2 | Freestyle warm elegant | 8 | 43 | 5 | 14pt | 44pt | 59.9 KB |
| E3 | Enterprise full (template+brand+content+images) | 8 | 57 | 7 | 11pt | 52pt | 39.5 KB |
| E4 | Enterprise README.md only | 6 | 31 | 6 | 11pt | 52pt | 43.4 KB |
| E5A | Proposal A (mckinsey) | 4 | 27 | 0 | 11pt | 52pt | 31.8 KB |
| E5B | Proposal B (alt palette) | 4 | 27 | 0 | 11pt | 52pt | 31.8 KB |
| E5C | Proposal C (alt mood) | 4 | 27 | 0 | 11pt | 52pt | 31.8 KB |

### Content Verification (E3 Enterprise — all 8 goal types)

| Slide | Goal | Content Rendered |
|-------|------|-----------------|
| 0 | hook | Title + subtitle + hero image |
| 1 | problem | Title + 4 bullets + image |
| 2 | solution | Title + 4 bullets + image |
| 3 | features | Title + 3 cards (AI Engine, Live Dashboard, Integration) |
| 4 | data | Title + table diagram (5 rows) |
| 5 | code | Title + code block (python, 4 lines) |
| 6 | exercise | Title + badge "Exercise 5 min" + 3 steps |
| 7 | cta | Title + subtitle |

### Quality Checks

- All font sizes >= 11pt (no unreadable text)
- All slides have >= 3 shapes (no blank slides)
- All slides have text content (no empty slides)
- Images correctly placed: hero→hook, product→features
- README.md correctly parsed into 6 pages with goal inference
- Proposals differentiated: A=indigo-deep, B=slate-minimal, C=lavender-dream

### Build Mode Evaluation (v0.8.0)

| ID | Scenario | Slides | Shapes | Colors | Fonts | File Size |
|----|----------|--------|--------|--------|-------|-----------|
| BA | McKinsey (sidebar+table+numbered cards) | 8 | 206 | 8 | Georgia+Calibri+Consolas | 40.9 KB |
| BB | Cyberpunk (grid+terminal+dashboard) | 8 | 272 | 7 | Orbitron+JetBrains Mono | 42.7 KB |
| BC | Creative (circles+emoji+before-after) | 8 | 190 | 9 | Fredoka+Nunito | 41.2 KB |

**Key difference from FreeStyle**: Each Build proposal has completely different page structure, layout strategy, and visual language — not just palette/font swap.
