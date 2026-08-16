# AGENTS.md

## Project: PPT Design Skill

AI-powered PPT generation — 3 modes: FreeStyle (one-liner), Enterprise (brand compliance), Build (pixel-perfect). 40,000+ style combos, fully editable .pptx.

**⚠️ python-pptx full API reference: [`src/ppt_pro_max/docs/python-pptx-reference.md`](src/ppt_pro_max/docs/python-pptx-reference.md)** — 170+ shape types, 73 chart types, tables, connectors, freeform, hyperlinks, media, effects, 3D, OOXML. **Must read before writing python-pptx code.**

**⚠️ READ BEFORE deleting any file: [`skill/ARCHITECTURE.md`](skill/ARCHITECTURE.md)** — Module registry with risk levels (CORE/ENTRY/INTERNAL/ISOLATED). Run `python scripts/check_deps.py <module_name>` to check deletion impact.

## Commands

### Generate PPT
```bash
python -m ppt_pro_max "AI startup investor pitch" --style "dark cyberpunk" --fetch-images --llm-provider seedream --llm-api-key $ARK_API_KEY
```

### Render Preview (debug build.py layouts without opening PowerPoint)
```bash
# Render every slide to PNG + HTML contact sheet (PowerPoint COM / LibreOffice fallback)
python -m ppt_pro_max.render_preview output/build_10pages.pptx --open

# Codex sandbox / CI / headless env: PowerPoint COM fails (WinError 1312).
# Use the LibreOffice engine explicitly, or just call preview() in build.py —
# render_preview auto-falls-back PowerPoint → LibreOffice.
python -m ppt_pro_max.render_preview output/build_10pages.pptx --engine libreoffice
```

**Installer auto-installs render deps**: `python install.py` detects LibreOffice +
poppler (pdftoppm) and installs any missing via winget, so the headless render
path works right after installation. Add `--no-render-deps` to skip. Check
status with `python install.py --check`.

**In build.py — preferred (LLM-friendly):** `from ppt_pro_max.build_helpers import *`
```python
preview("output.pptx")                     # auto engine (COM → LibreOffice fallback)
preview("output.pptx", engine="libreoffice")  # Codex sandbox / headless
```
`preview()` renders every slide to PNG + an HTML contact sheet; Codex can then
visually review real layout renders without PowerPoint. Requires LibreOffice
(soffice.bin) + poppler (pdftoppm) for the headless path.

### Standalone Image Generation
```bash
# AI image generation (no PPT needed)
python -m ppt_pro_max image "futuristic AI city" --llm-provider seedream --llm-api-key $ARK_API_KEY

# Search stock photos
python -m ppt_pro_max image "team meeting" --image-mode search --unsplash-key $KEY

# Auto mode: AI generation → fall back to search
python -m ppt_pro_max image "product launch" --llm-provider seedream -v
```

Python API:
```python
from ppt_pro_max import fetch_image
result = fetch_image("futuristic AI city", mode="generate", llm_provider="seedream", llm_api_key="...")
print(result["path"])  # Local file path
```

### Run Tests
```bash
# 从项目根目录 (conftest.py 自动将 src/ 加入 sys.path, 确保加载 V2 源码)
python -m pytest tests/ -q

# 完整回归 (跳过 7 个已知数据依赖问题)
python -m pytest tests/ -q --ignore=tests/test_group_audit.py \
  --ignore=tests/test_image_fetcher.py --ignore=tests/test_pptx_capabilities.py \
  --ignore=tests/test_xml_extraction.py --ignore=tests/test_analyze_template.py
```

> **⚠️ 源码优先级**: 系统 site-packages 可能装有旧版 ppt_pro_max。pytest 由 conftest.py 保证加载 `src/`。但手动运行 `python build.py` 或 `python -m ppt_pro_max` 时, 必须 `pip install -e .` 或设 `$env:PYTHONPATH = "src"`, 否则会加载旧版代码。

### Lint
```bash
python -m ruff check src/
```

## Architecture

4-phase pipeline: StoryPlanner → DesignDecider → ContentGenerator → PPTRenderer

- `src/ppt_pro_max/__init__.py` — `generate_ppt()` + `fetch_image()` API (entry point)
- `src/ppt_pro_max/planner/story_planner.py` — Phase 1: narrative planning
- `src/ppt_pro_max/decider/design_decider.py` — Phase 2: design decisions
- `src/ppt_pro_max/content/content_generator.py` — Phase 3: content generation
- `src/ppt_pro_max/renderer/ppt_renderer.py` — Phase 4: PPT rendering
- `src/ppt_pro_max/renderer/theme_composer.py` — 40,000+ style combinations
- `src/ppt_pro_max/renderer/image_fetcher.py` — 4 image generation engines (Seedream/GPT Image/DALL-E/Wanx) + 1 enhancer (Kimi)
- `src/ppt_pro_max/render_preview.py` — Debug tool: batch-render .pptx slides to PNG + HTML contact sheet (PowerPoint COM / LibreOffice fallback)

### Enterprise / Renderer (unified)

- `src/ppt_pro_max/enterprise/precision_renderer.py` — Unified renderer: `render_slide()` dispatches by goal (hook→hero, content→bullets, features→cards, data→chart, code→code-block, exercise→exercise, overview→sidebar)
- `src/ppt_pro_max/enterprise/scanner.py` — Scans project dir for template.pptx, brand.json, content.json, README.md, images
- `src/ppt_pro_max/enterprise/content_parser.py` — Parses content.json AND README.md (P4: H1→pages, H2→bullets, code blocks, tables, images, goal inference with English+Chinese keywords)
- `src/ppt_pro_max/enterprise/image_matcher.py` — Image assignment: keyword-based `match_images()` + size-aware `assign_images_by_size()` (P5: >1500px→background, 800-1500→scene, <800→icon) + `auto_generate_image_prompts()` for AI image fetcher
- `src/ppt_pro_max/enterprise/proposal_generator.py` — P6: Generate 2-3 style preview PPTs (4 slides each: hook/problem/features/cta) with differentiated palettes/moods
- `src/ppt_pro_max/enterprise/slide_extractor.py` — P9: Extract content + layout from existing PPT
- `src/ppt_pro_max/enterprise/smartart_extractor.py` — P12: SmartArt XML parsing (data/layout/colors/quickStyle, no drawing needed)
- `src/ppt_pro_max/enterprise/group_extractor.py` — P12: GroupShape recursive extraction (texts/images/structure)
- `src/ppt_pro_max/enterprise/ole_extractor.py` — P12: OLE/embedded object metadata extraction
- `src/ppt_pro_max/enterprise/delivery_gate.py` — Delivery quality gate (predecessor of BuildQA)
- `src/ppt_pro_max/build_qa.py` — BuildQA: three-tier QA (fatal/warning/review) for Build-mode PPTs
- `src/ppt_pro_max/enterprise/template_analyzer.py` — Template analysis for VI Build
- `src/ppt_pro_max/enterprise/version_manager.py` — Output versioning

### Advanced Design Effects (AD-P1~P7)

- `src/ppt_pro_max/renderer/text_effects.py` — Text-level effects: gradient fill, outline, shadow, glow, 3D, alpha, vertical text, rotation, letter spacing (operates on `a:rPr`)
- `src/ppt_pro_max/renderer/blip_fill.py` — Image-in-shape via `a:blipFill`: circle/hexagon/diamond image, 22 artistic effects, duotone, grayscale, brightness/contrast, saturation
- `src/ppt_pro_max/renderer/image_processor.py` — 7 Pillow filters (grayscale, sepia, duotone, ink_wash, blur, vignette, edge_fade) with caching
- `src/ppt_pro_max/renderer/visual_effects.py` — Shape-level effects + Phase 4: `Shape3D` dataclass, `apply_3d()`, `apply_bevel()`, `PATTERN_TYPES` (31 patterns), `apply_pattern_fill()`, `apply_frosted_glass()`
- `src/ppt_pro_max/renderer/animation.py` — Phase 5: `EXIT_PRESETS` (8), `EMPHASIS_PRESETS` (8), `add_exit_animation()`, `add_emphasis_animation()`, morph transition
- `src/ppt_pro_max/renderer/decoration_library.py` — Phase 6: brush divider, seal stamp, scroll frame, neon border, grid background, glass panel, ink splash
- `src/ppt_pro_max/renderer/theme_composer.py` — Phase 7: `_MOOD_TEXT_EFFECT_MAP`, `_MOOD_IMAGE_EFFECT_MAP`; `compose()` returns `text_effect_preset` + `image_effect`

### SVG Compiler

- `src/ppt_pro_max/renderer/svg_compiler/` — SVG→native editable PPTX shapes compiler
- `src/ppt_pro_max/build_helpers.py:svg_chart()` — Tier 3 component-level brush: `svg_chart(slide, svg_text, x, y, w, h, C=C)`

#### When to use `svg_chart()` vs build_helpers atoms

**Decision tree (3 layers):**

**Layer 1 — If build_helpers has a direct atom, use it:**

| Need | Atom function | Notes |
|------|--------------|-------|
| Single rect/circle/shape | `rect()`, `oval()`, `shape()` | |
| KPI number card | `kpi_card()` | Grouped: bg + accent bar + number + label + trend |
| Horizontal bar chart | `bar_chart()` | Grouped: bg bars + value bars + labels |
| Comparison bars (left/right) | `comparison_bars()` | |
| Donut/pie chart | `donut_chart()` | Native chart when sectors>1, shape fallback for 1 sector |
| Native chart (bar/line/pie/radar/bubble/...) | `native_chart()` | 20+ chart types, editable data |
| Code block | `code_block()` | Dark bg + language badge + monospace |
| Flow chart single node | `flow_process()`, `flow_decision()`, `flow_data()`, `flow_document()` | |
| Text / multiline | `text()`, `multiline()` | CJK font injection automatic |
| Image in shape | `circle_image()`, `hex_image()`, `diamond_image()`, `star_image()`, `heart_image()` | |
| Single funnel shape | `funnel()` | MSO_SHAPE.FUNNEL only — no labels or segments |

**Layer 2 — If atoms cannot compose the target, use `svg_chart()`:**

| Need | Why svg_chart | Key SVG features used |
|------|--------------|----------------------|
| Pyramid (3-5 tiers + side labels + connectors) | Multi-trapezoid coordinate alignment | `<polygon>`, `<text>`, `<line>` |
| Pentagon/radar (5+ vertices + radial lines + labels) | Radial vertex layout | `<polygon>`, `<circle>`, `<text>` |
| Multi-segment funnel (graduated width + per-segment labels) | `funnel()` is a single shape, no labels | `<polygon>`, `<text>`, `linearGradient` |
| Venn diagram (overlapping circles + transparency) | Circle intersection + alpha blending | `<circle>`, `fill-opacity` |
| Gauge/dashboard (arc + pointer + tick marks) | Arc paths + rotation | `<path>` (A command), `transform` |
| Org chart / tree (multi-level + polyline connectors) | Rectangles + bent connectors | `<rect>`, `<polyline>`, `<text>` |
| Sankey diagram (flows + curved connections) | Bezier paths | `<path>` (C/Q commands) |

**Layer 3 — Mixed (recommended default):**
```python
slide = page_header(...)                                          # atom
       + svg_chart(slide, pyramid_svg, 2, 1.5, 6, 5.5, C=C)    # complex region
       + kpi_card(slide, 8.5, 1.5, ...)                          # atom
       + text(slide, 0.5, 7, ...)                                 # atom
```

**Quick decision signals:**

| Signal | Choose |
|--------|--------|
| ≥ 3 shapes needing **coordinate interdependence** (pyramid tiers aligned, radar vertices radiating) | `svg_chart()` |
| `defs`/`linearGradient`/`radialGradient`/`clipPath` needed | `svg_chart()` |
| Freeform curves (`path` with Q/C Bezier commands) | `svg_chart()` |
| Single responsibility (one shape, one text, one card) | atom |
| Atom function fully covers the need | atom |

#### SVG Compiler internals

- **Font scale**: SVG `font-size` is in SVG user units (viewBox pixels), not points. Conversion: `scaled_fs = max(parent_fs * scale * 72.0, 6.0)` where `scale = region_w / viewBox_w`. The `72.0` factor converts inches to points (1 inch = 72 pt). See `_text.py:324`.
- **Text overflow**: SVG `<text>` may overflow the viewBox — the compiler preserves SVG-specified font sizes and does not shrink text to fit the region. Textbox width uses natural measurement (`metrics.width_inches + 0.2`).
- **CJK measurement**: When PIL `truetype` fails and content has CJK characters, falls back to `estimate_text_size()` (character-based width table). See `_text.py:_has_cjk()`.
- **Overlap warnings**: `_detect_text_overlaps()` uses O(n²) pairwise check on text boxes — emits warnings for overlaps > 0.05". Adjacent text boxes within a card (badge + title) trigger cosmetic warnings; these are expected and do not affect rendered output.
- **Scaling mode**: Currently "contain" only (`min(w/vw, h/vh)`). No "cover" or "stretch" option yet.
- **`<use>` element**: Supported — `_render_use()` resolves `href`/`xlink:href` against `_defs`, applies x/y offset via Affine, and recursively walks the referenced element.
- **`gradientTransform`**: Supported — `apply_gradient()` transforms gradient coordinates via `tf.apply()` for both radial and linear gradients.
- **Stroke dash/cap/join**: Supported — `parse_stroke_style()` and `apply_stroke_style()` are called in `_render_shape()`.
- **Sanitizer**: Uses `_fix_self_closing_lxml()` (not regex) — preserves child elements like `<rect><title>...</title></rect>`.

#### SVG Compiler known gaps

| Gap | Location | Impact |
|-----|----------|--------|
| `_BASELINE_MAP` defined twice | `_text.py:42-51` and `_text.py:94-103` | Maintenance risk — future edits may only change one copy |
| `tspan` `dx`/`dy` offsets — partial | `_text.py:_render_tspan_text` | `dx` approximated via spacer run with `spc`; `dy` triggers new line when >0.5 units, adjusts `space_before`. Intra-line `dy` not pixel-perfect |
| `rect` `rx`/`ry` — implemented | `_compiler.py:_rounded_rect_cubics` | Rounded rectangles rendered as freeform with 4 cubic Bezier arcs per corner |
| Scaling modes — implemented | `_compiler.py:_to_inches`, `svg_chart(scaling=...)` | "contain" (default), "cover", "stretch" all supported |

### Installer & Skill Source

- `skill/` — **唯一 skill 源目录**（SKILL.md, AGENTS.md, .env.example, scripts/, data/, src/ junction）
- `installer/install.py` — 安装器主入口（从 skill/ 复制到各平台目录）
- `installer/detect.py` — 平台检测逻辑（13 个 AI 平台）
- `installer/renderer.py` — SKILL.md 模板渲染（平台差异化内容）
- `installer/platforms/*.json` — 每平台配置 JSON
- `skill.json` — 跨平台元数据（版本、平台列表、安装命令）
- `install.py`（根目录）— 向后兼容入口，委托到 `installer/install.py`

## Prerequisites

- Python 3.x
- **Pillow >= 10.0** required for cover-fit image cropping
- **设计数据库已内置** — 7 个 CSV 数据集（colors/typography/styles/products/landing/ui-reasoning/motion）打包在 `src/ppt_pro_max/data/`，无需外部安装

## Key Constraints

- **python-pptx 1.0.2**: `PP_TRANSITION_TYPE` does NOT exist, must use XML for transitions
- **Cover-fit images**: Use `_add_picture_cover()` which Pillow pre-crops — never use `add_picture` with stretch
- **Cache-first**: All image engines check cache before API call to prevent duplicate charges
- **Source of truth**: `skill/` 目录是唯一 skill 源；`src/ppt_pro_max/` 是 Python 包源码
- **Windows**: Use `python` not `python3`
- **Pipeline unified (P2)**: Rendering always goes through PrecisionRenderer.render_slide(), no dual-path
- **FreeStyle unified**: FreeStyle uses PrecisionRenderer — never falls back to old PPTRenderer
- **Content priority**: content.json > README.md > StoryPlanner (P4 integration)
- **Image assignment flow**: match_images() → assign_images_by_size() → auto_generate_image_prompts() → ImageFetcher (P5 integration)
- **Proposal flow**: `generate_ppt(proposal=True)` → 3 preview PPTs → user picks → `generate_ppt(confirmed_proposal="B")` (P6-P7)
- **SmartArt storage**: 4 XML parts only (data/layout/colors/quickStyle), no drawing (PowerPoint auto-rebuilds), colors.xml must store original (P12)
- **BuildQA**: three-tier QA (fatal/warning/review) — decorative bleeds are review, real content overflows are fatal

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
| P9 | SlideExtractor | Done | 17 in test_slide_extractor.py |
| P12 | SmartArt/GroupShape/OLE XML extractors | Done | 8 in test_p9_p14.py |
| BuildQA | Three-tier QA (fatal/warning/review) for Build mode | Done | included in test_build_qa.py |
| DQ | Design Quality Upgrades (Tier 1+2+3, 28 upgrades) | Done | 95 in test_design_quality.py + 25 in test_design_integration.py |
| AD-P1 | Text effects (gradient, outline, shadow, glow, 3D, vertical, rotation, alpha, spacing) | Done | 23+40+5 = 68 in test_text_effects/api/spc |
| AD-P2 | Image effects & fill (blipFill, circle/hex/diamond image, 22 artistic, 7 Pillow filters) | Done | 15+27+13 = 55 in test_blip_fill/image_effects/api |
| AD-P3 | Style system expansion (5 palettes, 5 fonts, 5 decorations, 4 layouts, 5 moods, 5 CJK) | Done | 108 in test_style_expansion |
| AD-P4 | 3D shapes & pattern fill (Shape3D, bevel, 31 patterns, frosted glass) | Done | 38 in test_3d_pattern_frosted |
| AD-P5 | Animation & transition expansion (morph, 8 exit presets, 8 emphasis presets) | Done | 33 in test_animation_expansion |
| AD-P6 | Decoration library (brush divider, seal stamp, scroll, neon, grid, glass, ink splash) | Done | 45 in test_decoration_library |
| AD-P7 | Mode integration (mood→text_effect_preset, mood→image_effect, compose() API) | Done | 33 in test_mode_integration |

## End-to-End Evaluation

**1,377 tests passed, 7 skipped, 0 failures. Lint clean on all modified files.**

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
