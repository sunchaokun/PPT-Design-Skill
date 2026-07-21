---
name: ppt-design-skill
version: 0.9.1
description: "AI-powered PPT generation — 40,000+ style combinations, narrative-driven, design-intelligent, AI images, fully editable .pptx. Dual-mode: FreeStyle + Enterprise. 8 goal-type layouts, 35 moods, README parsing, size-aware image assignment, proposal preview, brand compliance, component chart library. Engines: Seedream, GPT Image, DALL-E, Wanx, Kimi."
argument-hint: "[topic] [--style style-description] [--fetch-images] [--project DIR] [--proposal]"
license: MIT
metadata:
  author: sunchaokun
  category: design
  tags: [ppt, presentation, deck, pitch, slides, python-pptx, enterprise, brand, diagram, proposal]
---

# PPT Design Skill

AI-powered PPT generation — dual-mode engine, 40,000+ style combos, 10 diagram types, brand compliance, version control, fully editable .pptx.

## Execution Workflow

ALWAYS follow this 5-step workflow. Each step requires user confirmation before proceeding. Do NOT skip steps or generate final PPT directly — rework is extremely costly.

### Step 1: Requirements & Framework

- Understand: topic, audience, language, scenario
- Read any user-provided materials (README, docs, data files)
- Design the skeleton: total pages, per-page goal, core title for each page
- Determine: language (zh/en), business_mode, style direction
- Present to user as text outline, confirm before proceeding

### Step 2: Visual Proposals (3 styles)

- Generate 3 preview PPTs with different styles but same framework content
- Write a lightweight content.json (framework-level titles + short placeholder bullets) to a temp file, then call pipeline 3 times with different `--style`
- Example: `generate_ppt(query, content_file=temp_json_path, style="dark cyberpunk")` × 3 styles
- Note: current ProposalGenerator uses hardcoded 4-page preview; when `content_file` is passed, the freestyle/enterprise path is used instead
- User picks one style direction (A/B/C) or requests adjustments
- Low rework cost: only style parameters change

### Step 3: Detailed Content → content.json

- Write full content for every page: titles, subtitles, bullets, cards, code, exercise, chart data
- MUST be query-specific and domain-accurate — NEVER use generic template content
- MUST follow the Content Design Rules below to trigger design capabilities
- Save as content.json, present key content to user for review
- User confirms content accuracy before proceeding

### Step 4: Draft Generation & Revision

- Generate full PPT: `generate_ppt(query, content_file=..., style=confirmed_style, fetch_images=True, ...)`
- Verify output: check page count, file size, content rendering
- For revisions: modify content.json and regenerate, or use `--pages` / `--beautify`

### Step 5: Final Delivery

- User confirms satisfaction
- Pipeline auto-saves with version control

### Content Design Rules (CRITICAL — maximizes design quality)

When writing content.json, follow these rules to produce the best possible rendering output.

| Rule | Why | Example |
|------|-----|---------|
| features: first card is the most important, with longer body text | First card gets featured treatment (gradient bar, larger title, higher elevation) | Card 1: "智能推理引擎 — 自动选择最优框架，量化压缩70%，P99<10ms" vs Card 2: "全链路监控" (shorter) |
| data/overview pages: write 6+ bullets (short, <50 chars each) | 6+ bullets trigger two-column layout for better density | 6 concise data points instead of 3 long ones |
| tech topics: include a code page with `language` + `source` | Code pages add technical credibility | `{"code": {"language": "python", "source": "deploy(model)\nserver.start()"}}` |
| education/training: include an exercise page with `duration` + `steps` | Exercise pages add interactivity | `{"exercise": {"instructions": "...", "duration": "5 min", "steps": [...]}}` |
| topic transitions: insert `{"goal": "section", "title": "..."}` | Section dividers create visual rhythm (auto-rendered with oversized number + gradient line) | Between problem→solution or features→data |
| hook: short subtitle (<40 chars); cta: long subtitle (>60 chars) | Different subtitle lengths trigger different hero compositions (bottom-fade vs split-left) | hook: "5分钟取代5周" vs cta: "免费额度包含1000次推理/月，无需信用卡" |
| vary bullet density across pages: some 3-bullet, some 6+ | Varying density makes the deck feel natural, not templated; layout engine adapts positioning | Don't make every page the same density |
| use concrete numbers and data | Specific claims are more persuasive than vague statements | "GPU成本年增3倍" not "成本持续增长" |

## When to Activate

- User asks to create/generate/design a **PPT/presentation/deck/slide deck**
- User wants a **pitch deck, product demo, sales presentation, investor deck**
- User wants to **convert content/outline into PowerPoint**
- User wants **brand-compliant** presentations with template + version control
- User wants **page-level CRUD** on existing PPT (add/delete/swap/move pages)
- User wants **diagrams** in PPT (flowchart, funnel, timeline, SWOT, etc.)

## Dual-Mode Architecture

### FreeStyle Mode (Quick Exploration)

One command, AI generates everything — content, design, images.

```bash
python -m ppt_pro_max "AI startup investor pitch"

# Natural language style (40K+ combos)
python -m ppt_pro_max "fintech pitch" --style "warm fintech"
python -m ppt_pro_max "product launch" --style "dark cyberpunk"
python -m ppt_pro_max "brand strategy" --style "elegant luxury"

# AI images (Seedream recommended)
python -m ppt_pro_max "AI pitch" --fetch-images --llm-provider seedream

# Exact atom control
python -m ppt_pro_max "pitch" --palette wine-burgundy --fonts elegant-serif --decoration gold-trim --layout-variant centered

# Design dials
python -m ppt_pro_max "pitch" --variance 7 --motion 5 --density 6
```

### Enterprise Mode (Brand Compliance)

Project directory driven — template + brand + content + version control.

```bash
# Initialize project
mkdir my-project && cd my-project
# Place: template.pptx, brand.json, content.json, logo.png, images/

# Generate with brand compliance
python -m ppt_pro_max "路演" --project ./my-project --density 6

# Review before generating
python -m ppt_pro_max "路演" --project ./my-project --review

# Page revision (delete, swap, move, insert)
python -m ppt_pro_max "" --project ./my-project --pages "-3,2<>5,10>3,+6"

# Version history
python -m ppt_pro_max "" --project ./my-project --history

# Business mode
python -m ppt_pro_max "培训" --project ./my-project --business-mode education
```

## Python API

```python
from ppt_pro_max import generate_ppt, fetch_image

# FreeStyle
result = generate_ppt("AI startup pitch", style="dark cyberpunk", fetch_images=True)

# Enterprise
result = generate_ppt("路演", project="./my-project", density=6)

# Page revision
result = generate_ppt("", project=".", pages="-3,2<>5")

# Version history
result = generate_ppt("", project=".", history=True)

print(f"Generated: {result['output_path']}, {result['page_count']} pages")

# Standalone image generation (no PPT needed)
img = fetch_image("futuristic AI city", mode="generate", llm_provider="seedream", llm_api_key="...")
print(img["path"])  # Local file path

# Search stock photos
img = fetch_image("team meeting", mode="search", unsplash_access_key="...")

# Auto: try AI generation, fall back to search
img = fetch_image("product launch", mode="auto", llm_provider="seedream", llm_api_key="...")
```

## 4-Phase Pipeline

1. **Story Planning** → strategy + page structure + emotion arc
2. **Design Decisions** → per-page layout/color/typography from 40K+ combos
3. **Content Generation** → copy formulas (PAS/FAB/AIDA) + image keywords
4. **PPT Rendering** → python-pptx direct, 12 master layouts, QA gates

## Design Atoms (40,000+ Combos)

| Atom | Count | Examples | Status |
|------|-------|----------|--------|
| Color Palettes | 25 | ocean-blue, cyber-neon, golden-luxury, wine-burgundy, midnight-navy, monochrome-dark... | All working |
| Font Pairs | 20 | modern-sans, elegant-serif, tech-mono, contrast-mix, sharp-modern... | All working |
| Decorations | 10 | accent-bar, neon-lines, gold-trim, diamond-bullets, gradient-bar, sidebar-nav, minimal-dots, circle-accent, no-decoration, full-bleed-overlay | All rendered via DecorationRenderer |
| Layout Variants | 8 | standard, centered, sidebar-left, sidebar-right, grid-2x2, asymmetric... | Consumed by PrecisionRenderer |

Natural language: `--style "warm fintech"` auto-selects matching atoms. Decoration and layout-variant atoms are now fully consumed by PrecisionRenderer — they control title decoration, margin positioning, and card style.

## 10 Diagram Types

| Type | Description | Data Format |
|------|-------------|-------------|
| Flowchart | Process flow, auto horizontal/vertical | nodes + connectors |
| Funnel | Decreasing width stages | stages (items) |
| Timeline | Alternating top/bottom labels | events (items) |
| SWOT | 4-quadrant analysis | strengths/weaknesses/opportunities/threats |
| Matrix | Comparison grid | rows + columns |
| Cycle | Circular arrangement | stages (items) |
| Table | Alternating row colors | headers + rows |
| Hierarchy | Parent-child tree | nodes with parent |
| Pyramid | Stacked levels | levels (items) |
| Venn | 2-3 set intersection | sets with labels |

## Image Engines

| Engine | Provider | Env Key | Default Model |
|--------|----------|---------|---------------|
| Seedream | Volcengine | ARK_API_KEY | doubao-seedream-5-0-260128 |
| GPT Image | OpenAI | OPENAI_API_KEY | gpt-image-1 |
| DALL-E 3 | OpenAI | OPENAI_API_KEY | dall-e-3 |
| Wanx | Alibaba | DASHSCOPE_API_KEY | wanx-v1 |
| Kimi K2.6 | Moonshot | MOONSHOT_API_KEY | kimi-k2-0711-preview |

Image modes: `placeholder` (default), `search` (Unsplash/Pexels), `generate` (AI), `enhance` (Kimi keyword optimization + search).

All engines use **cache-first** — same image never generated twice.

### Standalone Image Generation (No PPT Required)

The image module can be used independently — no need to generate a full PPT.

**CLI subcommand:**
```bash
# AI image generation
python -m ppt_pro_max image "futuristic AI city" --llm-provider seedream --llm-api-key $ARK_API_KEY

# Search Unsplash/Pexels
python -m ppt_pro_max image "startup team meeting" --image-mode search --unsplash-key $KEY

# Kimi enhance + search
python -m ppt_pro_max image "product demo" --image-mode enhance --llm-provider kimi

# Custom size + verbose info
python -m ppt_pro_max image "AI robot" --llm-provider seedream --width 2560 --height 1440 -v
```

**Python API:**
```python
from ppt_pro_max import fetch_image

# Generate AI image
result = fetch_image(
    "futuristic AI city",
    mode="generate",
    llm_provider="seedream",
    llm_api_key="...",
    width=1920,
    height=1080,
)
print(result["path"])  # Local file path, or None if failed
print(result["provider"])  # e.g. "seedream"

# Search stock photos
result = fetch_image("team meeting", mode="search", unsplash_access_key="...")

# Auto mode: try AI generation, fall back to search
result = fetch_image("product launch", mode="auto", llm_provider="seedream", llm_api_key="...")
```

**Parameters:**

| Parameter | Default | Description |
|-----------|---------|-------------|
| `keywords` | (required) | Image search/generation keywords |
| `mode` | `"auto"` | `placeholder` / `search` / `generate` / `enhance` / `auto` |
| `emotion` | `""` | Emotion hint: curiosity, hope, confidence, warmth, urgency, fear |
| `goal` | `""` | Slide goal: hook, problem, solution, features, cta, product |
| `width` | `1920` | Image width in pixels |
| `height` | `1080` | Image height in pixels |
| `llm_provider` | auto-detect | seedream / gpt-image / dalle / gemini / wanx / kimi |
| `llm_api_key` | env var | API key (or set env: ARK_API_KEY, OPENAI_API_KEY, etc.) |
| `auto_detect` | `True` | Auto-detect LLM config from host tools |

**Returns:** `{"path": str|None, "mode": str, "provider": str, "keywords": str, "width": int, "height": int}`

## Animation System

- **12 Transitions**: fade, push, wipe, split, cover, dissolve, wheel, wedge, blinds, checker, comb, random
- **11 Entrances**: appear, fly_in, fade_in, zoom_in, float_up, bounce, etc.
- **Motion mapping**: 1-2 = transitions only, 3-5 = fade_in entrance, 6-10 = fly_in entrance
- Applied via XML injection (python-pptx 1.0.2 has no native transition API)

## Enterprise Project Structure

```
my-project/
├── template.pptx    # Brand template (optional)
├── brand.json       # Brand spec (colors, fonts, logo, footer, watermark)
├── content.json     # Page content (title, bullets, image, diagram, code, exercise)
├── logo.png         # Company logo
├── images/          # Local images
└── output/
    ├── v1/          # Version 1
    │   ├── presentation.pptx
    │   └── meta.json
    └── v2/          # Version 2 (auto-incremented)
```

## content.json Format

```json
{
  "meta": {"title": "...", "author": "..."},
  "slides": [
    {
      "goal": "hook|problem|solution|features|cta|content|data|code|exercise|section|overview",
      "title": "Page Title",
      "subtitle": "Optional subtitle",
      "bullets": ["Point 1", "Point 2"],
      "image": "images/photo.png",
      "cards": [{"title": "...", "text": "..."}],
      "diagram": {"type": "flowchart", "data": {...}},
      "code": {"language": "python", "source": "..."},
      "exercise": {"instructions": "...", "duration": "5 min", "steps": [...]},
      "component_type": "group",
      "component_category": "process"
    }
  ]
}
```

**Goal types and rendering behavior:**

| Goal | Rendering | Notes |
|------|-----------|-------|
| `hook` | Hero: full-bleed image + gradient overlay + title | First slide; short subtitle recommended |
| `cta` | Hero: full-bleed image + gradient overlay + title | Last slide; long subtitle recommended |
| `section` | Section divider: oversized number + title + gradient line | Auto-handled by render_slide() when goal="section"; Pipeline can auto-insert on topic shifts |
| `problem`/`solution`/`content` | Title + gradient line accent + bullets + optional image | Standard content slide; 6+ bullets → two-column |
| `features` | Title + cards row | First card gets featured treatment (gradient bar, 22pt title, higher elevation) |
| `data`/`overview` | Title + bullets or diagram | 6+ bullets trigger two-column layout |
| `code` | Title + code block (always dark bg #1E293B) + language badge | `language` + `source` required |
| `exercise` | Title + badge (ALL CAPS, solid variant) + instructions + numbered steps | `duration` + `steps` recommended |

## brand.json Format

```json
{
  "colors": {
    "primary": "#1E3A5F", "on-primary": "#FFFFFF",
    "accent": "#E8A838", "background": "#0A1E3D",
    "foreground": "#F0F4F8", "muted": "#1A2E4A",
    "muted-foreground": "#8A9BB5"
  },
  "logo": {"position": "top_right", "width_inches": 1.0, "skip_cover": true},
  "footer": {"text": "Company Name", "show_page_number": true},
  "watermark": {"text": "CONFIDENTIAL", "opacity": 0.1},
  "spacing": {
    "strip_style": "auto",
    "section_dividers": true
  }
}
```

**spacing options:**
- `strip_style`: `"auto"` (varied: left bar / bottom line / none, cycling by page), `"left"` (always left bar), `"none"` (no strip). Default: `"auto"`
- `section_dividers`: `true` (auto-insert section dividers on topic shifts), `false` (no dividers). Default: `true`

## Page Revision Syntax

```
--pages "3,5 +6 -8 3>5 3<>7"
```

| Syntax | Action | Example |
|--------|--------|---------|
| `N` | Keep page N | `3,5` keep pages 3 and 5 |
| `+N` | Insert new page at N | `+6` insert at position 6 |
| `-N` | Delete page N | `-3` delete page 3 |
| `N>M` | Move page N to position M | `10>3` move 10 to 3 |
| `N<>M` | Swap pages N and M | `2<>5` swap 2 and 5 |

All page numbers are 1-based, refer to ORIGINAL document.

## .env Configuration

Copy `.env.example` to `.env` and fill in API keys:

```env
ARK_API_KEY=your-key-here
OPENAI_API_KEY=sk-...
DASHSCOPE_API_KEY=...
MOONSHOT_API_KEY=...
UNSPLASH_ACCESS_KEY=...
PEXELS_API_KEY=...
```

Search order: CWD/.env → package root/.env → ~/.ppt-pro-max/.env

## CLI Reference

```
ppt-design "query" [options]
ppt-design image "keywords" [image-options]

# Style
  --style TEXT          Natural language style description
  --palette NAME        Color palette (25 options)
  --fonts NAME          Font pair (20 options)
  --decoration NAME     Decoration style (10 options)
  --layout-variant NAME Layout variant (8 options)
  --mood TEXT           Mood hint
  --style-seed INT      Reproducible style seed

# Content
  --slides N            Override slide count
  --content FILE        JSON content file
  --variance 1-10       Design variance
  --motion 1-10         Animation intensity
  --density 1-10        Content density

# Images (PPT mode)
  --image-mode MODE     placeholder|search|generate|enhance
  --fetch-images        Shortcut for --image-mode search
  --llm-provider PROV   seedream|gpt-image|dalle|wanx|kimi
  --llm-api-key KEY     API key (or .env)
  --llm-base-url URL    API base URL override
  --llm-model MODEL     Model name override

# Enterprise
  --project DIR         Project directory (Enterprise mode)
  --business-mode MODE  pitch|education|training|report
  --review              Preview plan before generating
  --review-file FILE    Save review plan to JSON
  --output-version N    Specify version number
  --from-version N      Base on specified version
  --pages OPS           Page operations (-3,2<>5,10>3,+6)
  --history             Show version history
  --component-library PATH  Component library DB path (auto-detected if omitted)

# Beautify
  --beautify PPTX       Beautify existing PPT file
  --beautify-mode MODE  light (color/font swap) or full (rebuild + component injection)

# Output
  --persist             Save design system as MASTER.md
  --dry-run             Output design decisions only
  -o, --output PATH     Output .pptx path

# Standalone Image (ppt-design image)
  keywords              Image keywords (required)
  --image-mode MODE     placeholder|search|generate|enhance|auto
  --fetch-images        Shortcut for --image-mode search
  --emotion EMOTION     Emotion hint (curiosity, hope, confidence, warmth, urgency)
  --goal GOAL           Slide goal (hook, problem, solution, features, cta)
  --width N             Image width (default: 1920)
  --height N            Image height (default: 1080)
  -v, --verbose         Print extra info (mode, provider, etc.)
  --llm-provider PROV   seedream|gpt-image|dalle|gemini|wanx|kimi
  --llm-api-key KEY     API key (or .env)
  --llm-base-url URL    API base URL override
  --llm-model MODEL     Model name override
  --unsplash-key KEY    Unsplash API key
  --pexels-key KEY      Pexels API key
  --no-auto-detect      Disable auto-detection of LLM config
```

## Component Library (Professional Chart Templates)

A SQLite-indexed library of GroupShape/SmartArt templates extracted from real PPT files, with coordinate normalization for universal scaling. Count varies by build — run `query_component_library()` for current catalog.

### Library Overview

| Category | Count | Use When |
|----------|-------|----------|
| infographic | 4,101 | Data visualization, statistics, KPI dashboards |
| process | 672 | Workflows, step-by-step, pipelines, procedures |
| hierarchy | 548 | Org charts, reporting structures, tree diagrams |
| chart | 132 | Bar/pie/line chart layouts, data comparison |
| timeline | 42 | Milestones, roadmaps, chronological events |
| swot | 39 | Strategic analysis, 4-quadrant frameworks |
| features | 2 | Feature showcases |
| comparison | 1 | Side-by-side comparisons |

### How to Query the Library

```python
from ppt_pro_max import query_component_library

# Browse full catalog
catalog = query_component_library()
# Returns: {"group": {"infographic": {"count": 4101, "min_nodes": 3, "max_nodes": 12}, ...}, ...}

# Search by type + category
results = query_component_library(type="group", category="process")
# Returns list of matching components sorted by node_count DESC

# Search with node_count for precise matching
results = query_component_library(type="group", category="hierarchy", node_count=5)
```

### How to Use Components in content.json

Add `component_type` and `component_category` to any slide. The Pipeline auto-matches from the library and injects a professional chart template scaled to fit:

```json
{
  "goal": "content",
  "title": "项目流程",
  "bullets": ["需求分析", "方案设计", "开发实现", "测试上线"],
  "component_type": "group",
  "component_category": "process"
}
```

Matching logic:
1. Exact match: `type + category + node_count` (node_count = len(bullets))
2. Fuzzy match: same type+category, closest node_count
3. Fallback: DiagramEngine or bullet list rendering

### Component Selection Strategy (Decision Guide)

**When to use components vs built-in layouts:**

| Scenario | Use Component Library | Use Built-in Layout |
|----------|----------------------|---------------------|
| Process/flow with 3-8 steps | `process` component | Only if no library match |
| Org chart / reporting structure | `hierarchy` component | Never use built-in for this |
| Data dashboard / KPI grid | `infographic` component | Only for simple 2-3 metrics |
| Timeline / milestones | `timeline` component | Only if no library match |
| SWOT analysis | `swot` component | Only if no library match |
| Simple 3-card features | Built-in `features` cards | Better brand consistency |
| Code block | Built-in code renderer | Components don't help here |
| Bullet list (2-5 items) | Built-in bullet renderer | Simpler = better |

**Selection priority:**
1. **Complex diagrams** (hierarchy, process, timeline, swot) → ALWAYS try component library first
2. **Data-heavy pages** (infographic, chart) → Component library for visual richness
3. **Simple content** (bullets, cards, code) → Built-in PrecisionRenderer layouts

**Node count matching tips:**
- `node_count` = number of text items in the chart (e.g., 4 steps → node_count=4)
- Prefer exact node_count match for best layout fit
- If between two node counts, choose the larger one (extra nodes can be hidden/empty)
- For `hierarchy`: node_count = total positions (CEO + 3 VPs = 4)
- For `process`: node_count = number of steps
- For `swot`: node_count is typically 4 (one per quadrant)

### Design Quality Optimization

**1. Component + brand color synergy:**
Components auto-apply brand colors via `schemeClr → srgbClr` replacement. For best results:
- Ensure `brand.json` has complete color roles: `primary`, `accent`, `foreground`, `muted`, `background`
- Dark themes work best with `infographic` and `chart` components (high contrast)
- Light themes work best with `process` and `hierarchy` components (clean lines)

**2. Component density vs page density:**
- Component pages should have `--density` 5-7 (not too dense, chart needs breathing room)
- Title + component chart is the ideal page composition — avoid adding extra bullets alongside a component
- If both `bullets` and `component_type` are set, bullets become the data source for the chart

**3. Mixing components and built-in elements:**
- Hook/CTA pages: always use built-in hero layout (full-bleed image + overlay)
- Content pages with diagrams: use component library
- Content pages with simple lists: use built-in bullet renderer
- Never put a component on a hook/CTA page

**4. Beautify mode optimization:**
- `--beautify-mode light`: color/font replacement only, preserves original layout (fast, safe)
- `--beautify-mode full`: extracts content → matches components → rebuilds with PrecisionRenderer (higher quality, but restructures slides)
- For existing professional PPTs: use `light` mode
- For PPTs needing structural improvement: use `full` mode

### Database Management

```bash
# Build library from PPT素材
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材目录" --min-node-count 3

# Rebuild from scratch (required after coordinate normalization changes)
rmdir /s /q component_library
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材目录" --min-node-count 3

# Multiple directories (dedup is automatic)
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材A"
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材B"
```

**min_node_count=3 is the recommended threshold** — filters out decorative fragments (single labels, icon+text pairs) that PrecisionRenderer can generate better from scratch. The library's value is in complex chart templates (3+ nodes) that are expensive to build manually.

### Coordinate Normalization

All component coordinates are normalized to 0-1 range at extraction time. At injection, they are denormalized to the target bounding box. This means:
- Components render correctly at any size (small card, half-page, full-page)
- Font sizes are clamped to 8pt-72pt to handle extreme aspect ratios
- Nested GroupShapes are recursively normalized using each group's own chOff/chExt

## Design Quality Standards

All 28 upgrades from `docs/design-quality-upgrade-plan.md` are now implemented (v0.7.0). 824 tests pass.

**Implemented rendering capabilities:**

| Feature | Module | Effect |
|---------|--------|--------|
| Layout Engine | `layout_engine.py` | Content-adaptive positioning, margin computation |
| Typography Scale | `typography.py` | 5-level hierarchy (display→caption), letter-spacing |
| Color Depth | `color_system.py` | OKLCH-based 11 tint/shade levels, 5 alpha variants |
| Gradient Overlay | `precision_renderer.py` | Hero slides: transparent top → dark bottom (no flat black) |
| Elevation System | `elevation.py` | 5 shadow levels + dark-mode glow |
| Conditional Brand Strip | `precision_renderer.py` | Strip on ~1/3 of slides, varied patterns |
| Image Color-Grading | `image_processor.py` | 8-12% primary overlay, JPEG-preserving cache |
| Featured Card | `precision_renderer.py` | First card: gradient bar, 22pt title, higher elevation |
| Luminance Detection | `precision_renderer.py` | `_is_dark()` via luminance, not hardcoded hex list |
| Code Block Redesign | `precision_renderer.py` | Always-dark bg, styled language badge, muted-foreground text |
| CJK Font Pairing | `theme_mapper.py` | 12 Latin↔CJK pairings, `a:ea`/`a:cs` XML set on runs |
| Adaptive Margins | `layout_engine.py` | Density-aware whitespace (presenting/reading/balanced) |
| Badge System | `precision_renderer.py` | ALL CAPS, tracked, tinted/solid/outline variants |
| Section Dividers | `precision_renderer.py` | Oversized number + title + gradient line |
| Decoration Renderer | `decoration_renderer.py` | All 10 styles rendered (accent-bar, neon-lines, gold-trim, etc.) |
| Layout Variant Consumption | `precision_renderer.py` | `layout_variant` dict controls margins, title alignment |
| Noise Texture | `image_processor.py` | Per-deck seeded, 2% opacity, Box-Muller Gaussian |
| Progress Bar | `precision_renderer.py` | Thin bar at bottom, replaces muted bar when present |
| Corner Radius Scale | `precision_renderer.py` | CORNER_RADIUS_SCALE: sm=4, md=8, lg=12, pill=50 |
| Gradient Lines | `precision_renderer.py` | `add_gradient_line()` with alpha fade |
| Image Masking | `precision_renderer.py` | Rounded-rect frame with padding |
| Two-Column Bullets | `precision_renderer.py` | 6+ bullets → 2 columns with vertical separator |
| Hero Pattern Selection | `precision_renderer.py` | 4 patterns: gradient, split-left, bottom-fade, asymmetric |

## Key Constraints

- **python-pptx 1.0.2**: No `PP_TRANSITION_TYPE`, must use XML for transitions/animations
- **Cover-fit images**: Use `_add_picture_cover()` with Pillow pre-crop — never stretch
- **Cache-first**: All image engines check cache before API call
- **Two-pass rebuild**: Page revision uses rebuild (not in-place) to avoid ZIP corruption
- **1-based pages**: All `--pages` numbers refer to original document
- **Windows**: Use `python` not `python3`
- **Component library**: min_node_count=3 for quality; rebuild DB after normalization logic changes
- **Component priority**: complex diagrams (hierarchy/process/timeline/swot) → always try library first; simple content (bullets/cards/code) → built-in renderer
- **OOXML alpha**: `a:alpha val` = percentage × 1000 (e.g., 80% = 80000, NOT 0.8)
- **OOXML letter-spacing**: `a:spc val` = tracking_em × font_size_pt × 100 (font-size-dependent, NOT percentage)
- **apply_shadow() signature**: `apply_shadow(shape, blur_pt, distance_pt, direction_deg=90, color="#000000", alpha_pct=25)` — note `direction_deg` comes before `color`
- **add_text() signature**: `add_text(slide, text, x, y, w, h, font=None, size=20, color_role="foreground", bold=False, align="left")` — CJK companion font auto-set via `a:ea`/`a:cs`
- **add_rounded_rect() signature**: `add_rounded_rect(slide, x, y, w, h, fill_role=None, fill_hex=None, border_role=None, border_hex=None, gradient=False, shadow=False, corner_radius="md")` — `corner_radius` accepts `"sm"|"md"|"lg"|"pill"` or int pt value
- **BrandSpec**: has `spacing` dict (for `strip_style`, `section_dividers`), no `extra` field
- **GradientFill**: use `GradientFill` + `GradientStop` for alpha gradients; `apply_gradient(shape, color1, color2)` does NOT support alpha

## Dependencies

- python-pptx >= 1.0.2 (required)
- Pillow >= 10.0 (required)
- python-dotenv >= 1.0 (optional, for .env support)
- ui-ux-pro-max >= 1.0.0 (optional)
