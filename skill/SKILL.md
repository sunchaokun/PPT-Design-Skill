---
name: ppt-design-skill
version: 0.9.2
description: "AI-powered PPT generation — 40,000+ style combinations, narrative-driven, design-intelligent, AI images, fully editable .pptx. Three modes: FreeStyle + Build + VI Build. 8 goal-type layouts, 35 moods, README parsing, size-aware image assignment, proposal preview, brand compliance, component chart library. Engines: Seedream, GPT Image, DALL-E, Wanx, Kimi."
argument-hint: "[topic] [--style style-description] [--fetch-images] [--proposal]"
license: MIT
metadata:
  author: sunchaokun
  category: design
  tags: [ppt, presentation, deck, pitch, slides, python-pptx, brand, diagram, proposal, build, vi-build]
---

# PPT Design Skill

**⚠️ READ BEFORE coding: [`src/ppt_pro_max/docs/python-pptx-reference.md`](src/ppt_pro_max/docs/python-pptx-reference.md)** — 170+ shape types, 73 chart types, tables, connectors, freeform, hyperlinks, media, effects, 3D, OOXML. python-pptx has far more capabilities than rect/oval/textbox.

AI-powered PPT generation — three-mode engine, 40,000+ style combos, 10 diagram types, brand compliance, version control, fully editable .pptx.

## ⚠️ Non-Negotiable Sections (DO NOT compress or remove)

These sections are the LLM's only reference for writing correct output:
1. **content.json Format** — LLM must know the exact schema to write valid content
2. **brand.json Format** — LLM must know brand spec structure for VI Build mode
3. **Build Helpers API** — LLM must know function signatures to write build.py
4. **Content Design Rules** — LLM must know which content patterns trigger which rendering
5. **Key Constraints** — LLM must know API signatures to write correct python-pptx code
6. **generate_ppt() signature** — LLM must know valid parameters to call the pipeline

## Execution Workflow

ALWAYS follow this 5-step workflow. Each step requires user confirmation before proceeding. Do NOT skip steps or generate final PPT directly — rework is extremely costly.

### Step 1: Requirements & Framework

- Understand: topic, audience, language, scenario
- Read any user-provided materials (README, docs, data files)
- Design the skeleton: total pages, per-page goal, core title for each page
- Determine: language (zh/en), business_mode, style direction
- **Design Read**: declare VARIANCE (1-10), MOTION (1-10), DENSITY (1-10) based on audience and scenario
- Present to user as text outline, confirm before proceeding

**Dial → Action Map (V/M/D → LLM decisions using EXISTING content.json fields and generate_ppt() params):**

| VARIANCE | Action |
|----------|--------|
| 1-3 | Use `goal:"content"` with centered layouts; equal-width cards; `--layout-variant centered` or `standard` in generate_ppt() call |
| 4-7 | Mix `goal:"content"` with `goal:"features"`; feature first card; `--layout-variant sidebar-left` or `asymmetric` in generate_ppt() call; vary which pages have images |
| 8-10 | Use diverse goal types per page; avoid any repeated layout family; `--layout-variant asymmetric`; insert section dividers between every topic shift |

| MOTION | Action |
|--------|--------|
| 1-3 | No special action — default transitions only |
| 4-7 | Ensure cover slide has `goal:"hook"` (gets fade-in); section dividers get entrance animation automatically |
| 8-10 | Same as 4-7 plus: request `--motion 8` in generate_ppt() call; more section dividers for animation variety |

| DENSITY | Action (bullet count per page in content.json) |
|---------|------------------------------------------------|
| 1-3 | 2-3 bullets per content page; insert breathing pages (goal:"section" or goal:"content" with ≤2 bullets) after every 2 content pages |
| 4-7 | 3-5 bullets per content page; mix: some pages 3 bullets, some 6+ (triggers two-column) |
| 8-10 | 6+ bullets on data/overview pages; use `component_type:"group"` + `component_category:"infographic"` for dense pages; no breathing pages |

### Step 2: Visual Proposals (3 styles)

- Generate 3 preview PPTs with different styles but same framework content
- Write a lightweight content.json (framework-level titles + short placeholder bullets) to a temp file, then call pipeline 3 times with different `--style`
- Example: `generate_ppt(query, content_file=temp_json_path, style="dark cyberpunk")` × 3 styles
- Note: current ProposalGenerator uses hardcoded 4-page preview; when `content_file` is passed, the freestyle path is used instead
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
| features: first card featured with longer body | First card gets gradient bar + 22pt title + higher elevation | Card 1: "智能推理引擎 — 自动选择最优框架" vs Card 2: "全链路监控" |
| 6+ bullets → two-column layout | Better density; layout engine auto-splits | 6 concise data points instead of 3 long ones |
| tech topics: include code page | Code pages add technical credibility | `{"code": {"language": "python", "source": "..."}}` |
| education/training: include exercise page | Exercise pages add interactivity | `{"exercise": {"duration": "5 min", "steps": [...]}}` |
| topic transitions: insert section divider | Visual rhythm (oversized number + gradient line) | Between problem→solution |
| hook: short subtitle (<40 chars); cta: long (>60) | Different hero compositions | hook: "5分钟取代5周" vs cta: "免费额度包含1000次推理/月" |
| vary bullet density (some 3-bullet, some 6+) | Varying density feels natural; 10+ items → cards/grid/table, never list | Don't make every page the same density |
| use concrete real data; no fake precision | "GPU成本年增3倍" not "成本持续增长"; no fabricated 92%/4.1× | Real data only; mark as "example" if hypothetical |
| ≤5 bullets: single column | 6+: two-column; 10+: use cards/grid/infographic component, never list | 3 bullets → single col; 7 bullets → two-col |
| no filler verbs (赋能/领先/一站式/生态/革新/引领) | AI-generated buzzwords destroy credibility | Use plain functional language |
| quotes ≤3 lines, attribution = name+title | PPT quotes are fragments, not full reviews | "Name, CTO, Company" — never name alone |
| theme lock: one theme per deck, no mid-deck switch | Dark stays dark, light stays light; micro-variation OK | #0A1E3D → #0F2847 OK; #0A1E3D → #FFF8F0 NOT OK |

## When to Activate

- User asks to create/generate/design a **PPT/presentation/deck/slide deck**
- User wants a **pitch deck, product demo, sales presentation, investor deck**
- User wants to **convert content/outline into PowerPoint**
- User wants **brand-compliant** presentations with template + version control
- User wants **page-level CRUD** on existing PPT (add/delete/swap/move pages)
- User wants **diagrams** in PPT (flowchart, funnel, timeline, SWOT, etc.)

## Three-Mode Architecture

| | FreeStyle | Build Script | **VI Build** |
|---|---|---|---|
| **Use case** | Quick exploration, prototyping | Delivery-grade, no template | **Enterprise VI compliance** |
| **Trigger** | Default | Hand-write `build.py` | `analyze_template.py` + `build_helpers` |
| **Content source** | AI auto-generates | Hardcoded per page in build.py | LLM reads template analysis, generates build.py |
| **Brand compliance** | Style atom combos | Design Token dict `C` | **Extracted VI Token from template** |
| **Layout control** | Auto-match goal type | **Per-element x/y/w/h** | **Preserve framework pages + build_helpers for new** |
| **Font control** | Theme-level | **Run-level per character** | **Run-level + template font inheritance** |
| **Template reuse** | None | None | **Framework pages preserved + decorations/LOGO copied** |
| **Quality ceiling** | ★★★ | ★★★★★ | ★★★★★ |

> **Recommended workflow**: FreeStyle prototype → VI Build (with enterprise template) or Build Script (no template) for precision delivery.

### FreeStyle Mode (Quick Exploration)

One command, AI generates everything — content, design, images.

```bash
python -m ppt_pro_max "AI startup investor pitch"

# Natural language style (40K+ combos)
python -m ppt_pro_max "fintech pitch" --style "warm fintech"
python -m ppt_pro_max "product launch" --style "dark cyberpunk"

# AI images (Seedream recommended)
python -m ppt_pro_max "AI pitch" --fetch-images --llm-provider seedream

# Exact atom control
python -m ppt_pro_max "pitch" --palette wine-burgundy --fonts elegant-serif --layout-variant centered

# Design dials
python -m ppt_pro_max "pitch" --variance 7 --motion 5 --density 6
```

### Build Mode (Pixel-Perfect Delivery)

LLM writes `build.py` scripts from blank canvas, using build_helpers for maximum per-element control.

```bash
# LLM generates build.py, then:
python build.py
```

See **Build Helpers API** section below for function reference.

### VI Build Mode (Enterprise Template Compliance)

LLM reads template analysis, generates build.py that preserves framework pages (cover/TOC/back cover) and uses `build_helpers` for new content pages.

```bash
# Step 1: Analyze template
python -m ppt_pro_max.analyze_template template.pptx > analysis.txt

# Step 2: Give analysis.txt to LLM, which generates build.py

# Step 3: Run build.py
python build.py
```

**VI Build workflow in build.py:**

```python
from ppt_pro_max.build_helpers import *

# VI Token extracted from template analysis
C = {
    'primary': '#2E6504', 'accent': '#7DA92F', 'muted': '#81C784',
    'light': '#C8E6C9', 'white': '#FFFFFF', 'background': '#FFFFFF',
    'card_bg': '#F9F9F9', 'text_dark': '#1A1A1A', 'text_body': '#333333',
    'text_muted': '#666666', 'divider': '#CCCCCC',
    'font_heading': '微软雅黑', 'font_body': '微软雅黑',
}

# Load template (NOT Presentation() from scratch)
prs = Presentation('template.pptx')
template_slide = prs.slides[0]  # Reference for copying decorations/LOGO

# Framework pages (cover, TOC, back cover) are preserved — do NOT delete them
# Add new content pages:
s = add_slide(prs)
copy_decorations(s, template_slide)  # Copy visual elements from template
copy_logo(s, template_slide, color_hints=['#2E6504'])  # Copy company LOGO
page_header(s, 'Revenue Overview', 'FY2025 Performance', C)
kpi_card(s, 0.65, 1.8, 3.8, 1.35, '12.8亿', '年度产值', '+8.3%', C=C)

prs.save('output.pptx')
```

**Key differences from Build Script:**
- Start with `Presentation('template.pptx')` NOT `Presentation()`
- Framework pages (cover/TOC/back cover) are preserved untouched
- Use `copy_decorations()` / `copy_logo()` to maintain VI consistency
- VI Token (`C` dict) extracted from `analyze_template.py` output, not hand-written

## Design Constraints

### Typography
- Cover title: 44-52pt, ≤2 lines | Inner title: 32-36pt | Body: 14-16pt | Bullets: 13-14pt | Caption: 11-12pt
- Min 4 font-size levels per deck — 2 levels (title+body only) is forbidden
- NO Calibri/Arial as default font (PPT's AI default, same as Inter in web)
- Serif: only for editorial / luxury / heritage scenes — NOT for tech/startup/data
- Max 2 font families per deck (heading + body, +1 monospace for code)
- Emphasize with bold or color shift — NO mixing serif+sans for "contrast emphasis"
- Italic title line-height ≥1.1× (descender clearance)
- Each page: short title (≤8 words) + short subtitle (≤25 chars) + 1 visual OR 1 CTA

### Color
- Max 1 accent color per deck, used on EVERY page (consistency lock)
- NO default-blue gradient cover when style is unspecified
- NO default gold+navy for "premium" scenes (#1A1A2E / #C9A96E family)
- Warm/cool gray: pick one, use throughout — no mixing
- Chart colors derived from main palette — no rainbow
- Dark theme: text ≥60% luminance above background (projection-grade contrast)
- Light theme: no light-gray text on white (invisible on projector)

### Layout
- VARIANCE > 4: avoid all-centered; use left-aligned / sidebar / asymmetric
- Same layout family max 1 occurrence per deck ("Our Products" ≠ "Core Advantages" visually)
- 8+ slides: ≥4 distinct visual layouts
- Left-right alternation ≤2 times; 3rd = break pattern
- Eyebrow count ≤ceil(page_count / 3); NO section-number eyebrows (00/INDEX, 001·核心能力)
- No split-header as default (title left + small text right)
- Bento grid: ≥2-3 cells with visual variation (not all-white text cards)
- Spec sheets: NO 10-line bullet lists — use card grid / highlight+fold / grouped sections
- NO 20+ row data tables in PPT — PPT is presentation, not document

### Page Roles
- MUST plan 5 roles: breathing (low-density rest) / section-divider / data-impact (big number) / visual-anchor (full-bleed image) / cta
- No 3 consecutive "title + bullets" pages
- 6+ slides: ≥1 section divider
- High-density page → must be followed by low-density page

### Visual Assets
- Cover MUST have real visual (not text + gradient block)
- Even minimalist style: ≥2-3 pages with images
- NO fake product screenshots (text-box dashboards/terminals/task-lists — #1 AI Tell in PPT)
- Logo wall: use real logo images, not text spans

### Consistency Locks
- Corner radius: ONE system per deck — sharp (0pt) / soft (8-12pt) / pill — no mixing
- Font pair: consistent throughout (heading + body, all slides)
- Accent color + warm/cool gray + theme lock: enforced in Color rules and Content Design Rules

### AI Tells Blacklist (HARD BAN unless user explicitly requests)

- No cover version labels (V0.6/BETA/内测版)
- No "Brand · No.01" style sub-labels
- No section-number eyebrows (00/INDEX, 001·核心能力) — use natural language
- No card/image numbering labels (01/04, 1/3)
- Max 1 middle-dot (·) per metadata line — no "foo · bar · baz · qux"
- No decorative status dots
- No em-dash (—) or Chinese em-dash (——) — use comma/hyphen/semicolon
- No linebreak+italic "design trick"
- No vertical rotated text
- No crosshair/fine-grid decoration (only for organizing real content)
- No fake product UI (text-box dashboard/terminal/task-list)
- No fake version footers (v0.6.2-rc.1, "last sync 4s ago")
- No "silently used by" / "默默服务" social-proof headlines — use natural language or skip
- No "来自一线" / "实战笔记" artisan labels — use plain functional labels
- No city/time/weather bars (99% of scenarios)
- No eyebrow micro-metadata sentences
- No generic step labels ("Phase 1/2/3", "步骤 01/02/03") — use verb+noun
- No overlaid labels on images ("Brand · 02")
- No decorative photo credits (场景 III · 35mm) — skip or use one-line caption
- No version footers on marketing slides (v1.4.2, Build 0048)
- No inventory counters as decoration ("已预约 412/800")
- No bottom-of-cover decoration strips (品牌. 创新. 技术.)
- No floating explanation text top-right of section titles
- No divider lines on every row of long lists
- No progress bars with filled background tracks for comparison
- No scroll hints (Scroll, ↓)

### Design Vocabulary (pattern → trigger condition → content.json mapping)

**Covers**: Asymmetric Split Hero → `goal:"hook"`, image on one side | Editorial Manifesto → `goal:"hook"`, no image, large type | Full-Bleed Image → `goal:"hook"`, image with overlay | Data-Impact → `goal:"hook"`, big number + one-liner | Minimal Typography → `goal:"hook"`, text-only, extreme whitespace

**Inner pages**: Sidebar+Content → `--layout-variant sidebar-left` in generate_ppt(), consulting/reports | Split Text-Image → `goal:"content"`, image field | Bento Grid → `goal:"features"`, 4+ cards | Big Number Focus → `goal:"data"`, single metric | Card Row → `goal:"features"`, 3 cards | Comparison Split → `goal:"content"`, two-column contrast | Timeline Horizontal → `component_category:"timeline"` | Quote Spotlight → `goal:"content"`, quote in bullets | Code Terminal → `goal:"code"` | Full-Width Visual → `goal:"content"`, full-bleed image

**Data**: Table Diagram → `goal:"data"`, diagram type:"table" | Chart Focus → `goal:"data"`, diagram type: chart | Metric Dashboard → `component_category:"infographic"` | Infographic Component → `component_type:"group"` | Number Grid → `goal:"data"`, 2×2 metrics

**Content relationship → visual strategy**: Sequential → Timeline | Contrast → Comparison Split | Primary+secondary → unequal layout | Equal-weight → Card Row | Hierarchical → Hierarchy component | Evidence → center + orbit | Process → Cycle/Process component | Data-driven → Big Number/Chart

### Redesign Protocol
- Greenfield: start from Dial baseline
- Redesign-Preserve: audit brand tokens → incremental evolution
- Redesign-Overhaul: visually equivalent to greenfield
- Audit before modifying: brand tokens / information architecture / content blocks / patterns to keep / patterns to kill
- Modernization levers (in order): fonts → spacing → colors → animations → key-page rebuild → full replacement
- Never silently change: page order / navigation labels / logo / legal copy

### Design System Mapping
- Consulting/finance → sidebar + component_library process/hierarchy
- Tech talks → code block + component_library infographic
- Education → exercise page + built-in bullets
- Creative proposals → custom blocks + AI images
- Brand launch → full-bleed images + minimal text
- ONE design system per deck — no mixing McKinsey sidebar with Cyberpunk neon

### Performance & Accessibility
- <50 shapes per slide | images: cover-fit crop, never stretch | cache-first
- Public-sector / accessibility scenes: motion ≤3 | unknown audience: motion ≤5
- Dark mode: no pure black (#000) or pure white (#FFF) — use near-black/near-white
- Z-order: background < content < decoration < overlay

### Scope Exclusions
- Pure data tables → Excel | Multi-step forms → Web app | Real-time collaboration → Dedicated app | Interactive dashboards → Power BI/Tableau | Long documents (>50 pages) → Word/PDF

### Pre-Flight Check
- [ ] Basics: fonts ≥11pt | pages ≥3 shapes | text on every page | images correct | no broken links
- [ ] Consistency: accent color throughout | corner radius uniform | font pair uniform | theme locked | font-size levels ≥4
- [ ] Typography: cover title ≤2 lines | subtitle ≤20 chars | inner title ≤2 lines | no rotated text | line-height ≥1.1 for italic
- [ ] Layout: adjacent pages different layout family | 8+ pages ≥4 layouts | first card featured | no 3 consecutive same-structure pages | alternation ≤2
- [ ] Labels: eyebrow ≤ceil(pages/3) | no numbered eyebrows | no image overlay labels | no status dots | no generic step labels
- [ ] Color: no default-blue cover (unless specified) | no default gold+navy | chart colors from palette | dark-theme contrast sufficient | no light-gray on white
- [ ] Rhythm: 6+ pages have divider | hook ≠ cta visually | density varies | high-density followed by low | ≤1 core message per page
- [ ] Content: no AI Tells violations | no fake precision numbers | bullets have logical relation | quotes ≤3 lines with attribution
- [ ] Visuals: cover has real visual | minimalist ≥2-3 pages with images | bento ≥2-3 cells varied | logo wall uses images
- [ ] Scene: projection contrast OK | print doesn't rely on animation | large-screen numbers ≥36pt
- [ ] Animation: each has stated purpose | motion>4 has real animations | marquee ≤1/page
- [ ] Dial: values derived from Design Read | variance>4 has asymmetric layouts | density varies across pages

## Python API

```python
from ppt_pro_max import generate_ppt, fetch_image

# FreeStyle
result = generate_ppt("AI startup pitch", style="dark cyberpunk", fetch_images=True)

# With content.json
result = generate_ppt("pitch", content_file="content.json", style="warm fintech", fetch_images=True)

# With design dials
result = generate_ppt("pitch", content_file="content.json", style="professional",
                      layout_variant="sidebar-left", motion=5, density=6, variance=7)

# Proposal flow
result = generate_ppt("pitch", proposal=True, style="dark cyberpunk")

# Standalone image generation
img = fetch_image("futuristic AI city", mode="generate", llm_provider="seedream", llm_api_key="...")
print(img["path"])
```

**Key generate_ppt() parameters:** `query`, `style`, `content_file`, `layout_variant`, `variance`, `motion`, `density`, `fetch_images`, `proposal`, `confirmed_proposal`, `component_library`, `palette`, `fonts`, `decoration`, `mood`

## 4-Phase Pipeline

1. **Story Planning** → strategy + page structure + emotion arc
2. **Design Decisions** → per-page layout/color/typography from 40K+ combos
3. **Content Generation** → copy formulas (PAS/FAB/AIDA) + image keywords
4. **PPT Rendering** → python-pptx direct, 12 master layouts, QA gates

## Design Atoms (40,000+ Combos)

| Atom | Count | Examples |
|------|-------|----------|
| Color Palettes | 25 | ocean-blue, cyber-neon, golden-luxury, wine-burgundy, midnight-navy, monochrome-dark... |
| Font Pairs | 20 | modern-sans, elegant-serif, tech-mono, contrast-mix, sharp-modern... |
| Decorations | 10 | accent-bar, neon-lines, gold-trim, diamond-bullets, gradient-bar, sidebar-nav, minimal-dots, circle-accent, no-decoration, full-bleed-overlay |
| Layout Variants | 8 | standard, centered, sidebar-left, sidebar-right, grid-2x2, asymmetric... |

Natural language: `--style "warm fintech"` auto-selects matching atoms. Decoration and layout-variant atoms are consumed by PrecisionRenderer — they control title decoration, margin positioning, and card style.

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

Image modes: `placeholder` (default), `search` (Unsplash/Pexels), `generate` (AI), `enhance` (Kimi keyword optimization + search). All engines use **cache-first**.

## Animation System

- **12 Transitions**: fade, push, wipe, split, cover, dissolve, wheel, wedge, blinds, checker, comb, random
- **11 Entrances**: appear, fly_in, fade_in, zoom_in, float_up, bounce, etc.
- **Motion mapping**: 1-2 = transitions only, 3-5 = fade_in entrance, 6-10 = fly_in entrance
- Applied via XML injection (python-pptx 1.0.2 has no native transition API)

## Project Structure (Build & VI Build)

```
my-project/
├── template.pptx    # Brand template (VI Build only)
├── brand.json       # Brand spec (colors, fonts, logo, footer, watermark)
├── content.json     # Page content for FreeStyle/generate_ppt()
├── build.py         # Build Script or VI Build entry point
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

## Component Library (Professional Chart Templates)

A SQLite-indexed library of GroupShape/SmartArt templates extracted from real PPT files, with coordinate normalization for universal scaling.

### Library Overview

| Category | Count | Use When |
|----------|-------|----------|
| infographic | 4,101 | Data visualization, statistics, KPI dashboards |
| process | 672 | Workflows, step-by-step, pipelines, procedures |
| hierarchy | 548 | Org charts, reporting structures, tree diagrams |
| chart | 132 | Bar/pie/line chart layouts, data comparison |
| timeline | 42 | Milestones, roadmaps, chronological events |
| swot | 39 | Strategic analysis, 4-quadrant frameworks |

### How to Query the Library

```python
from ppt_pro_max import query_component_library

catalog = query_component_library()
results = query_component_library(type="group", category="process")
results = query_component_library(type="group", category="hierarchy", node_count=5)
```

### How to Use Components in content.json

Add `component_type` and `component_category` to any slide:

```json
{
  "goal": "content",
  "title": "项目流程",
  "bullets": ["需求分析", "方案设计", "开发实现", "测试上线"],
  "component_type": "group",
  "component_category": "process"
}
```

Matching logic: exact match (type+category+node_count) → fuzzy match (closest node_count) → fallback (DiagramEngine or bullets).

### Component Selection Strategy

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

**Node count matching:** node_count = number of text items (4 steps → node_count=4). Prefer exact match; if between two, choose larger.

## Build Helpers API (for Build/VI Build mode)

LLM writes `build.py` scripts using these functions. Import: `from ppt_pro_max.build_helpers import *`

### Color Dictionary (C)

```python
C = {
    'primary': '#2E6504', 'accent': '#7DA92F', 'muted': '#81C784',
    'light': '#C8E6C9', 'white': '#FFFFFF', 'background': '#FFFFFF',
    'card_bg': '#F9F9F9', 'text_dark': '#1A1A1A', 'text_body': '#333333',
    'text_muted': '#666666', 'divider': '#CCCCCC',
    'font_heading': '微软雅黑', 'font_body': '微软雅黑',
}
```

### Functions

| Function | Purpose | Key Params |
|----------|---------|------------|
| `add_slide(prs)` | Add blank slide | Auto-finds blank layout |
| `page_header(slide, title, subtitle, C)` | Title + subtitle + divider line | Standard content page header |
| `kpi_card(slide, left, top, width, height, number, label, trend, C)` | KPI metric card | number (big), label, trend (+8.3%) |
| `bar_chart(slide, left, top, data, C)` | Horizontal bar chart | data: [(label, pct, val), ...] |
| `comparison_bars(slide, left, top, metrics, C)` | Before/after comparison | metrics: [(label, v_old, v_new, pct_old, pct_new), ...] |
| `donut_chart(slide, cx, cy, radius, inner_radius, sectors, C)` | Donut chart (simplified) | sectors: [(name, pct_str, color), ...] |
| `highlight_cards(slide, left, top, cards, C)` | Highlight card row | cards: [(title, desc, accent_color), ...] |
| `text(slide, left, top, width, height, txt, font_size, color, bold, align, font_name, C)` | Single-line text | color: role name or hex |
| `multiline(slide, left, top, width, height, lines, font_size, color, C)` | Multi-line text | lines: list of strings |
| `rect(slide, left, top, width, height, fill, line, C)` | Rectangle | fill/line: role name or hex |
| `rrect(slide, left, top, width, height, fill, line, C)` | Rounded rectangle | Same as rect |
| `oval(slide, left, top, width, height, fill, line, C)` | Ellipse | Same as rect |
| `top_bar(slide, color, C)` | Top accent bar | Brand color strip |
| `copy_decorations(slide, template_slide)` | Copy decorations from template | Skips long text (>50 chars) and images |
| `copy_logo(slide, template_slide, color_hints)` | Copy LOGO from template | Only finds GROUP shapes (shape_type==6) |

### Color Resolution
- Hex value: `'#2E6504'` → used directly
- Role name: `'primary'` → looks up `C['primary']`
- Missing role: returns `'#000000'` (never crashes)

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
- **layout_variant**: NOT a content.json field — use `--layout-variant` CLI param or `layout_variant=` kwarg in generate_ppt()
- **animation**: NOT a content.json field — use `--motion` CLI param or `motion=` kwarg in generate_ppt()

## CLI Quick Reference

```
python -m ppt_pro_max "query" [--style STYLE] [--layout-variant VARIANT] [--motion 1-10] [--density 1-10] [--variance 1-10] [--content FILE] [--fetch-images] [--proposal] [-o PATH]
python -m ppt_pro_max image "keywords" [--llm-provider PROV] [--llm-api-key KEY] [--image-mode MODE] [-v]
```

## Dependencies

- python-pptx >= 1.0.2 (required)
- Pillow >= 10.0 (required)
- python-dotenv >= 1.0 (optional, for .env support)
- ui-ux-pro-max >= 1.0.0 (optional)
