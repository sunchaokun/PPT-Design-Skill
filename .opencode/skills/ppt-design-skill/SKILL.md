---
name: ppt-design-skill
version: 0.3.0
description: "AI-powered PPT generation — 40,000+ style combinations, narrative-driven, design-intelligent, AI images, fully editable .pptx. Dual-mode: FreeStyle + Enterprise. 10 diagram types, brand compliance, version control, page revision. Engines: Seedream, GPT Image, DALL-E, Wanx, Kimi."
argument-hint: "[topic] [--style style-description] [--fetch-images] [--project DIR]"
license: MIT
metadata:
  author: sunchaokun
  category: design
  tags: [ppt, presentation, deck, pitch, slides, python-pptx, enterprise, brand, diagram]
---

# PPT Design Skill

AI-powered PPT generation — dual-mode engine, 40,000+ style combos, 10 diagram types, brand compliance, version control, fully editable .pptx.

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
from ppt_pro_max import generate_ppt

# FreeStyle
result = generate_ppt("AI startup pitch", style="dark cyberpunk", fetch_images=True)

# Enterprise
result = generate_ppt("路演", project="./my-project", density=6)

# Page revision
result = generate_ppt("", project=".", pages="-3,2<>5")

# Version history
result = generate_ppt("", project=".", history=True)

print(f"Generated: {result['output_path']}, {result['page_count']} pages")
```

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
| Decorations | 10 | accent-bar, neon-lines, gold-trim, diamond-bullets, gradient-bar, sidebar-nav... |
| Layout Variants | 8 | standard, centered, sidebar-left, sidebar-right, grid-2x2, asymmetric... |

Natural language: `--style "warm fintech"` auto-selects matching atoms.

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
      "goal": "hook|problem|solution|features|cta|...",
      "title": "Page Title",
      "subtitle": "Optional subtitle",
      "bullets": ["Point 1", "Point 2"],
      "image": "images/photo.png",
      "cards": [{"title": "...", "text": "..."}],
      "diagram": {"type": "flowchart", "data": {...}},
      "code": {"language": "python", "source": "..."},
      "exercise": {"instructions": "...", "duration": "5 min", "steps": [...]}
    }
  ]
}
```

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
  "watermark": {"text": "CONFIDENTIAL", "opacity": 0.1}
}
```

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

# Images
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

# Output
  --persist             Save design system as MASTER.md
  --dry-run             Output design decisions only
  -o, --output PATH     Output .pptx path
```

## Key Constraints

- **python-pptx 1.0.2**: No `PP_TRANSITION_TYPE`, must use XML for transitions/animations
- **Cover-fit images**: Use `_add_picture_cover()` with Pillow pre-crop — never stretch
- **Cache-first**: All image engines check cache before API call
- **Two-pass rebuild**: Page revision uses rebuild (not in-place) to avoid ZIP corruption
- **1-based pages**: All `--pages` numbers refer to original document
- **Windows**: Use `python` not `python3`

## Dependencies

- python-pptx >= 1.0.2 (required)
- Pillow >= 10.0 (required)
- python-dotenv >= 1.0 (optional, for .env support)
- ui-ux-pro-max >= 1.0.0 (optional)
