<div align="center">

# PPT Design Skill

**Generate professional .pptx presentations from a single sentence**

Triple-mode engine · Narrative-driven · Brand-compliant · AI images · **40,000+ Style Combinations** · **28 Design Quality Upgrades** · **Build Script Precision**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![pptx](https://img.shields.io/badge/python--pptx-1.0.2-green.svg)](https://pypi.org/project/python-pptx/)

Compatible with OpenCode · Claude Code · Codex · Cursor

[中文](../README.md) | [Usage Guide](usage-guide.md) | English

</div>

---

## ✨ Showcase

> 5 styles, 5 scenarios — each with cover + content page, AI images by Seedream

### 🏢 Professional Modern — Enterprise Investor Pitch

<img src="showcase/showcase-professional-slide1.jpg" width="45%"/> <img src="showcase/showcase-professional-slide2.jpg" width="45%"/>

*Navy blue corporate · Gold accents · Left sidebar navigation · 2×2 metric cards*

### 🌌 Dark Tech — AI Product Launch

<img src="showcase/showcase-dark-tech-slide1.jpg" width="45%"/> <img src="showcase/showcase-dark-tech-slide2.jpg" width="45%"/>

*Cyberpunk dark · Neon blue/purple/pink · Consolas monospace · 3-column feature cards*

### 🏛️ Warm Elegant — Luxury Brand Strategy

<img src="showcase/showcase-warm-elegant-slide1.jpg" width="45%"/> <img src="showcase/showcase-warm-elegant-slide2.jpg" width="45%"/>

*Golden marble · Georgia serif · Centered editorial layout · Diamond bullet points*

### 🚀 Vibrant Startup — Fundraising Pitch Deck

<img src="showcase/showcase-vibrant-startup-slide1.jpg" width="45%"/> <img src="showcase/showcase-vibrant-startup-slide2.jpg" width="45%"/>

*Purple-pink gradient · Segoe UI · Progress bar metrics · Semi-transparent stat pills*

### 🌿 Nature Calm — Sustainability Impact Report

<img src="showcase/showcase-nature-calm-slide1.jpg" width="45%"/> <img src="showcase/showcase-nature-calm-slide2.jpg" width="45%"/>

*Forest green · Circle accents · 4-column impact cards · Narrow left sidebar*

---

## 🔥 Features

| Feature | Description |
|---------|-------------|
| **Triple-Mode Engine** | FreeStyle rapid generation + Enterprise brand compliance + Build Script per-page precision |
| **Narrative Engine** | 3 strategies (YC Seed Deck / Product Demo / Sales Pitch) + Duarte Sparkline emotion arcs |
| **40,000+ Style Combos** | 25 palettes × 20 font pairs × 10 decorations × 8 layout variants |
| **Natural Language Style** | Describe your style: `--style "warm fintech"` / `--style "dark cyberpunk"` |
| **28 Design Quality Upgrades** | OKLCH color depth · Shadow elevation · Gradient overlay · Progress bar · Corner radius · CJK font pairing · Noise texture · Two-column bullets · 4 Hero patterns · Section dividers · Badge system · Gradient lines · Image masking · Decoration renderer · Code block redesign · Card upgrade · Adaptive margins · Typography scale |
| **10 Diagram Types** | Flowchart / Funnel / Timeline / SWOT / Matrix / Cycle / Table / Hierarchy / Pyramid / Venn |
| **Build Script Mode** | 10 page templates + Design Token system + post-build checks, delivery-grade quality |
| **Brand Visual Design** | Auto-mapped brand colors: backgrounds, accent bars, branded title/body text, logo placement |
| **Page Revision** | `--pages` CRUD: delete / swap / move / insert with full content preservation |
| **Version Management** | v1 → v2 → v3 auto-numbering, meta.json tracks per-page goal/title |
| **python-pptx Direct** | Fully editable .pptx output, 356x faster than HTML→screenshot |
| **12 Master Layouts** | 13.333"×7.5" 16:9 precise coordinates |
| **AI Image Engines** | Seedream / GPT Image / DALL-E / Wanx — 4 generation engines + Kimi enhancement |
| **Animation System** | 12 transition types + 10 entrance effects, motion 1-10 smart mapping |
| **Code/Exercise Blocks** | Dark code blocks + language badges + exercise badges + step lists — education-ready |
| **CJK Fonts** | 12 CJK font pairings with auto-fallback (Microsoft YaHei / STSong / SimHei etc.) |
| **QA Gates** | 5 automated quality checks + `--review` proposal confirmation |

---

## 🚀 Quick Start

### Install

```bash
git clone https://github.com/sunchaokun/PPT-Design-Skill.git
cd PPT-Design-Skill

# One-click install — auto-detect AI platform + install skill + pip deps
python install.py

# Manual install
pip install -e .
```

### FreeStyle — Generate from a Sentence

```bash
# Minimal usage
ppt-design "AI startup investor pitch"

# Natural language style — 40,000+ combinations
ppt-design "fintech pitch" --style "warm fintech"
ppt-design "product launch" --style "dark cyberpunk tech"
ppt-design "ESG report" --style "calm nature"

# AI images + style + animation
ppt-design "investor pitch" --style "dark cyberpunk" \
  --fetch-images --llm-provider seedream --llm-api-key $ARK_API_KEY \
  --motion 7 --density 6
```

### Enterprise — Brand-Compliant + Version Control

```bash
# 1. Create project directory
mkdir my-pitch && cd my-pitch

# 2. Add brand assets (all optional)
#    template.pptx / brand.json / content.json / logo.png / images/

# 3. Generate
ppt-design "AI Platform" --project . --density 6 --motion 5

# 4. Revise pages
ppt-design "" --project . --pages "-3,2<>5"

# 5. View history
ppt-design "" --project . --history
```

### Build Script — Per-Page Precision (Delivery-Grade Quality)

```python
# build.py — Direct python-pptx per-page precise control
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
sl = prs.slides.add_slide(prs.slide_layouts[6])

# Every element: exact x, y, w, h, font, size, color
tb = sl.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(8), Inches(1.5))
run = tb.text_frame.paragraphs[0].add_run()
run.text = "Title"
run.font.name = "Space Grotesk"  # run-level font, PowerPoint respects
run.font.size = Pt(52)
run.font.bold = True

prs.save("output/presentation.pptx")
```

> 10 page templates + Design Token system + post-build check script — see [Usage Guide §5](usage-guide.md#5-build-script-模式--精确控制)

---

## 🏗️ Triple-Mode Architecture

### FreeStyle Pipeline

```
Input → Story Planning → Design Decisions → Content Generation → PPT Rendering → .pptx
```

Use for: quick exploration, style experiments, personal presentations

### Enterprise Pipeline

```
Project Dir → Asset Scan → Brand Merge → Content Parse → Visual Design → Render+Animate → Version Mgmt → .pptx
                 ↓            ↓             ↓              ↓
            template     brand.json    content.json    accent bar
            logo.png     colors/fonts  diagram         brand overlay
            images/      logo config   code/exercise   footer/watermark
```

Use for: enterprise compliance, team collaboration, multi-version iteration

### Build Script Mode

```
build.py → Design Token → Page Templates → python-pptx → .pptx → check.py validation
              ↓              ↓              ↓
          colors/fonts   10 templates   per-element precision
          one-line theme  x/y/w/h      run-level fonts
```

Use for: **final delivery, precise control, quality assurance**

> **Recommended workflow**: FreeStyle prototype → Enterprise content fill → Build Script precision delivery

### Project Directory Structure

```
my-project/
├── template.pptx      # Optional: brand template
├── brand.json         # Optional: brand specification
├── content.json       # Optional: real content
├── logo.png           # Optional: company logo
├── images/            # Optional: image pool
└── output/            # Auto-generated
    ├── v1/presentation.pptx + meta.json
    └── v2/presentation.pptx + meta.json
```

---

## 📋 content.json — Precise Content Control

```json
{
  "meta": {"title": "Acme Corp — Series B Pitch"},
  "slides": [
    {"goal": "hook", "title": "The Future of AI", "subtitle": "Acme Corp", "image": "images/hero.png"},
    {"goal": "problem", "title": "The Problem", "bullets": ["75% fail", "Data silos"]},
    {"goal": "solution", "title": "Our Solution", "bullets": ["Unified gateway"], "image": "images/product.png"},
    {"goal": "features", "title": "Key Features", "cards": [
      {"title": "Fast", "text": "Sub-100ms"},
      {"title": "Secure", "text": "SOC2+HIPAA"}
    ]},
    {"goal": "market", "title": "Market", "diagram": {"type": "funnel", "data": {"items": [{"text": "TAM $120B"}, {"text": "SOM $8B"}]}}},
    {"goal": "code_demo", "title": "Quick Start", "code": {"language": "python", "source": "from acme import AIPlatform\nplatform = AIPlatform(key='x')"}},
    {"goal": "exercise", "title": "Try It", "exercise": {"instructions": "Deploy in 5 min", "duration": "5 min", "steps": ["Sign up", "Deploy"]}},
    {"goal": "cta", "title": "Join Us", "subtitle": "contact@acme.ai"}
  ]
}
```

> Full field reference: [Usage Guide §5](usage-guide.md#5-contentjson-内容格式)

---

## 🎨 brand.json — Brand Visual Specification

```json
{
  "colors": {
    "primary": "#2563EB",
    "accent": "#F97316",
    "foreground": "#1A1A2E",
    "muted-foreground": "#94A3B8",
    "background": "#FFFFFF",
    "muted": "#F1F5F9"
  },
  "fonts": {"heading": "Calibri", "body": "Calibri"},
  "logo": {"position": "top_right", "width_inches": 1.2, "skip_slides": ["hook"]},
  "footer": {
    "show_page_number": true,
    "page_number_format": "{n} / {total}"
  }
}
```

> Full field reference: [Usage Guide §6](usage-guide.md#6-brandjson-品牌格式)

---

## ✏️ Page Revision — CRUD Operations

```bash
# Delete page 3
ppt-design "" --project . --pages "-3"

# Swap pages 2 and 5
ppt-design "" --project . --pages "2<>5"

# Move page 10 to after page 3
ppt-design "" --project . --pages "10>3"

# Insert blank page after page 6
ppt-design "" --project . --pages "+6"

# Combined operations
ppt-design "" --project . --pages "-3,2<>5,10>3,+6"
```

> Full syntax: [Usage Guide §7](usage-guide.md#7-页面修订语法)

---

## 🖼️ Image Engines

| Engine | Type | CLI | Default Model |
|--------|------|-----|---------------|
| `placeholder` | Gradient placeholder | Default | — |
| `search` | Unsplash / Pexels | `--image-mode search` | — |
| `seedream` | AI generate | `--llm-provider seedream` | `doubao-seedream-5-0-260128` |
| `gpt-image` | AI generate | `--llm-provider gpt-image` | `gpt-image-1` |
| `dalle` | AI generate | `--llm-provider dalle` | `dall-e-3` |
| `wanx` | AI generate | `--llm-provider wanx` | `wanx-v1` |
| `kimi` | Enhanced search | `--llm-provider kimi` | `kimi-k2-0711-preview` |

### Seedream Available Models

| Model | Description |
|-------|-------------|
| `doubao-seedream-5-0-260128` | **Default**, Seedream 5.0 |
| `doubao-seedream-5-0-pro-260628` | Seedream 5.0 Pro |
| `doubao-seedream-4-5-251128` | Seedream 4.5 |

All AI engines include **cache-first** — same image never generated twice.

---

## 🎨 Design System — 40,000+ Style Combinations

| Atom | Count | Examples |
|------|-------|----------|
| 🎨 Color Palettes | 25 | ocean-blue, cyber-neon, golden-luxury, wine-burgundy, monochrome-dark... |
| ✏️ Font Pairs | 20 | modern-sans, serif-editorial, tech-mono, elegant-serif, contrast-mix... |
| 🖌️ Decorations | 10 | accent-bar, neon-lines, gold-trim, gradient-bar, circle-accent, sidebar-nav... |
| 📐 Layout Variants | 8 | standard, centered, sidebar-left, grid-2x2, wide-cards, full-width... |

**25 × 20 × 10 × 8 = 40,000 combinations**

### Natural Language Style

```bash
ppt-design "investor pitch" --style "warm fintech"          # → ocean-blue + clean-corporate + accent-bar
ppt-design "product launch" --style "dark cyberpunk"         # → cyber-neon + tech-mono + neon-lines
ppt-design "brand strategy" --style "elegant luxury"         # → golden-luxury + elegant-serif + gold-trim
ppt-design "ESG report" --style "calm nature"                # → sage-calm + humanist-sans + circle-accent
ppt-design "startup pitch" --style "bold startup vibrant"    # → royal-purple + bold-sans + gradient-bar
```

### Preset Themes

| Theme | Palette | Fonts | Decoration | Layout |
|-------|---------|-------|------------|--------|
| Professional | midnight-navy | clean-corporate | accent-bar | sidebar-left |
| Dark Tech | cyber-neon | tech-mono | neon-lines | wide-cards |
| Warm Elegant | golden-luxury | serif-editorial | gold-trim | centered |
| Vibrant Startup | neon-gradient | bold-sans | gradient-bar | grid-2x2 |
| Nature Calm | forest-green | humanist-sans | circle-accent | sidebar-left |

---

## 🏆 Design Quality Upgrades — 28 Professional Enhancements

v0.7.0 introduces 28 design quality upgrades across three tiers:

### Tier 1 — Visual Foundations (10)

| # | Upgrade | Description |
|---|---------|-------------|
| 1.1 | **Layout Engine** | `LayoutEngine` + `Rect` + `ContentLayout` — unified coordinate computation |
| 1.2 | **Typography Scale** | `TypeScale` dataclass with density/mode-adaptive font size ratios |
| 1.3 | **OKLCH Color Depth** | Perceptually uniform 9-level color scale + 5-level alpha hierarchy |
| 1.4 | **Gradient Overlay** | `add_gradient_overlay()` replaces flat overlay — softer cover/CTA pages |
| 1.5 | **5-Level Shadow Elevation** | `ELEVATION_SCALE` from subtle→floating — clear visual hierarchy |
| 1.6 | **Smart Brand Strip** | ~1/3 of pages skip brand color strip — avoids visual fatigue |
| 1.7 | **Image Color Grading** | `grade_image_to_palette()` unifies image tones to brand palette |
| 1.8 | **Card Upgrade** | Title 20pt / body 14pt, `featured` cards get gradient bar |
| 1.9 | **Dark Mode Fix** | OKLCH luminance detection (0.299R+0.587G+0.114B) — no more misdetection |
| 1.10 | **Code Block Redesign** | Always-dark bg #1E293B + separate language badge + muted text color |

### Tier 2 — Typography Enhancements (6)

| # | Upgrade | Description |
|---|---------|-------------|
| 2.1 | **CJK Font Pairing** | 12 Latin+CJK combinations (e.g. Space Grotesk + Microsoft YaHei), auto-fallback |
| 2.2 | **Adaptive Margins** | presenting (0.6") / reading (1.2") / balanced (0.9") density-aware margins |
| 2.3 | **Badge System** | `add_badge()` — uppercase, 3 variants (default/solid/outline), CJK width-aware |
| 2.4 | **Section Dividers** | `goal="section"` → oversized number + title + gradient line |
| 2.5 | **Decoration Renderer** | `DecorationRenderer` unifies all 10 decoration styles |
| 2.6 | **Layout Variant Consumption** | `render_slide()` reads content_margin_left / title_alignment / decoration_style |

### Tier 3 — Advanced Visual (7)

| # | Upgrade | Description |
|---|---------|-------------|
| 3.1 | **Noise Texture** | Per-deck seeded Gaussian noise — adds subtle texture |
| 3.2 | **Progress Bar** | Thin bottom progress indicator, replaces simple page numbers |
| 3.3 | **Corner Radius System** | 4 levels (sm=4 / md=8 / lg=12 / pill=50) + `add_rounded_rect()` |
| 3.4 | **Gradient Lines** | `add_gradient_line()` — alpha fade for elegant title underlines |
| 3.5 | **Image Masking** | `add_masked_image()` — rounded-rect frame with 0.15" padding |
| 3.6 | **Two-Column Bullets** | 6+ bullets auto-split into 2 columns + vertical separator |
| 3.7 | **4 Hero Patterns** | gradient / split-left / bottom-fade / asymmetric cover layouts |

---

## 📁 Project Structure

```
PPT-Design-Skill/
├── pyproject.toml
├── install.py                        # One-click installer
├── SKILL.md                          # AI skill definition
├── AGENTS.md                         # Project instructions
├── docs/
│   ├── README_EN.md                  # This file
│   ├── usage-guide.md                # Full usage guide (Chinese, incl. Build Script)
│   └── showcase/                     # Showcase screenshots
├── examples/                         # Example PPTs
│   ├── showcase-professional.pptx
│   ├── showcase-dark-tech.pptx
│   ├── showcase-warm-elegant.pptx
│   ├── showcase-vibrant-startup.pptx
│   └── showcase-nature-calm.pptx
├── src/ppt_pro_max/
│   ├── __init__.py                   # generate_ppt() API
│   ├── cli.py                        # ppt-design CLI
│   ├── enterprise/                   # Enterprise Pipeline
│   │   ├── pipeline.py               # Main orchestrator
│   │   ├── precision_renderer.py     # Unified renderer (8 goal layouts + 28 design upgrades)
│   │   ├── brand_spec.py             # Brand specification
│   │   ├── content_parser.py         # Content parser (content.json + README.md)
│   │   ├── image_matcher.py          # Image matching + size classification + AI prompts
│   │   ├── proposal_generator.py     # Style preview (2-3 proposals)
│   │   ├── slide_extractor.py        # PPT content extraction (beautify mode)
│   │   ├── component_library.py      # Component library (SQLite index + dedup)
│   │   ├── component_renderer.py     # Component rendering bridge
│   │   ├── page_revision.py          # Page revision engine
│   │   ├── density_profile.py        # Density profiles
│   │   └── ...
│   ├── renderer/
│   │   ├── ppt_renderer.py           # FreeStyle renderer
│   │   ├── diagram_engine.py         # 10-type diagram engine
│   │   ├── diagram/                  # Diagram implementations
│   │   │   ├── flowchart.py / funnel.py / timeline.py / swot.py
│   │   │   ├── matrix.py / cycle.py / table.py
│   │   │   ├── hierarchy.py / pyramid.py / venn.py
│   │   │   └── connector_router.py / text_measurer.py / ...
│   │   ├── animation.py              # 12 transitions + 10 entrance effects
│   │   ├── theme_composer.py         # 40,000+ style combinations + 35 moods
│   │   ├── typography.py             # Typography scale (TypeScale)
│   │   ├── color_system.py           # OKLCH color depth + alpha levels
│   │   ├── elevation.py              # 5-level shadow elevation
│   │   ├── layout_engine.py          # Layout engine + adaptive margins
│   │   ├── image_processor.py        # Image color grading + noise texture
│   │   ├── decoration_renderer.py    # 10 decoration style renderer
│   │   ├── visual_effects.py         # Gradient / glow / letter spacing
│   │   ├── image_fetcher.py          # 5 engines + caching
│   │   └── ...
│   ├── planner/story_planner.py      # Story planning
│   ├── decider/design_decider.py     # Design decisions
│   └── content/content_generator.py  # Content generation
├── tests/                            # 824 tests
└── e2e-test-project/                 # E2E test project
```

## License

MIT
