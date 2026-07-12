<div align="center">

# PPT Design Skill

**Generate professional .pptx presentations from a single sentence**

Dual-mode engine · Narrative-driven · Brand-compliant · AI images · **40,000+ Style Combinations**

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
| **Dual-Mode Engine** | FreeStyle rapid generation + Enterprise brand compliance + version management + page revision |
| **Narrative Engine** | 3 strategies (YC Seed Deck / Product Demo / Sales Pitch) + Duarte Sparkline emotion arcs |
| **40,000+ Style Combos** | 25 palettes × 20 font pairs × 10 decorations × 8 layout variants |
| **Natural Language Style** | Describe your style: `--style "warm fintech"` / `--style "dark cyberpunk"` |
| **10 Diagram Types** | Flowchart / Funnel / Timeline / SWOT / Matrix / Cycle / Table / Hierarchy / Pyramid / Venn |
| **Brand Visual Design** | Auto-mapped brand colors: backgrounds, accent bars, branded title/body text, logo placement |
| **Page Revision** | `--pages` CRUD: delete / swap / move / insert with full content preservation |
| **Version Management** | v1 → v2 → v3 auto-numbering, meta.json tracks per-page goal/title |
| **python-pptx Direct** | Fully editable .pptx output, 356x faster than HTML→screenshot |
| **12 Master Layouts** | 13.333"×7.5" 16:9 precise coordinates |
| **AI Image Engines** | Seedream / GPT Image / DALL-E / Wanx — 4 generation engines + Kimi enhancement |
| **Animation System** | 12 transition types + 10 entrance effects, motion 1-10 smart mapping |
| **Code/Exercise Blocks** | Dark code blocks + exercise badges + step lists — education-ready |
| **CJK Fonts** | East Asian font fallback chain (Microsoft YaHei / STSong) |
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

---

## 🏗️ Dual-Mode Architecture

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

## 📁 Project Structure

```
PPT-Design-Skill/
├── pyproject.toml
├── install.py                        # One-click installer
├── SKILL.md                          # AI skill definition
├── AGENTS.md                         # Project instructions
├── docs/
│   ├── README_EN.md                  # This file
│   ├── usage-guide.md                # Full usage guide (Chinese)
│   └── showcase/                     # Showcase screenshots
├── src/ppt_pro_max/
│   ├── __init__.py                   # generate_ppt() API
│   ├── cli.py                        # ppt-design CLI
│   ├── enterprise/                   # Enterprise Pipeline
│   │   ├── pipeline.py               # Main orchestrator
│   │   ├── enterprise_renderer.py    # Enterprise renderer
│   │   ├── brand_spec.py             # Brand specification
│   │   ├── content_parser.py         # Content parser
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
│   │   ├── theme_composer.py         # 40,000+ style combinations
│   │   ├── image_fetcher.py          # 5 engines + caching
│   │   └── ...
│   ├── planner/story_planner.py      # Story planning
│   ├── decider/design_decider.py     # Design decisions
│   └── content/content_generator.py  # Content generation
├── tests/                            # 412 tests
└── e2e-test-project/                 # E2E test project
```

## License

MIT
