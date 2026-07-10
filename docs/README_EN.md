<div align="center">

# PPT Design Skill

**Generate professional .pptx presentations from a single sentence**

Narrative-driven · Design-intelligent · AI images · Fully editable · **40,000+ Style Combinations**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![pptx](https://img.shields.io/badge/python--pptx-1.0.2-green.svg)](https://pypi.org/project/python-pptx/)

Compatible with OpenCode · Claude Code · Codex · Cursor

[中文](../README.md) | English

</div>

---

## ✨ Showcase

> 5 styles, 5 scenarios — each with cover + content page, AI images by Seedream Pro

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
| **Narrative Engine** | 3 strategies (YC Seed Deck / Product Demo / Sales Pitch) + Duarte Sparkline emotion arcs |
| **Design Intelligence** | Reuses ui-ux-pro-max 5100+ design knowledge base for context-aware decisions |
| **40,000+ Style Combos** | 25 palettes × 20 font pairs × 10 decorations × 8 layout variants |
| **Natural Language Style** | Describe your style: `--style "warm fintech"` / `--style "dark cyberpunk"` |
| **python-pptx Direct** | Fully editable .pptx output, 356x faster than HTML→screenshot |
| **12 Master Layouts** | 13.333"×7.5" 16:9 precise coordinates, covering 95% of scenarios |
| **9 Copy Formulas** | PAS / FAB / AIDA / Social Proof / Cost of Inaction / Proof Stack... |
| **AI Image Engines** | Seedream / GPT Image / DALL-E / Wanx — 4 generation engines + Kimi K2.6 image enhancement |
| **Design Dials** | variance / motion / density 3-axis control |
| **CJK Fonts** | East Asian font fallback chain (Microsoft YaHei / STSong) |
| **QA Gates** | 5 automated quality checks (pages / titles / fonts / consistency / placeholders) |

---

## 🚀 Quick Start

### One-Click Install (Recommended)

```bash
# Clone repository
git clone https://github.com/sunchaokun/PPT-Design-Skill.git
cd PPT-Design-Skill

# One-click install — auto-detect AI platform + install skill files + pip deps
python install.py

# Specify platform
python install.py --platform claude     # Claude Code only
python install.py --platform opencode   # OpenCode only
python install.py --platform all        # All platforms

# Check installation status
python install.py --check

# Global install (install to home directory)
python install.py --global
```

The installer automatically:
1. **Detects** which AI coding platforms are in use (Claude Code / OpenCode / Codex / Cursor...)
2. **Copies** SKILL.md + scripts to each platform's skills directory
3. **Installs** Python dependencies (python-pptx, Pillow)
4. **Generates** AGENTS.md / CLAUDE.md project instruction files

### Manual Install

```bash
pip install -e .
# Optional: search engine support
pip install -e ".[search]"
```

### CLI

```bash
# Generate PPT from a single sentence
ppt-design "AI startup investor pitch"

# Natural language style (40K+ combos)
ppt-design "fintech pitch" --style "warm fintech"
ppt-design "product launch" --style "dark cyberpunk"
ppt-design "brand strategy" --style "elegant luxury"
ppt-design "ESG report" --style "calm nature"

# Exact atom control
ppt-design "pitch" --palette wine-burgundy --fonts elegant-serif --decoration gold-trim --layout-variant centered

# Specify strategy and theme (backward compatible)
ppt-design "SaaS product demo" --strategy "Product Demo" --theme "dark-tech"

# AI-generated images (Seedream Pro)
ppt-design "investor pitch" --fetch-images --llm-provider seedream --llm-api-key YOUR_ARK_KEY

# AI-generated images (GPT Image)
ppt-design "product intro" --fetch-images --llm-provider gpt-image --llm-api-key sk-xxx

# Kimi K2.6 enhances search keywords → downloads high-quality images
ppt-design "investor pitch" --image-mode enhance --llm-provider kimi --llm-api-key sk-xxx

# Load real content from JSON
ppt-design "investor pitch" --content pitch-data.json

# Design dials
ppt-design "investor pitch" --variance 8 --motion 6 --density 7

# Dry-run (design decisions only)
ppt-design "investor pitch" --dry-run
```

### Python API

```python
from ppt_pro_max import generate_ppt

# Minimal usage
result = generate_ppt("AI startup investor pitch")
print(f"Generated: {result['output_path']}, {result['page_count']} pages")

# Natural language style — infinite combinations
result = generate_ppt("fintech pitch", style="warm fintech")
result = generate_ppt("product launch", style="dark cyberpunk")
result = generate_ppt("brand strategy", style="elegant luxury minimal")

# Exact atom control
result = generate_ppt(
    query="AI startup investor pitch",
    palette="wine-burgundy",
    fonts="elegant-serif",
    decoration="gold-trim",
    layout_variant="centered",
)

# Full configuration — AI images + style combo
result = generate_ppt(
    query="AI startup investor pitch",
    strategy="YC Seed Deck",
    theme="dark-tech",
    slides=12,
    fetch_images=True,
    llm_provider="seedream",
    llm_api_key="ark-xxx",
    llm_model="doubao-seedream-5-0-pro-260628",
    variance=7,
    motion=5,
    density=6,
    output="my-pitch.pptx",
)
```

### Content JSON Format

```json
{
  "company": "Acme AI",
  "product": "AI Marketing Platform",
  "tagline": "Your AI marketing team. Always on.",
  "metrics": {"users": "10K+", "retention": "95%", "growth": "3x", "arr": "$2M"},
  "pain_points": [
    {"title": "Content Overload", "desc": "Need 10x content with same headcount"},
    {"title": "Tool Fatigue", "desc": "15+ tools that don't talk to each other"}
  ],
  "chart_data": {
    "mrr": {"labels": ["Sep","Oct","Nov","Dec"], "values": [5,12,28,45]}
  }
}
```

---

## 🏗️ 4-Layer Architecture

```
Input → Phase 1: Story Planning → Phase 2: Design Decisions → Phase 3: Content Generation → Phase 4: PPT Rendering → .pptx
               ↓                        ↓                          ↓                           ↓
        story_planner            design_decider            content_generator            ppt_renderer
        (strategy+emotion arc)  (layout+color+typography) (copy formulas+images)      (python-pptx direct)
```

| Layer | Module | Responsibility |
|-------|--------|----------------|
| Narrative | `planner/story_planner.py` | Strategy → Page structure → Emotion arc → Duarte Sparkline |
| Design | `decider/design_decider.py` | Layout → Color → Typography → Chart type → Transitions |
| Content | `content/content_generator.py` | Copy formulas → Data config → Image keywords → Variable filling |
| Expression | `renderer/ppt_renderer.py` | Theme → 12 layouts → Charts → Images → Animations → QA |

---

## 🖼️ Image Engines

| Engine | Type | CLI | Notes |
|--------|------|-----|-------|
| `placeholder` | Gradient placeholder | Default | No API key needed |
| `search` | Unsplash / Pexels | `--image-mode search` | API key required |
| `seedream` | AI generate | `--llm-provider seedream` | ByteDance Seedream Pro (recommended) |
| `gpt-image` | AI generate | `--llm-provider gpt-image` | OpenAI GPT Image |
| `dalle` | AI generate | `--llm-provider dalle` | OpenAI DALL-E 3 |
| `wanx` | AI generate | `--llm-provider wanx` | Alibaba Wanx |
| `kimi` | Enhanced search | `--llm-provider kimi` | Kimi K2.6 keyword optimization → search |

All AI generation engines include **cache-first** — same image never generated twice, saving costs.

---

## 🎨 Design Atoms — 40,000+ Style Combinations

> Compose 4 atom types for infinite styles:

| Atom | Count | Examples |
|------|-------|----------|
| 🎨 Color Palettes | 25 | ocean-blue, cyber-neon, golden-luxury, wine-burgundy, terracotta, monochrome-dark... |
| ✏️ Font Pairs | 20 | modern-sans, serif-editorial, tech-mono, elegant-serif, bold-sans, contrast-mix... |
| 🖌️ Decorations | 10 | accent-bar, neon-lines, gold-trim, diamond-bullets, gradient-bar, circle-accent, sidebar-nav... |
| 📐 Layout Variants | 8 | standard, centered, sidebar-left, sidebar-right, grid-2x2, wide-cards, asymmetric, full-width |

**25 × 20 × 10 × 8 = 40,000 combinations** — plus fuzzy natural language matching!

### Natural Language Style

```bash
ppt-design "investor pitch" --style "warm fintech"        # → ocean-blue + clean-corporate + accent-bar + sidebar-left
ppt-design "product launch" --style "dark cyberpunk"      # → cyber-neon + tech-mono + neon-lines + wide-cards
ppt-design "brand strategy" --style "elegant luxury"      # → golden-luxury + elegant-serif + gold-trim + centered
ppt-design "ESG report" --style "calm nature"             # → sage-calm + humanist-sans + circle-accent + standard
ppt-design "startup pitch" --style "bold startup vibrant"  # → royal-purple + bold-sans + gradient-bar + grid-2x2
```

### Exact Atom Control

```bash
ppt-design "pitch" --palette wine-burgundy --fonts elegant-serif --decoration gold-trim --layout-variant centered
ppt-design "launch" --palette copper-industrial --fonts tech-contrast --decoration no-decoration --layout-variant full-width
```

### Preset Themes (backward compatible)

| Theme | Palette | Fonts | Decoration | Layout |
|-------|---------|-------|------------|--------|
| Professional | midnight-navy | clean-corporate | accent-bar | sidebar-left |
| Dark Tech | cyber-neon | tech-mono | neon-lines | wide-cards |
| Warm Elegant | golden-luxury | serif-editorial | gold-trim | centered |
| Vibrant Startup | neon-gradient | bold-sans | gradient-bar | grid-2x2 |
| Nature Calm | forest-green | humanist-sans | circle-accent | sidebar-left |

---

## 🔗 Relationship with ui-ux-pro-max

**Independent repo + dependency reference** model:

- **Calls** ui-ux-pro-max's search engine and design knowledge base (5100+ CSV entries)
- **Does NOT modify** any ui-ux-pro-max code
- **Adds** PPT-specific code and data
- **Adapter layer** isolates upstream API changes

---

## 📁 Project Structure

```
PPT-Design-Skill/
├── pyproject.toml
├── install.py                        # One-click installer
├── SKILL.md                          # AI skill definition
├── AGENTS.md                         # Project instructions
├── CLAUDE.md                         # Claude Code quick ref
├── docs/showcase/                    # Showcase screenshots
├── src/ppt_pro_max/
│   ├── __init__.py                   # generate_ppt() API
│   ├── cli.py                        # ppt-design CLI
│   ├── adapters/                     # Adapter layer
│   │   ├── ui_ux_adapter.py          # ui-ux-pro-max adapter
│   │   └── slide_search_adapter.py   # slide_search adapter
│   ├── planner/story_planner.py      # Phase 1: Story planning
│   ├── decider/design_decider.py     # Phase 2: Design decisions
│   ├── content/content_generator.py  # Phase 3: Content generation
│   ├── renderer/
│   │   ├── ppt_renderer.py           # Phase 4: PPT rendering
│   │   ├── theme_mapper.py           # Theme mapping + CJK fonts
│   │   ├── theme_composer.py         # 40,000+ style combinations
│   │   ├── layout_registry.py        # 12 master layouts
│   │   ├── chart_builder.py          # Chart builder
│   │   ├── image_fetcher.py          # Image fetching (4 generation + 1 enhancement)
│   │   └── effects.py               # Shadow/glow/gradient
│   └── qa/qa_gates.py               # 5 quality checks
├── data/ppt/                         # PPT-specific data
├── examples/                         # Showcase PPTs
└── tests/                            # 47 tests
```

## License

MIT
