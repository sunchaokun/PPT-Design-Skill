---
name: ppt-design-skill
version: 0.2.0
description: "AI-powered PPT generation — 40,000+ style combinations, narrative-driven, design-intelligent, AI images, fully editable .pptx. Actions: create, generate, design, build PPT/presentation/deck. Themes: professional, dark-tech, warm-elegant, vibrant-startup, nature-calm + infinite combos. Engines: Seedream, GPT Image, DALL-E, Wanx, Kimi."
argument-hint: "[topic] [--style style-description] [--theme preset] [--fetch-images]"
license: MIT
metadata:
  author: sunchaokun
  category: design
  tags: [ppt, presentation, deck, pitch, slides, python-pptx, design, narrative]
  stacks: [python, pptx]
---

# PPT Design Skill

Professional PPT generation skill for AI coding assistants. Generates complete, fully editable .pptx presentations from a single sentence — with 40,000+ style combinations, narrative-driven structure, and AI-powered image generation.

## When to Activate

Activate when the user asks to:
- Create, generate, or design a **PPT/presentation/deck/slide deck**
- Make a **pitch deck, product demo, sales presentation, or investor deck**
- Convert **content/outline** into a PowerPoint file
- Design **slides** for a meeting, conference, or workshop
- Build a **fundraising deck, brand deck, or strategy presentation**

### Skip

- Pure backend logic or API development
- Database design or infrastructure work
- Non-visual scripts or automation tasks

**Decision criteria**: If the user wants a **visual presentation file (.pptx)**, activate this skill.

## Quick Start

```bash
# One-liner PPT generation
python -m ppt_pro_max "AI startup investor pitch"

# With natural language style (40,000+ combinations)
python -m ppt_pro_max "fintech pitch" --style "warm fintech"
python -m ppt_pro_max "product launch" --style "dark cyberpunk"
python -m ppt_pro_max "brand strategy" --style "elegant luxury minimal"

# With AI-generated images
python -m ppt_pro_max "AI pitch" --fetch-images --llm-provider seedream --llm-api-key YOUR_KEY

# With exact design atoms
python -m ppt_pro_max "pitch" --palette wine-burgundy --fonts elegant-serif --decoration gold-trim --layout-variant centered
```

## Python API

```python
from ppt_pro_max import generate_ppt

# Minimal
result = generate_ppt("AI startup investor pitch")

# Natural language style
result = generate_ppt("fintech pitch", style="warm fintech")

# Full control
result = generate_ppt(
    query="AI startup investor pitch",
    style="dark cyberpunk tech",
    strategy="YC Seed Deck",
    slides=12,
    fetch_images=True,
    llm_provider="seedream",
    llm_api_key="ark-xxx",
    output="my-pitch.pptx",
)

print(f"Generated: {result['output_path']}, {result['page_count']} pages")
print(f"Style atoms: {result.get('theme_atoms', {})}")
```

## Workflow

### Step 1: Parse Intent

Extract from user query:
- **Product type**: SaaS, AI, hardware, service, brand, sustainability...
- **Audience**: Investors, customers, team, conference...
- **Goal**: Invest, buy, understand, agree...
- **Style**: Use `--style` for natural language or exact atoms

### Step 2: Narrative Planning (Phase 1)

`story_planner.py` → strategy + page structure + emotion arc:
- **YC Seed Deck**: Hook → Problem → Solution → Traction → Market → Team → Financial → CTA
- **Product Demo**: Hook → Vision → Problem → Demo → Features → Proof → CTA
- **Sales Pitch**: Hook → Pain → Agitation → Solution → Proof → Offer → CTA

### Step 3: Design Decisions (Phase 2)

`design_decider.py` → per-page layout/color/typography/chart/transition:
- Layout: title-slide, content-with-title, three-column-cards, four-metrics, big-number, quote, chart-focus, image-plus-text, cta-closing...
- Theme: 40,000+ combinations via composable design atoms

### Step 4: Content Generation (Phase 3)

`content_generator.py` → copy formulas + image keywords:
- 9 copy formulas: PAS, FAB, AIDA, Social Proof, Cost of Inaction, Proof Stack...
- 18 goal-specific generators with real persuasive content
- Auto-detect context (deep tech, fintech, sustainability, sales, product...)

### Step 5: PPT Rendering (Phase 4)

`ppt_renderer.py` → python-pptx direct output:
- 12 master layouts with precise inch coordinates (13.333" × 7.5" 16:9)
- Cover-fit image insertion (Pillow pre-crop, no distortion)
- Dark overlay for hero slides
- CJK font fallback (Microsoft YaHei / STSong)
- QA gates: 5 automated quality checks

### Step 6: Output

- Save .pptx file
- Report: page count, strategy, theme, style atoms used
- Remind user to review and customize content

## Design Atoms — 40,000+ Combinations

### Color Palettes (25)

ocean-blue, midnight-navy, cyber-neon, neon-gradient, golden-luxury, rose-gold, forest-green, sage-calm, sunset-warm, terracotta, cherry-red, royal-purple, arctic-frost, slate-minimal, charcoal-bold, coral-energy, teal-fresh, indigo-deep, copper-industrial, monochrome, monochrome-dark, lavender-dream, mint-fresh, wine-burgundy, sky-bright

### Font Pairs (20)

modern-sans, geometric-sans, bold-sans, clean-corporate, serif-editorial, elegant-serif, literary-serif, tech-mono, mono-clean, swiss-style, humanist-sans, friendly-round, sharp-modern, classic-formal, contrast-mix, tech-contrast, warm-mix, startup-mix, minimal-mix, editorial-mix

### Decoration Styles (10)

accent-bar, neon-lines, gold-trim, minimal-dots, diamond-bullets, gradient-bar, circle-accent, sidebar-nav, no-decoration, full-bleed-overlay

### Layout Variants (8)

standard, centered, sidebar-left, sidebar-right, wide-cards, grid-2x2, asymmetric, full-width

### Natural Language Style Matching

| Style Description | Auto-Selected Atoms |
|---|---|
| warm fintech | ocean-blue + clean-corporate + accent-bar + sidebar-left |
| dark cyberpunk | cyber-neon + tech-mono + neon-lines + wide-cards |
| elegant luxury | golden-luxury + elegant-serif + gold-trim + centered |
| calm nature | sage-calm + humanist-sans + circle-accent + standard |
| bold startup | royal-purple + bold-sans + gradient-bar + grid-2x2 |
| minimal corporate | slate-minimal + modern-sans + no-decoration + standard |
| industrial tech | copper-industrial + tech-contrast + accent-bar + full-width |

## Image Engines (5)

| Engine | Provider | CLI Flag | Notes |
|--------|----------|----------|-------|
| Seedream Pro | ByteDance Volcengine | `--llm-provider seedream` | Recommended, high quality |
| GPT Image | OpenAI | `--llm-provider gpt-image` | GPT Image 2/1.5/1 |
| DALL-E 3 | OpenAI | `--llm-provider dalle` | Classic AI generation |
| Wanx | Alibaba DashScope | `--llm-provider wanx` | Chinese market |
| Kimi K2.6 | Moonshot | `--llm-provider kimi` | Keyword enhance + search |

All engines use **cache-first** — same image never generated twice.

## CLI Reference

```
ppt-design "query" [options]

Style:
  --style TEXT          Natural language style ("warm fintech", "dark cyberpunk")
  --palette NAME        Color palette (25 options)
  --fonts NAME          Font pair (20 options)
  --decoration NAME     Decoration style (10 options)
  --layout-variant NAME Layout variant (8 options)
  --mood TEXT           Mood hint (professional, tech, warm, elegant, vibrant, nature)
  --style-seed INT      Reproducible random seed

Structure:
  --strategy NAME       Override strategy (YC Seed Deck, Product Demo, Sales Pitch)
  --theme NAME          Theme preset (backward compatible)
  --slides N            Override slide count
  --content FILE        JSON file with real content
  --variance 1-10       Design variance (centered → bold)
  --motion 1-10         Animation intensity (subtle → dramatic)
  --density 1-10        Content density (spacious → dense)

Images:
  --fetch-images        Auto-fetch images
  --image-mode MODE     placeholder/search/generate/enhance/auto
  --llm-provider PROV   Image engine (seedream/gpt-image/dalle/wanx/kimi)
  --llm-api-key KEY     API key for image engine
  --llm-model MODEL     Model name override

Output:
  --persist             Persist design system as MASTER.md
  --dry-run             Design decisions only (no .pptx)
  -o, --output PATH     Output file path
```

## Content JSON Format

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

## Design Principles

1. **Narrative first** — PPT is storytelling, not decoration
2. **Infinite styles** — 40,000+ combinations via composable design atoms
3. **Theme consistency** — All pages share one theme (colors + fonts + decoration)
4. **Master layouts** — 12 predefined layouts with precise coordinates
5. **Context-aware** — Layout/color/typography adapt to each page's goal and emotion
6. **Fully editable** — Output is native .pptx, not screenshots
7. **Cache-first images** — Same image never generated twice, saving API costs

## Dependencies

- python-pptx >= 1.0.2 (required)
- Pillow >= 10.0 (required, for cover-fit image cropping)
- ui-ux-pro-max >= 1.0.0 (optional, for search + design knowledge)
