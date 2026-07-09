---
name: ppt-design-skill
version: 0.2.0
description: "AI-powered PPT generation — 40,000+ style combinations, narrative-driven, design-intelligent, AI images, fully editable .pptx. Actions: create, generate, design, build PPT/presentation/deck. Themes: professional, dark-tech, warm-elegant, vibrant-startup, nature-calm + infinite combos via --style. Engines: Seedream, GPT Image, DALL-E, Wanx, Kimi."
argument-hint: "[topic] [--style style-description] [--fetch-images]"
license: MIT
metadata:
  author: sunchaokun
  category: design
  tags: [ppt, presentation, deck, pitch, slides, python-pptx]
---

# PPT Design Skill

Professional PPT generation for AI coding assistants. 40,000+ style combos, narrative-driven, fully editable .pptx.

## When to Activate

- User asks to create/generate/design a **PPT/presentation/deck/slide deck**
- User wants a **pitch deck, product demo, sales presentation, investor deck**
- User wants to **convert content/outline into PowerPoint**

## Quick Start

```bash
# One-liner
python -m ppt_pro_max "AI startup investor pitch"

# Natural language style (40K+ combos)
python -m ppt_pro_max "fintech pitch" --style "warm fintech"
python -m ppt_pro_max "product launch" --style "dark cyberpunk"
python -m ppt_pro_max "brand strategy" --style "elegant luxury"

# AI images (Seedream Pro recommended)
python -m ppt_pro_max "AI pitch" --fetch-images --llm-provider seedream --llm-api-key YOUR_KEY

# Exact atoms
python -m ppt_pro_max "pitch" --palette wine-burgundy --fonts elegant-serif --decoration gold-trim --layout-variant centered
```

## Python API

```python
from ppt_pro_max import generate_ppt

result = generate_ppt("AI startup pitch", style="dark cyberpunk", fetch_images=True)
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
| Color Palettes | 25 | ocean-blue, cyber-neon, golden-luxury, wine-burgundy... |
| Font Pairs | 20 | modern-sans, serif-editorial, tech-mono, contrast-mix... |
| Decorations | 10 | accent-bar, neon-lines, gold-trim, diamond-bullets... |
| Layout Variants | 8 | sidebar-left, centered, grid-2x2, asymmetric... |

Natural language: `--style "warm fintech"` auto-selects matching atoms.

## Image Engines

Seedream Pro (recommended) · GPT Image · DALL-E 3 · Wanx · Kimi K2.6

All engines use **cache-first** — same image never generated twice.

## CLI

```
ppt-design "query" [--style TEXT] [--palette NAME] [--fonts NAME] [--decoration NAME]
  [--layout-variant NAME] [--mood TEXT] [--style-seed INT]
  [--strategy NAME] [--theme NAME] [--slides N] [--content FILE]
  [--fetch-images] [--llm-provider PROV] [--llm-api-key KEY]
  [--persist] [--dry-run] [-o PATH]
```

## Dependencies

- python-pptx >= 1.0.2 (required)
- Pillow >= 10.0 (required)
- ui-ux-pro-max >= 1.0.0 (optional)
