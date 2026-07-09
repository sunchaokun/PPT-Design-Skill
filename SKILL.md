---
name: ppt-design-skill
version: 0.1.0
description: AI-powered PPT generation — narrative-driven, design-intelligent, fully editable .pptx
author: sunchaokun
category: design
tags: [ppt, presentation, design, narrative, python-pptx]
strict: false
---

# PPT Design Skill

Professional PPT generation skill for AI coding assistants. Generates complete .pptx presentations from a single sentence.

## Activation

Activate when the user asks to:
- Create, generate, or design a PPT/presentation/deck/slide deck
- Make a pitch deck, product demo, sales presentation, or investor deck
- Convert content/outline into a PowerPoint file

## Workflow

### Step 1: Parse Intent

Extract from user query:
- **Product type**: What kind of product/project? (SaaS, AI, hardware, service...)
- **Audience**: Who is the presentation for? (investors, customers, team, conference...)
- **Goal**: What should the audience do? (invest, buy, understand, agree...)
- **Tone**: Formal / casual / bold / elegant?

### Step 2: Narrative Planning

Use `story_planner.py` to:
1. Select a presentation strategy (YC Seed Deck, Product Demo, Sales Pitch, etc.)
2. Generate page structure with goals and emotions
3. Create emotion arc (curiosity → frustration → hope → confidence → urgency)

### Step 3: Design Decisions

Use `design_decider.py` to:
1. Generate design system (colors, fonts, spacing) via ui-ux-pro-max adapter
2. Per-page layout, typography, color treatment decisions
3. Map design system to PPT theme

### Step 4: Content Generation

Use `content_generator.py` to:
1. Apply copy formulas (PAS, FAB, AIDA, Social Proof...)
2. Generate placeholder content with structure
3. Configure chart data if applicable

### Step 5: PPT Rendering

Use `ppt_renderer.py` to:
1. Create .pptx with theme and master layouts
2. Fill placeholders with content
3. Add charts, transitions, and effects
4. Run QA gates

### Step 6: Output

- Save .pptx file
- Report: page count, strategy used, theme applied
- Remind user to replace placeholder content

## CLI Reference

```bash
ppt-design "query" [options]

Options:
  --strategy NAME      Override presentation strategy
  --theme NAME         Theme preset (professional, dark-tech, warm-elegant, vibrant-startup, nature-calm)
  --slides N           Override slide count
  --content FILE       JSON file with real content
  --variance 1-10      Design variance (centered → bold)
  --motion 1-10        Animation intensity (subtle → dramatic)
  --density 1-10       Content density (spacious → dense)
  --fetch-images       Auto-download images from Unsplash/Pexels
  --persist            Persist design system as MASTER.md
  --dry-run            Output design decisions only (no .pptx)
  -o, --output PATH    Output file path
```

## Design Principles

1. **Narrative first** — PPT is storytelling, not decoration
2. **Theme consistency** — All pages share one theme (colors + fonts)
3. **Master layouts** — 12 predefined layouts, not free-form positioning
4. **Framework + placeholders** — Professional structure, user fills content
5. **Context-aware** — Layout/color/typography decisions adapt to each page's goal and emotion
6. **Fully editable** — Output is native .pptx, not screenshots

## Dependencies

- python-pptx >= 1.0.2 (required)
- ui-ux-pro-max >= 1.0.0 (optional, for search + design knowledge)
