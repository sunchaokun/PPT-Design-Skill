---
name: ppt-design-skill
version: 0.8.0
description: "AI-powered PPT generation — 3 modes: FreeStyle (one-liner), Enterprise (brand compliance), Build (pixel-perfect). Requires ui-ux-pro-max for 192 palettes, 84 styles, 74 font pairs, 161 anti-patterns. 5,500+ chart templates. Engines: Seedream, GPT Image, DALL-E, Wanx, Kimi."
argument-hint: "[topic] [--style style] [--fetch-images] [--project DIR] [--build]"
license: MIT
metadata:
  author: sunchaokun
  category: design
  tags: [ppt, presentation, deck, pitch, slides, python-pptx, enterprise, brand, diagram, proposal]
  requires-skills: [ui-ux-pro-max]
---

# PPT Design Skill

AI-powered PPT generation — design intelligence from ui-ux-pro-max, 5,500+ chart templates, brand compliance, fully editable .pptx.

## Execution Workflow

ALWAYS follow this 5-step workflow. Do NOT skip steps — rework is extremely costly.

### Step 1: Requirements & Design Intelligence

- Understand: topic, audience, language, scenario
- Read any user-provided materials (README, docs, data files)
- **ui-ux-pro-max auto-provides** (no manual selection needed):
  - Product type → industry-specific color palette (192 palettes)
  - Style recommendation → effects, accessibility, dark/light mode
  - Typography pairing → industry-appropriate fonts (74 pairings)
  - Landing pattern → conversion-optimized slide sequence (34 patterns)
  - Anti-patterns → per-industry "what NOT to do" (161 rules)
  - Reasoning rules → conditional design logic
- Design skeleton: total pages, per-page goal, core title
- Present outline, confirm before proceeding

### Step 2: Visual Proposals (Build Mode, 1-3 Options)

**ALWAYS use Build Mode for proposals** — it produces truly differentiated designs (different layouts, visual languages, page structures), unlike FreeStyle which only swaps colors/fonts.

- Generate 1-3 Build Mode PPTs based on user needs:
  - **1 proposal**: User has clear style preference → single Build script with that style
  - **2 proposals**: User wants comparison → two contrasting styles (e.g., business vs creative)
  - **3 proposals**: User wants full exploration → three distinct design strategies
- Each proposal uses a **completely different design strategy** (not just palette swap):
  - Different layout patterns (sidebar vs grid vs centered)
  - Different visual language (data-driven vs narrative vs visual-impact)
  - Different shape vocabulary (rectangles vs rounded vs circles)
  - Different color philosophy (monochrome accent vs dual-tone vs multi-color)
- Use ui-ux-pro-max `get_design_system()` to derive colors, fonts, effects per proposal
- Write a `build_X.py` script per proposal using python-pptx directly
- User picks direction (A/B/C) or requests adjustments

**Build Mode template structure:**
```python
# build_proposal.py — each proposal is a standalone python-pptx script
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# 1. Define color palette (from ui-ux-pro-max design_system)
C = {'bg': '#...', 'primary': '#...', 'accent': '#...', ...}

# 2. Helper functions: _rect, _rrect, _oval, _text, _multiline

# 3. Build each slide with unique layout strategy
#    - Cover: sidebar | split-screen | centered-hero | circles
#    - Problem: numbered-cards | error-log | emoji-cards
#    - Solution: sidebar+cards | code-block | step-circles
#    - Data: comparison-table | dashboard-metrics | before-after
#    - Features: accent-bar-cards | neon-cards | bubble-cards
#    - CTA: dark-full | grid-overlay | warm-circles

prs.save('proposal-X-style.pptx')
```

### Step 3: Detailed Content → content.json

- Write full content per page (see content.json format and Content Rules below)
- **CRITICAL: Respect ui-ux-pro-max anti-patterns** — e.g., legal: no "AI purple/pink gradients"; healthcare: no "Bright neon colors"
- **CRITICAL: ALWAYS use component library for diagram pages** — MANDATORY workflow:
  1. **Query library first**: `query_component_library(type="group", category="<category>")` to check available templates
  2. **Set component fields**: Add `component_type` and `component_category` to every page with 3+ bullets that describe a process/hierarchy/SWOT/timeline/infographic
  3. **Never use plain bullets for diagrams** — hierarchy, process, timeline, SWOT pages MUST use component library
  4. **Auto-inference fallback**: If you forget to set component fields, the pipeline auto-infers them from bullet keywords, but explicit is better
- Present to user for review

### Step 4: Draft Generation & Revision (Build Mode)

- Rename selected proposal: `build_A.py` → `build.py` (canonical working script)
- Expand from 8-page skeleton to full content (15-20+ pages), add images, data, polish
- **Version management** (Build scripts are plain text — git tracks every change):

| Change Type | Strategy | Example |
|-------------|----------|---------|
| Content tweak | Edit `build.py`, git commit | Fix wording, add bullet, adjust number |
| Layout tweak | Edit `build.py`, git commit | Move shape, resize, recolor |
| Style overhaul | Snapshot `build.py` → `build_v1.py`, start fresh | Switch from sidebar to grid layout |
| Output versioning | Auto-increment filename | `output_v1.pptx` → `output_v2.pptx` |

```python
# build.py — version management pattern
import os, glob

def _next_version(basename="output", ext=".pptx"):
    existing = glob.glob(f"{basename}_v*{ext}")
    nums = [int(f.split("_v")[1].split(".")[0]) for f in existing if "_v" in f]
    return f"{basename}_v{max(nums, default=0) + 1}{ext}"

output = _next_version()  # output_v1.pptx, output_v2.pptx, ...
prs.save(output)
print(f"Saved: {output}")
```

- Each revision: edit `build.py` → run → user reviews → commit
- Major style change: `cp build.py build_v1.py` then rewrite `build.py`
- If using FreeStyle/Enterprise → `generate_ppt(query, content_file=..., style=confirmed_style, fetch_images=True, ...)`

### Step 5: Final Delivery

- User confirms final version
- `build.py` is the single source of truth — re-run anytime to regenerate identical output
- Deliver: final `.pptx` + `build.py` (client can edit/rebuild)

## Design Intelligence (powered by ui-ux-pro-max)

All design decisions flow through ui-ux-pro-max. **Do NOT manually pick colors/fonts.**

| Database | Count | What It Provides |
|----------|-------|-------------------|
| Product Types | 192 | Industry detection, style recommendation |
| Color Palettes | 192 | Full semantic palettes (10+ color roles) |
| UI Styles | 84 | Effects, accessibility, dark/light mode |
| Font Pairings | 74 | Heading+body with mood/best-for |
| Landing Patterns | 34 | Section order, CTA placement, conversion strategy |
| Anti-Patterns | 161 | Per-industry "what NOT to do" |
| Reasoning Rules | 161 | Conditional design logic |
| Chart Types | 25 | Data type → optimal visualization |

**How it works (automatic):**
1. User says "fintech investor pitch"
2. StoryPlanner → "Fintech/Crypto", style rec "Glassmorphism + Dark Mode"
3. DesignDecider → palette (#F59E0B primary, #8B5CF6 accent), anti-patterns ("Playful design + Unclear fees")
4. ThemeComposer → searches colors/typography/styles CSV for best match
5. PrecisionRenderer receives full design context

## Component Library (5,500+ Professional Chart Templates)

**MANDATORY: Use component library for ALL diagram pages.** The pipeline auto-infers component fields from bullets, but you should explicitly set them for best results. These are real professional PPT chart templates, not hand-drawn.

### MANDATORY Workflow (before writing content.json)

1. **Query the library** to check available categories and node_counts:
   ```python
   from ppt_pro_max import query_component_library
   catalog = query_component_library()  # See all categories
   process = query_component_library(type="group", category="process")  # Search specific
   hierarchy = query_component_library(type="group", category="hierarchy", node_count=5)  # Precise
   ```
2. **Set component fields** in content.json for every diagram page
3. **Pipeline auto-inference** as fallback — but explicit is always better

| Category | Count | Use When |
|----------|-------|----------|
| infographic | 4,101 | Data dashboards, KPI grids, statistics |
| process | 672 | Workflows, pipelines, step-by-step |
| hierarchy | 548 | Org charts, reporting structures |
| chart | 132 | Bar/pie/line chart layouts |
| timeline | 42 | Milestones, roadmaps |
| swot | 39 | Strategic analysis, 4-quadrant |

### Auto-Inference (write natural bullets, system matches template)

| Bullet Keywords | Auto-Category | Example |
|----------------|--------------|---------|
| 优势/劣势/机会/威胁, strength/weakness/opportunity/threat | `swot` | ["技术优势", "市场劣势", "增长机会", "潜在威胁"] |
| 组织/架构/层级/CEO/CTO, org/team/structure | `hierarchy` | ["CEO", "CTO", "VP Engineering"] |
| 流程/步骤/阶段, step/phase/process | `process` | ["需求分析", "方案设计", "开发实现", "测试上线"] |
| 时间线/里程碑, timeline/milestone | `timeline` | ["Q1 发布", "Q2 拓展", "Q3 盈利"] |
| KPI/指标/数据, metric/dashboard | `infographic` | ["收入 ¥1.2亿", "增长 45%"] |

### Selection Priority

1. **Complex diagrams** (hierarchy, process, timeline, swot) → ALWAYS component library first — NEVER use built-in bullets for hierarchy
2. **Data-heavy pages** (infographic, chart) → Component library for visual richness
3. **Simple content** (2-5 bullets, features cards, code) → Built-in PrecisionRenderer layouts

### Explicit Control in content.json

```json
{
  "goal": "content",
  "title": "项目流程",
  "bullets": ["需求分析", "方案设计", "开发实现", "测试上线"],
  "component_type": "group",
  "component_category": "process",
  "component_fit": "contain",
  "component_bounds": [0.9, 1.6, 11.5, 5.0]
}
```

- `component_fit`: `contain` (default, preserve aspect), `width`, `height`, `stretch`
- `component_bounds`: `[left, top, width, height]` in inches — explicit position/size override
- Without `component_bounds`: system auto-calculates content area based on title/image presence

Matching: exact node_count → closest node_count → fallback to DiagramEngine/bullets.

### Aspect Ratio & Fit Mode

Components auto-preserve original aspect ratio. Control with `component_fit`:

| fit | Effect | Use when |
|-----|--------|----------|
| `contain` (default) | Fit inside area, may have padding | General, safest |
| `width` | Fill width, height proportional | Horizontal process, timeline |
| `height` | Fill height, width proportional | Vertical org charts |
| `stretch` | Ignore aspect, fill entire area | Only when aspect ≈ area |

```json
{
  "component_type": "group",
  "component_category": "process",
  "component_fit": "width",
  "bullets": ["Step 1", "Step 2", "Step 3"]
}
```

### Coordinate Transform (lossless)

Components store original XML + original bounds EMU. Injection applies uniform scale+translate to ALL xfrm elements (off/ext/chOff/chExt), preserving `ext/chExt` ratio. PowerPoint's mapping `screen = (child - chOff) * (ext/chExt) + off` produces correct positions at any target size.

### Query Before Writing

```python
from ppt_pro_max import query_component_library
catalog = query_component_library()  # Full catalog
results = query_component_library(type="group", category="process")  # Search
results = query_component_library(type="group", category="hierarchy", node_count=5)  # Precise
```

**AI should query the library before writing content.json** to match available node_counts.

## Content Design Rules

| Rule | Why | Example |
|------|-----|---------|
| Respect anti-patterns | Industry-specific violations ruin credibility | Legal: no neon gradients; Healthcare: no motion-heavy |
| Use component library for diagrams | 5,500+ templates > bullet lists | SWOT → `component_category: "swot"` |
| features: first card longest | First card gets featured treatment | Card 1: "智能推理引擎 — 自动选择最优框架，量化压缩70%" |
| 6+ bullets on data pages | Triggers two-column layout | 6 concise data points instead of 3 long ones |
| Include a code page for tech | Adds technical credibility | `{"code": {"language": "python", "source": "..."}}` |
| Insert section dividers | Visual rhythm between topic shifts | `{"goal": "section", "title": "..."}` |
| hook: short subtitle; cta: long | Different hero compositions | hook: "5分钟取代5周" vs cta: long subtitle |
| Vary bullet density | Natural feel, not templated | Mix 3-bullet and 6+ bullet pages |
| Use concrete numbers | Specific > vague | "GPU成本年增3倍" not "成本持续增长" |

## When to Activate

- Create/generate/design a PPT/presentation/deck/slide deck
- Pitch deck, product demo, sales presentation, investor deck
- Convert content/outline into PowerPoint
- Brand-compliant presentations with template + version control
- Page-level CRUD on existing PPT (add/delete/swap/move pages)
- Diagrams in PPT (flowchart, funnel, timeline, SWOT, etc.)

## Three-Mode Architecture

| Mode | User | Scenario | One-liner |
|------|------|----------|-----------|
| **FreeStyle** | Everyone | One sentence → one PPT | "给我做个AI路演" → 30s done |
| **Enterprise** | Corporate teams | VI compliance, brand consistency, version control | brand.json + template.pptx lock style |
| **Build** | Designers / AI | Pixel-perfect, proposal comparison, version management | python-pptx script = source of truth |

### FreeStyle — 傻瓜模式

One sentence, one PPT. Simplest, fastest way to create a presentation.

```bash
python -m ppt_pro_max "AI startup investor pitch"
python -m ppt_pro_max "fintech pitch" --style "warm fintech"
python -m ppt_pro_max "AI pitch" --fetch-images --llm-provider seedream --llm-api-key $ARK_API_KEY
python -m ppt_pro_max "pitch" --palette wine-burgundy --fonts elegant-serif
```

- Auto: StoryPlanner → DesignDecider → ContentGenerator → PPTRenderer
- 35 mood keywords, 40,000+ style combos
- Good for: quick drafts, personal use, non-designers

### Enterprise — 企业模式

Brand compliance pipeline. Reads brand.json + template.pptx + content.json, enforces VI rules, manages versions.

```bash
python -m ppt_pro_max "路演" --project ./my-project --density 6
python -m ppt_pro_max "" --project ./my-project --review
python -m ppt_pro_max "" --project ./my-project --pages "-3,2<>5,10>3,+6"
python -m ppt_pro_max "" --project ./my-project --history
```

- Brand lock: colors, fonts, logo, footer, watermark from brand.json
- Template inheritance: slide masters, layouts from template.pptx
- Content priority: content.json > README.md > StoryPlanner
- Version control: output/v1/, output/v2/ auto-incremented
- Page operations: add, delete, swap, move pages
- Good for: corporate teams, agencies, brand-managed organizations

### Build — 设计师模式

Write python-pptx scripts directly. Full control over every shape, color, position. Each script is a standalone, version-manageable design.

```python
# build.py — full control over every shape, color, position
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

C = {'bg': '#0A0E17', 'primary': '#00F0FF', 'accent': '#FF2E97', ...}
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
# ... build slides with _rect, _rrect, _oval, _text helpers
prs.save('output.pptx')
```

- **Proposal phase**: 1-3 `build_X.py` scripts, each with completely different design strategy
- **Revision phase**: single `build.py`, git commit per change, auto-increment output versions
- **Delivery**: `.pptx` + `build.py` (client can edit/rebuild)
- Good for: proposal comparison, pixel-perfect output, designer-quality work, AI-assisted design

### Python API

```python
from ppt_pro_max import generate_ppt

result = generate_ppt("AI startup pitch", style="dark cyberpunk", fetch_images=True)
result = generate_ppt("路演", project="./my-project", density=6)
print(f"Generated: {result['output_path']}, {result['page_count']} pages")
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
      "component_category": "process",
      "component_fit": "contain"
    }
  ]
}
```

**Goal → Rendering:**

| Goal | Rendering |
|------|-----------|
| `hook`/`cta` | Hero: full-bleed image + gradient overlay + title |
| `section` | Divider: oversized number + title + gradient line |
| `problem`/`solution`/`content` | Title + gradient line + bullets + optional image; 6+ bullets → two-column |
| `features` | Title + cards row (first card featured) |
| `data`/`overview` | Title + bullets or diagram |
| `code` | Title + code block (dark bg) + language badge |
| `exercise` | Title + badge + instructions + numbered steps |

## 10 Diagram Types

| Type | Data Format |
|------|-------------|
| Flowchart | nodes + connectors |
| Funnel | stages (items) |
| Timeline | events (items) |
| SWOT | strengths/weaknesses/opportunities/threats |
| Matrix | rows + columns |
| Cycle | stages (items) |
| Table | headers + rows |
| Hierarchy | nodes with parent |
| Pyramid | levels (items) |
| Venn | sets with labels |

## Image Engines

| Engine | Provider | Env Key |
|--------|----------|---------|
| Seedream | Volcengine | ARK_API_KEY |
| GPT Image | OpenAI | OPENAI_API_KEY |
| DALL-E 3 | OpenAI | OPENAI_API_KEY |
| Wanx | Alibaba | DASHSCOPE_API_KEY |
| Kimi K2.6 | Moonshot | MOONSHOT_API_KEY |

Modes: `placeholder` (default), `search` (Unsplash/Pexels), `generate` (AI), `enhance` (Kimi). Cache-first — same image never generated twice.

## Enterprise Project Structure

```
my-project/
├── template.pptx    # Brand template (optional)
├── brand.json       # Brand spec (colors, fonts, logo, footer, watermark)
├── content.json     # Page content
├── logo.png         # Company logo
├── images/          # Local images
└── output/
    ├── v1/          # Version 1
    └── v2/          # Version 2 (auto-incremented)
```

## brand.json Format

```json
{
  "colors": {
    "primary": "#1E3A5F", "on-primary": "#FFFFFF",
    "accent": "#E8A838", "background": "#0A1E3D",
    "foreground": "#F0F4F8", "muted": "#1A2E4A"
  },
  "logo": {"position": "top_right", "width_inches": 1.0, "skip_cover": true},
  "footer": {"text": "Company Name", "show_page_number": true},
  "spacing": {"strip_style": "auto", "section_dividers": true}
}
```

## CLI Reference

```
ppt-design "query" [options]

  --style TEXT          Natural language style
  --palette/--fonts/--decoration/--layout-variant  Atom control
  --mood TEXT           Mood hint
  --style-seed INT      Reproducible seed
  --slides N            Override slide count
  --content FILE        JSON content file
  --variance/motion/density 1-10  Design dials
  --fetch-images        Enable image search/generation
  --llm-provider PROV   seedream|gpt-image|dalle|wanx|kimi
  --project DIR         Enterprise mode
  --business-mode MODE  pitch|education|training|report
  --review              Preview plan before generating
  --pages OPS           Page operations (-3,2<>5,10>3,+6)
  --beautify PPTX       Beautify existing PPT
  --beautify-mode MODE  light|full
  --component-library PATH  DB path (auto-detected)
  -o, --output PATH     Output .pptx path
```

## Page Revision Syntax

`--pages "3,5 +6 -8 3>5 3<>7"` — N=keep, +N=insert, -N=delete, N>M=move, N<>M=swap. 1-based, refers to ORIGINAL document.

## Component Library Management

```bash
# Build from PPT素材
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材目录" --min-node-count 3

# Multiple directories (dedup automatic)
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材A"
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材B"
```

## Dependencies

- python-pptx >= 1.0.2 (required)
- Pillow >= 10.0 (required)
- lxml >= 4.9.0 (required)
- **ui-ux-pro-max** (required — NO fallback)
  - Install: `npx ui-ux-pro-max-cli init --ai opencode`
  - Or set `UX_PRO_MAX_DIR` env var
  - Missing → `UiUxProMaxNotFoundError`, system will NOT start
  - All design decisions flow through ui-ux-pro-max — auto-selected per industry
- python-dotenv >= 1.0 (optional)
- httpx >= 0.24.0 (optional, image search)
- openai >= 1.0.0 (optional, AI images)
