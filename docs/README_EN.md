# PPT Design Skill

PPT Design Skill is a brief-first presentation design workflow powered
by the published [`pptx-designer`](https://pypi.org/project/pptx-designer/)
Python library.

Current release: `1.0`

The library generates the editable PPTX. The skill is responsible for the
design process and quality gate:

```text
brief -> page structure -> visual direction -> build -> PPTX -> PDF -> PNG ->
LLM visual review -> revision -> user confirmation -> delivery
```

## Install

Clone the repository first and run the following commands from its root:

```powershell
# Clone the skill repository
git clone https://github.com/sunchaokun/PPT-Design-Skill.git
cd PPT-Design-Skill

# Install the published pptx-designer runtime
python install.py
python skill/scripts/check_runtime.py
```

Install the skill bundle for supported coding assistants:

```powershell
python installer/install.py --platform all --force
python installer/install.py --platform deepseek-harness --project --force
```

For a single assistant, replace `all` with `opencode`, `codex`, `claude`, or
another supported platform key. Restart the assistant after installation.

The installer supports Claude Code, Codex, DeepSeek Harness, OpenCode, Cursor,
Windsurf, Roo Code, Gemini CLI, Trae, Continue, Droid, KiloCode, Augment, and
GitHub Copilot. See [installer/README.md](../installer/README.md).

## Three modes

| Mode | Use case | Implementation |
|---|---|---|
| Build Mode | Delivery-grade blank-canvas composition | Python + public `pptx_designer.tools.*` |
| FreeStyle Mode | Fast exploration or goal-driven generation | `generate_ppt(query=...)` or `generate_ppt(content=...)` |
| VI Build Mode | Existing template and enterprise brand compliance | Template + `extract_design_dna()` + controlled new pages |

### Build Mode

Use Build Mode when exact coordinates, custom composition, advanced diagrams,
or a stable design-token system are required.

### FreeStyle: `generate_ppt()`

FreeStyle is the library's goal-based generation path. A query is convenient
for a fast topic-driven draft; a structured `content` dictionary gives the LLM
more control over page goals and copy. Both use `generate_ppt()` and neither
provides pixel-level element placement.

### VI Build Mode

Use VI Build when the user supplies a corporate template or requires brand
compliance. Analyze the template, extract design DNA, preserve framework pages,
add content pages using the template's visual system, and review the complete
deck through PPTX -> PDF -> PNG. Read
[template-brand.md](../skill/references/template-brand.md).

Complex PowerPoint masters, SmartArt, animations, and private OOXML behavior
may require a controlled approximation and must not be promised as pixel-perfect.

### Build Mode implementation

Build Mode writes ordinary, reviewable Python using public
`pptx_designer` helpers. Use it for delivery-grade work, custom composition,
complex diagrams, brand systems, and exact placement. See the real design
cases in [examples/README.md](../examples/README.md).

## Quality gate

Do not claim completion because the Python script ran or a PPTX file exists.
Always export the confirmed PPTX -> PDF -> PNG path, inspect every rendered
slide, revise material visual issues, and obtain user confirmation.

The visual review is performed by the LLM reading the PNGs. The skill does not add a
separate visual scoring service.

Before generation, convert the user's brief into a small acceptance contract.
After rendering, compare every `MUST` condition with visible evidence in the
PNG and record `PASS`, `NEEDS_REVISION`, or `BLOCKED`. A general statement that
the deck “looks good” is not sufficient.

## Documentation

- [Chinese usage guide](usage-guide.md)
- [Skill workflow](../skill/SKILL.md)
- [Public API contract](../skill/references/public-api.md)
- [QA and delivery](../skill/references/qa-and-delivery.md)
- [Examples](../examples/README.md)

## Real design cases

The repository includes three complete editorial PPTX artifacts from the maintained
`pptx-designer` examples:

- [Luxury Fragrance Lookbook](../examples/output/luxury_fragrance_lookbook.pptx)
- [Couture Editorial Deck](../examples/output/couture_editorial_deck.pptx)
- [Architecture Vision Book](../examples/output/architecture_vision_book.pptx)

These are visual reference cases, not disposable smoke-test decks. They show
how the skill should use distinct page structures, atmosphere imagery, native
editable text and shapes, and a coherent editorial narrative.

### Visual preview

[![Three PPT Design Skill case studies](assets/cases/contact-sheet.png)](../examples/README.md)

The preview is rendered from the actual case PPTX files through the confirmed
PPTX -> PDF -> PNG path. Open the [Luxury Fragrance](assets/cases/luxury-fragrance-slide01.png),
[Couture Editorial](assets/cases/couture-editorial-slide01.png), or
[Architecture Vision](assets/cases/architecture-vision-slide01.png) representative page.

## What the skill does

The skill uses a mature presentation-design framework with `pptx-designer` as
its implementation engine.
It covers:

- audience and scenario analysis;
- page-level narrative planning;
- domain-specific visual paradigms;
- palette, typography, spacing, density, and image direction;
- FreeStyle, Build Mode, and VI Build Mode routing;
- reproducible Python generation;
- basic PPTX structural inspection;
- confirmed PPTX -> PDF -> PNG rendering;
- direct LLM review of every rendered PNG;
- targeted revision and user confirmation.

The skill is not a prompt-to-image service or a replacement for PowerPoint's
rendering engine. The Python library produces the editable file;
the skill is the design and quality-control layer around it.

## Design principles

The following principles remain mandatory:

1. Design for the audience and the communication goal.
2. Build a coherent visual system across the deck.
3. Prefer restraint, hierarchy, and meaningful variation over decoration.
4. Detect the domain before choosing a business-slide pattern.
5. Use real data and label assumptions.
6. Keep important information as native editable objects.
7. Do not call a deck complete until the rendered PNGs have been reviewed.

The skill explicitly rejects repeated default card grids, tiny unreadable text,
stretched images, random theme changes, fake precision, and full-page screenshots
used as a substitute for editable content.

## Mode selection in practice

FreeStyle has two input forms, but they use the same `generate_ppt()` library
pipeline:

```python
from pptx_designer import generate_ppt

# Topic-driven draft
generate_ppt("AI startup pitch", style="professional", output="output/pitch.pptx")

# Content-driven draft
generate_ppt(
    content={
        "title": "Business review",
        "pages": [
            {"goal": "hook", "title": "The signal is clear"},
            {"goal": "data", "title": "Key metrics", "bullets": ["Revenue: $12M"]},
        ],
    },
    style="professional",
    output="output/review.pptx",
)
```

Use Build Mode when exact coordinates, custom diagrams, a stable design-token
system, or delivery-grade composition is required. Build Mode scripts are
ordinary Python and should be reviewed and versioned like application code.

Use VI Build when a template or enterprise brand system must be preserved. The
template is analyzed, framework pages are protected, new pages inherit the
brand token system, and the complete result is checked through PNG rendering.

## Mandatory visual gate

After generation, run:

```powershell
powershell -ExecutionPolicy Bypass -File skill/scripts/render_pptx.ps1 `
  -InFile output/deck.pptx `
  -OutDir output/deck-rendered
```

Then inspect every PNG for hierarchy, readability, spacing, overflow, crop,
chart legibility, page rhythm, domain fit, and consistency. If the result is
not acceptable, revise the source and rerun the complete loop. A successful
Python process or valid PPTX package is not a visual acceptance criterion.

## Installation matrix

| Assistant | Global skill root | Project skill root |
|---|---|---|
| Claude Code | `~/.claude/skills` | `.claude/skills` |
| Codex | `~/.agents/skills` | `.codex/skills` |
| OpenCode | `~/.config/opencode/skills` | `.opencode/skills` |
| DeepSeek Harness | `~/.dsh/skills` | `.dsh/skills` |

Install with `installer/install.py`. The installer also supports Cursor,
Windsurf, Roo Code, Gemini CLI, Trae, Continue, Droid, KiloCode, Augment, and
GitHub Copilot. Python dependencies are installed separately through
`pptx-designer`; LibreOffice and Poppler are optional system dependencies for
the headless renderer.
