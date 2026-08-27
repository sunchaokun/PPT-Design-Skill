# PPT Design Skill

PPT Design Skill is a brief-first presentation design workflow powered
by the published [`pptx-designer`](https://pypi.org/project/pptx-designer/)
Python library.

Current release: `1.1`

The library generates the editable PPTX. The skill is responsible for the
design process and quality gate:

```text
brief -> page structure -> visual direction -> build -> PPTX -> PDF -> PNG ->
LLM visual review -> revision -> user confirmation -> delivery
```

## Choose the right mode

### For delivery-grade work, choose **Build Mode**

Build Mode is the recommended default when the deck will be used with clients,
executives, investors, or in a formal presentation. It provides the highest
control over page structure, visual direction, layout, and revision quality.

| Mode | Best for | Layout control | Speed | Recommendation |
|---|---|---:|---:|---:|
| **Build Mode** ⭐ | Client delivery, strategy, proposals, editorial, formal decks | Highest | Medium | **Default** |
| **FreeStyle Mode** | Fast exploration, quick drafts, known page goals | Medium: library-driven | Fastest | Explore first |
| **VI Build Mode** | Existing templates, masters, and brand systems | Template-controlled | Medium | Template first |

If visual quality and delivery confidence matter most, use **Build Mode**. Use
FreeStyle when speed matters more than exact composition; use VI Build when an
existing enterprise template must remain the visual source of truth.

`generate_ppt(query=...)` and `generate_ppt(content=...)` are two input forms of
the same FreeStyle mode, not separate rendering engines. Every mode still goes
through the PPTX -> PDF -> PNG visual review gate for delivery work.

## Install

Clone the repository first, then run the installer from its root. The installer
automatically installs the published `pptx-designer` Python package and copies
the skill bundle to the selected coding assistant:

```powershell
git clone https://github.com/sunchaokun/PPT-Design-Skill.git
cd PPT-Design-Skill

python installer/install.py --platform all --force
python skill/scripts/check_runtime.py
```

The installer supports Claude Code, Codex, DeepSeek Harness, OpenCode, Cursor,
Windsurf, Roo Code, Gemini CLI, Trae, Continue, Droid, KiloCode, Augment, and
GitHub Copilot. See [installer/README.md](../installer/README.md).

### Why LibreOffice is an optional dependency

PPTX generation only requires the Python package `pptx-designer`. LibreOffice
is needed only as the headless rendering fallback when Microsoft PowerPoint
COM is unavailable on Windows. In that fallback, LibreOffice (`soffice`)
converts PPTX to PDF and Poppler (`pdftoppm`) converts the PDF to PNG for
visual QA. It is not required to create the editable PPTX itself.

Run `python skill/scripts/check_runtime.py` to verify the environment. On
Windows, the check also looks in standard LibreOffice installation folders and
the registry, so LibreOffice does not need to be on PATH. Desktop applications
are never installed silently; use `python installer/install.py --render-deps`
only when you explicitly want winget to install LibreOffice and Poppler.

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

## Reviewed case studies

The repository maintains six complete, reviewed case studies. Each package
includes a reproducible source, editable PPTX, PDF export, PNG review evidence,
and a written visual-direction and acceptance record.

| Case | Pages | Design domain |
|---|---:|---|
| AI Agent Operating System | 12 | Technical architecture |
| AI Infrastructure Economics | 12 | Data and infrastructure research |
| Single-Cell CAR T Atlas | 12 | Scientific narrative |
| Louvre Abu Dhabi | 10 | Architecture and culture |
| Vertical City Retrofit | 14 | Urban strategy |
| COUTURE COLOR — Objects of Desire | 10 | Luxury beauty editorial |

Browse the full [case-study library](../examples/README.md) or the published
[gallery](https://sunchaokun.github.io/PPT-Design-Skill/). The Louvre Abu Dhabi
and COUTURE COLOR packages each use one final `build.py` entry point rather
than a collection of sequential draft scripts.

![COUTURE COLOR cover](../examples/site/assets/couture-color-objects-of-desire/slide01.png)

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
