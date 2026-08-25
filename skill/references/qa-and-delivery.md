# QA and delivery

Visual QA is performed by the LLM reviewing rendered PNGs. This skill does not
implement a separate visual model, scoring service, or automatic visual
revision engine.

## Basic structural checks

Before visual review:

- the build script exits successfully;
- the PPTX exists at the requested path;
- the PPTX can be reopened;
- slide count and page size are reasonable;
- expected images exist;
- SVG errors are absent and warnings are understood;
- important content is not a single rasterized page image.

Run `skill/scripts/inspect_pptx.py` for a compact report.

## Confirmed render path

Run on Windows:

```powershell
powershell -ExecutionPolicy Bypass -File skill/scripts/render_pptx.ps1 `
  -InFile output/deck.pptx `
  -OutDir output/rendered
```

The script prefers PowerPoint COM. It exports a PDF and each slide as a 1280×720
PNG. If PowerPoint COM is unavailable, it falls back to LibreOffice for PDF
conversion and Poppler's `pdftoppm` for PNG conversion.

## PNG review checklist: two-gate review

The review has two gates. The visual gate comes first because it catches weak
composition, underfilled pages, generic template feel, and a missing visual
thesis at the moment they are easiest to recognize. The defect gate is retained
from the mature workflow and must still catch serious design and delivery risks.

### Gate 1 — visual effect first

Look at the contact sheet and then each PNG at presentation scale. Ask first:

- Does the deck look client-ready at a glance?
- Does every page have a strong visual anchor, first read, and complete takeaway?
- Is the visual direction obvious and consistently executed?
- Are density, rhythm, composition, and whitespace intentional?
- Does the result feel designed for this audience and domain rather than like a
  generic generated template?

If the answer is no, mark the relevant slide `NEEDS_REVISION` even when the file
is technically valid and no object overlaps. Record the visual cause and revise
the source before spending time on final delivery checks.

### Gate 2 — serious defects and delivery risk

First compare the PNGs with the brief acceptance contract. For every row,
record status and evidence:

```text
R1 [MUST] PASS
Evidence: slides 1 and 4 use the approved dark palette and leave the requested
negative space around the product imagery.

R2 [MUST] NEEDS_REVISION
Evidence: slide 3 contains the three stages, but stage 2 is not readable at
presentation scale.
Action: enlarge the stage labels and rerender.
```

Then inspect every slide and record:

- hierarchy and first-read message;
- alignment, spacing, balance, and safe margins;
- readability at 16:9 presentation scale;
- text overflow, clipping, and overlap;
- chart, diagram, and image legibility;
- palette, typography, contrast, and consistency;
- domain and audience fit;
- whether the page looks deliberate rather than generic.

Apply a higher bar to whitespace: first identify the intended role of the empty
area (focus, pacing, hierarchy, or image breathing room). If no such role is
visible, treat a large unoccupied region as a design defect and mark the page
`NEEDS_REVISION`. Every page should have a clear visual anchor and a complete
takeaway. Minimalism means reducing nonessential elements, not delivering a
partially planned page.

Example issue record:

```text
Visual review: NEEDS_REVISION
Requirement: R2 [MUST]
Slide 4, right chart: labels are too small and the chart has six low-value categories.
Cause: chart region is too narrow for the selected data.
Action: keep the four material categories, enlarge labels, and rerender.
```

Do not claim `PASS` while a material issue remains. After every material
revision, rerun the build, render, and PNG review.

## Delivery package

For a delivery-grade task, provide:

- final `.pptx`;
- reproducible Python build script and structured content, if used;
- exported `.pdf`;
- PNG preview directory or contact sheet;
- concise structural and visual review result.

The user remains the final approver of subjective visual direction.
