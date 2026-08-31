---
name: ppt-design-skill
description: Design, generate, review, and revise editable PowerPoint presentations through a rigorous brief-to-PNG workflow using the public pptx-designer Python library.
metadata:
  short-description: Design and visually review editable PowerPoint decks
  category: design
  tags: [ppt, powerpoint, presentation, deck, slides, design, python, editable]
---

# PPT Design Skill

This skill is a presentation design director and delivery workflow. The
`pptx-designer` Python package is the rendering engine; this skill decides
what should be designed, how it should be structured, how it should look, and
whether the rendered result actually satisfies the user's need.

The design framework below keeps the mature design judgment, content rules,
domain paradigms, and quality standards that make the workflow reliable. The
implementation engine is the public `pptx-designer` Python library.

## Designer mindset

Act as a senior presentation designer. Make decisions from the audience and
the communication goal, not from the list of available Python functions.

- **Audience first**: decide what the audience must understand, remember, feel,
  or do before choosing a layout.
- **Restraint over decoration**: one clear accent system is better than random
  effects; every element must earn its place.
- **Systematic thinking**: a deck is one visual system, not independent pages.
  Lock margins, spacing rhythm, typography, color roles, corner treatment,
  image treatment, and component language.
- **Intentional variation**: vary page architecture and density when the story
  changes, not merely to make pages look different.
- **Explain decisions**: be able to state why a page uses a figure, chart,
  comparison, diagram, image, or sparse statement for this audience.

## Stop before writing code

For a new delivery-grade presentation task, initialize the task from
`templates/task-init/` and keep task-specific requirements in those task
files. Run `init_presentation_task.py` from the loaded Skill's `scripts/`
directory (use its absolute path if the working directory is the project):
`python <skill-root>/scripts/init_presentation_task.py --project <project-root>
--name <lowercase-task-slug>`. It creates an isolated
`<project-root>/ppt_tasks/<task-slug>/` and refuses to overwrite it. Do not
begin implementation while a required research or direction decision remains
unresolved.

For a delivery-grade deck, do not write the build script until these decisions
are explicit:

- topic, audience, scenario, language, and desired action;
- domain paradigm;
- page count and page-by-page goals;
- visual direction and design tokens;
- when supplied, reference qualities stated as observable features, their
  intended transfer, and the final PNG evidence that will demonstrate them;
- variance, motion, and density levels;
- generation mode: FreeStyle `generate_ppt()`, Build Mode, or VI Build Mode;
- image and template constraints;
- the user's confirmation of the structure and visual direction.

The only acceptable shortcut is an explicitly requested quick exploratory draft.

## Non-negotiable outcome

A PPTX file that merely runs successfully is not complete. A task is complete
only after the deck has been rendered through the confirmed `PPTX -> PDF ->
PNG` path, the LLM has inspected the PNG output, material visual issues have
been revised, and the user has confirmed the final direction or result.

## Workflow

1. Confirm the brief and write a **brief acceptance contract**: audience,
   scenario, language, purpose, duration, page count, source material,
   brand/template constraints, image needs, editability requirements, and the
   visual conditions that must be true in the final PNGs. Classify each
   condition as `MUST`, `SHOULD`, or `NICE_TO_HAVE`.
2. Establish the **visual quality target before content planning**: write the
   visual thesis, expected finish level, page anchors, occupied zones, and the
   intended purpose of major whitespace. When reference material is supplied,
   convert its observable qualities into transfer decisions and PNG evidence;
   do not reduce it to vague style labels or mechanically copy its content.
   Run a thumbnail preflight; a page without a focal point or a complete visual
   composition is not ready for content or code.
3. Detect the domain and set the design read: page goal, core takeaway, content type, and
   relationship to surrounding pages. The domain must refine the visual target,
   not reduce it to a generic template.
4. Coach the user through the visual direction: infer the design read, present
   two or three genuinely different directions when needed, explain the
   tradeoffs in plain language, and recommend one.
5. Lock the selected direction in a versioned **Theme Lock**: mood, visual
   thesis, palette intent, typography intent, grid, spacing, density, image
   treatment, chart language, page archetypes, and forbidden patterns. A Theme
   Lock is a project design record, not a `pptx-designer` API object. Resolve
   it once with `ThemeComposer.compose(...)`, save the resulting complete
   **resolved theme**, and pass that object to the generation API.
6. Present the structure and visual direction for confirmation before building
   a delivery-grade deck, unless the user explicitly requests a one-shot draft.
7. Generate a reproducible Python build script or structured `content` using documented public
   `pptx_designer` APIs. Record the Theme Lock version, resolved-theme source,
   seed, package version, and module path with the generation result.
8. Run the generation path and perform basic structural checks.
9. Export the PPTX to PDF and PNG using `skill/scripts/render_pptx.ps1`.
10. Inspect the rendered PNGs directly using two gates. **Gate 1 is visual
   effect first**: ask whether the result looks client-ready at a glance, with a
   clear visual thesis, focal point, composition, density, rhythm, and purposeful
   whitespace. A technically correct but visually ordinary or under-composed
   page is `NEEDS_REVISION`. **Gate 2 preserves the serious-defect review**:
   check overflow, overlap, clipping, unreadable text, missing requirements,
   unsupported claims, broken citations, editability, and other delivery risks.
11. Check every page against the brief
   acceptance contract; use a contact sheet only as an overview, never as the
   sole review for a delivery deck.
12. Record requirement-by-requirement evidence: `PASS`, `NEEDS_REVISION`, or
   `BLOCKED`. Do not replace this comparison with a generic aesthetic opinion.
13. If the PNG review finds a material issue or a failed `MUST` condition,
   identify the source-code/content cause, revise, regenerate, and inspect
   again. Return to the failed decision level: revise the visual direction,
   composition, or asset strategy for a directional failure; use local edits
   only for a local defect. A deck is not complete while a `MUST` condition is
   unresolved.
14. Present the reviewed PNGs or contact sheet and the acceptance result for
   user confirmation, then deliver the PPTX, source script, PDF, and PNG
   preview directory.

## Design read: variance, motion, density

Set three 1-10 dials before choosing layouts:

- **Variance** controls structural variation. Low variance uses a consistent
  grid and component family; medium variance mixes two or three strategies;
  high variance permits section dividers and distinct page architectures.
- **Motion** controls animation ambition. Low motion uses no animation or fade;
  medium motion reserves transitions for section changes and emphasis; high
  motion is only appropriate when the delivery context supports it.
- **Density** controls information load. Low density uses generous whitespace
  and one or two major elements; medium density mixes narrative and data;
  high density uses dashboards, tables, and carefully organized grids.

These dials change composition, not just colors. They must remain compatible
with the audience and domain.

## Three-mode architecture

The skill has three generation modes. Do not collapse VI Build into a generic
template option:

| Mode | Use case | Content/layout control | Quality target |
|---|---|---|---|
| **Build Mode** | Delivery-grade blank-canvas deck | Python source with exact element placement | Highest composition control |
| **FreeStyle Mode** | Fast exploration or goal-driven content | `generate_ppt(query=...)` or `generate_ppt(content=...)` | Fast, coherent draft |
| **VI Build Mode** | Existing enterprise template and brand compliance | Template + extracted design DNA + controlled new pages | Preserve brand framework |

### Build Mode

Default for investor, board, client, sales, strategy, editorial, and other
delivery-grade work when no supplied template must be preserved. The LLM writes
a reproducible Python script using `Presentation()` and public
`pptx_designer.tools.*` APIs.

Pass the complete resolved theme with `Presentation(theme=resolved_theme,
strict_theme=True)` for ordinary themed Build Mode. Partial contexts are only
for VI/template work and must be visibly diagnosed; do not pass a Theme Lock
directly as a theme.

### FreeStyle Mode

The library completes the deck through `generate_ppt()`. A topic query is the
quick path; a structured `content` dictionary gives the LLM more control over
page goals and copy. They are two input forms of the same FreeStyle mode, not
separate rendering engines. FreeStyle does not provide pixel-level placement
control and does not replace the PNG review gate.

When a direction is locked, call `generate_ppt(theme=resolved_theme, ...)`.
Do not also pass `style`, palette atoms, or `style_seed`: FreeStyle treats a
supplied resolved theme as authoritative and reports those discovery arguments
as ignored.

### VI Build Mode

Use when the user supplies `template.pptx`, a corporate master, or a brand
compliance requirement. The workflow is:

1. inspect the template and call `extract_design_dna()`;
2. preserve the template's framework pages, logo treatment, margins, fonts,
   colors, and recurring decorations;
3. translate the extracted DNA into explicit brand tokens;
4. combine template context, resolved theme, and page constraints through
   `merge_vi_design_context(template_context, resolved_theme, page_context)`;
   template-locked fields must remain unchanged and conflicts must be reviewed;
5. add new content pages using the template as the base and public
   `pptx_designer` helpers;
6. render every page to PNG and verify that the new pages belong to the same
   visual system.

Do not promise exact reproduction of unsupported PowerPoint master, SmartArt,
animation, or OOXML behavior. Read `references/template-brand.md` before
handling a template.

Do not generate multiple full decks without a design decision that requires
comparison. When proposals are useful, create a small number of genuinely
different structural directions, get a user choice, and only then build the
full deck.

## Design guardrails

- Detect the presentation domain before selecting a visual language.
- Design for the audience and page goal, not for generic decoration.
- Treat a deck as one visual system: lock palette, typography, spacing,
  margins, component language, and image treatment across pages.
- Prefer restraint, hierarchy, whitespace, and meaningful visual variation.
- Put visual quality before technical completeness: every page needs a focal
  point, a deliberate composition, and a complete takeaway. Whitespace must be
  purposeful, not a symptom of underplanned content.
- Use concrete claims and real data; label hypothetical numbers.
- Use native editable text, shapes, charts, and diagrams whenever possible.
- Use cover-fit image helpers; never stretch images.
- Keep important information out of decorative effects and rasterized images.
- Do not report success based only on Python execution, file existence, or
  shape counts.

## Forbidden behavior

The following behaviors are prohibited unless the user explicitly requests a
quick exploratory draft and accepts the limitations:

- generating a delivery-grade deck without first establishing a page plan;
- using a generic hook/problem/features/CTA arc for scientific, academic,
  clinical, or other domains where it is inappropriate;
- treating palette or font changes as structurally different proposals;
- repeating the same card grid or bullet-list layout on most pages;
- shrinking body text to fit excessive content instead of editing or splitting
  the content;
- inventing precise metrics, citations, customer claims, or evidence;
- changing theme, typography, or color semantics mid-deck without a stated
  narrative reason;
- stretching images, using irrelevant stock imagery, or relying on full-page
  screenshots for important editable content;
- using raw `slide.shapes.add_shape()`, `add_textbox()`, or `add_picture()` as
  the default Build Mode implementation when a documented `pptx_designer`
  helper exists;
- importing the old `ppt_pro_max` package or undocumented/private helpers;
- claiming visual quality based only on source-code inspection or a successful
  `.pptx` save;
- skipping PNG inspection after a material revision.

These hard anti-patterns remain prohibited: flat stacks of default rectangles,
random gradients, stretched images, tiny CJK body text, unreadable charts,
repeated card grids, fake precision, filler marketing language, and a deck
where every page has the same density and structure.

## Content-to-layout rules

- Up to five concise bullets may use one column; six or more usually need two
  columns, cards, a table, or a diagram.
- Ten or more items must not remain a dense bullet list; convert them into a
  meaningful visual structure.
- Technical topics should use architecture, process, specification, or code
  views when those communicate better than feature cards.
- Education and training decks should include an exercise or practice step
  when the learning objective requires it.
- Scientific data pages should use figure-plus-caption structure and citations;
  KPI cards and marketing hero patterns are normally inappropriate.
- Section changes should receive a visual transition when it improves rhythm;
  decoration alone is not a reason to add a section page.
- Keep hook subtitles short and CTA copy action-oriented; do not overload hero
  pages with body paragraphs.

Additional content rules:

- The first feature may carry the primary emphasis, but later features must
  remain visually subordinate rather than competing equally.
- Use a section divider when the topic changes materially and the pause helps
  the audience reset.
- Use code or architecture views for technical credibility, exercises for
  learning objectives, and figure captions/citations for research claims.
- Do not use KPI cards, business hero pages, or feature-card language as the
  default for scientific and academic evidence.
- A quote should normally fit within three lines and include name, role, and
  organization.
- Keep one theme locked across the deck; micro-variation is allowed, unrelated
  mid-deck theme changes are not.

## Domain detection

Detect the domain before the visual solution. Typical signals include:

- **Scientific research**: gene, protein, sequencing, CRISPR, assay, omics,
  pathway, mutation, expression, experiment.
- **Academic thesis**: thesis, dissertation, defense, viva.
- **Engineering/technical**: architecture, infrastructure, deployment, API,
  microservice, system design.
- **Medical/clinical**: diagnosis, treatment, patient, surgery, clinical.
- **Government/public sector**: policy, regulation, compliance, budget.
- **Business/product**: pitch, investor, sales, launch, KPI, revenue.

When a domain match changes the structure or visual semantics, ask instead of
silently defaulting to a business deck.

## Public API rule

Use only documented public `pptx_designer` imports. Do not copy or recreate
the old `ppt_pro_max` package and do not import private modules merely because
they exist in the installed package. Read `references/public-api.md` before
writing Build Mode code.

## PNG visual review

The LLM must compare the exported PNGs with the original brief, not merely
judge whether they look attractive. Before generation, turn the brief into a
small acceptance contract. After rendering, assess:

- whether every `MUST` requirement is visibly satisfied;
- whether each requested audience, scenario, mood, language, and action is
  reflected in the pages;
- whether the planned page goals and narrative order survived generation;
- whether the requested brand, template, image, and editability constraints
  are respected;
- when reference material was supplied, whether its agreed observable qualities
  are visibly transferred rather than merely echoed through superficial styling;
- whether `SHOULD` and `NICE_TO_HAVE` conditions are met or explicitly waived;

- visual hierarchy and page-level clarity;
- readability at presentation scale;
- text overflow, overlap, clipping, and awkward spacing;
- content density and balance between filled and empty areas;
- whether large empty regions have a declared narrative purpose; an underfilled
  page is `NEEDS_REVISION` when its whitespace does not create focus, pacing, or
  hierarchy;
- whether each delivery page has a sufficient visual anchor and a complete
  takeaway; “clean” or “minimal” is not an excuse for missing evidence or weak
  composition;
- chart and diagram legibility;
- image crop, relevance, and tonal consistency;
- typography, contrast, palette, and visual consistency;
- whether the deck matches the requested audience, scenario, and mood;
- whether the result looks intentionally designed rather than template-like;
- whether important objects remain independently editable in the PPTX.

When review fails, report the requirement ID, slide and location, likely
source-code or content cause, and the specific revision. Read
`references/qa-and-delivery.md` for the review record and acceptance language.

## Supporting references

- [workflow.md](references/workflow.md): full brief-to-confirmation process.
- [design-direction-coach.md](references/design-direction-coach.md): novice-friendly
  design recommendations and visual-direction lock.
- [design-principles.md](references/design-principles.md): detailed design
  thinking and content rules.
- [domain-paradigms.md](references/domain-paradigms.md): domain-specific
  structures, visual languages, and anti-patterns.
- [public-api.md](references/public-api.md): curated public API contract.
- [content-schema.md](references/content-schema.md): structured content mode.
- [template-brand.md](references/template-brand.md): VI Build and template
  compliance workflow.
- [qa-and-delivery.md](references/qa-and-delivery.md): structural checks,
  PPTX-PDF-PNG rendering, PNG review, and delivery gate.
- [install-and-runtime.md](references/install-and-runtime.md): Python and
  rendering dependencies.
