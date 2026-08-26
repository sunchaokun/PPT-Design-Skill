# Brief-to-delivery workflow

## 1. Brief confirmation

Capture the topic, audience, scenario, decision or action required, language,
duration, target page count, source materials, data provenance, brand rules,
template path, image policy and permitted-use scope, and editability
expectations.

If a decision materially changes the deck, ask before building: audience,
structure, visual direction, or whether the request is a quick draft versus a
delivery-grade deck.

### Brief acceptance contract

Before generation, convert the brief into a compact traceability table:

| ID | User requirement | Visual evidence expected in PNG | Priority |
|---|---|---|---|
| R1 | The deck should feel calm and premium | restrained palette, generous whitespace, controlled imagery | MUST |
| R2 | Explain the three-stage process | a readable process page with three distinct stages | MUST |
| R3 | Preserve the supplied logo | logo appears in the required position and scale | SHOULD |

Use the user's actual requirements, not the example wording above. Every
`MUST` item needs observable evidence in the rendered pages. If a requirement
is ambiguous, clarify it before building or mark the interpretation explicitly.

When the user supplies a reference, add only the material rows below; this is
not a request to copy its content or make a full design audit:

| Observable reference quality | Transfer decision | PNG evidence target |
|---|---|---|
| image scale and restrained type hierarchy | preserve the hierarchy; use project-appropriate imagery | cover first read |

Describe qualities concretely (for example composition, image treatment,
density, type hierarchy, or graphic role), not as labels such as “premium”.

## 2. Visual-first preflight

Before planning copy or coordinates, define the visual standard the client should
feel in the first three seconds. Record:

- the visual thesis in one sentence;
- two or three reference qualities to match (for example: editorial tension,
  scientific precision, or premium restraint);
- the page-level visual anchor for every page;
- the intended occupied zones and the reason for every major empty zone;
- the expected finish level: exploratory, presentation-ready, or client-delivery.

Run a thumbnail test on the planned page map: each page must have a visible
focal point, a clear first read, and enough designed content to feel complete.
Sparse pages are allowed only when the whitespace creates focus, pacing, or
hierarchy. If a page is empty because the content or composition is not planned,
it fails the preflight before any code is written.

Do not let technical feasibility, library presets, or a low density setting
lower the visual target. They are implementation constraints, not the quality
bar.

For supplied references, record whether each quality is preserved, abstracted,
or deliberately not used because of content, brand, or licensing constraints.
This makes the visual direction testable while preserving the freedom to design
an original deck.

## 3. Page plan

For every page define:

- page role and goal;
- one-sentence takeaway;
- evidence or content needed;
- visual form (hero, figure, chart, comparison, process, table, cards, code,
  timeline, or CTA);
- expected density;
- transition from the previous page.

Avoid making every page a title plus bullet list. Mix page roles deliberately,
but do not vary layouts without a communication reason.

## 4. Visual direction

If the user has not specified a visual direction, do not ask an abstract
“what style do you want?” question. Infer the likely design problem and present
two or three plain-language directions with a recommendation. Read
[design-direction-coach.md](design-direction-coach.md) for the response format.

Before presenting those directions, consult the package's built-in palettes,
typography, style presets, and theme composer as candidate evidence. After the
user confirms a direction, call the theme composer again to pin explicit,
reproducible implementation values. Do not allow a generic library preset to
override the user's domain, brand, or visual thesis.

Write a compact design brief containing:

- visual mood and audience fit;
- background, primary, accent, semantic, and text colors;
- heading/body font choices and CJK fallback;
- grid, safe margins, spacing rhythm, and corner treatment;
- image and chart treatment;
- decoration limits;
- pages that should be sparse, dense, or cinematic.

The same direction must govern the entire deck. Small variations are allowed;
unrelated theme changes are not.

## 5. Select the generation path

Use FreeStyle when the library's goal-based renderer is sufficient:

- query input for a fast topic-driven draft;
- structured `content` input when the page goals and copy are already known.

Use Build Mode when exact coordinates, custom composition, advanced diagrams,
or delivery-grade visual control are required. Both paths still go through the
same PDF/PNG rendering and LLM visual review gate.

Use VI Build Mode when the user supplies a PowerPoint template, corporate
master, or explicit brand-compliance requirement. Analyze the template with
`extract_design_dna()`, preserve its framework pages and brand tokens, add new
pages from the template base, and review both inherited and new pages through
the same PDF/PNG gate. Read [template-brand.md](template-brand.md) for the
specific preservation rules and unsupported PowerPoint behaviors.

## 6. Build and inspect loop

Keep the generated Python file as the reproducible source of truth. After each
material revision:

1. run the script;
2. reopen the PPTX with `python-pptx`;
3. export PDF and PNG;
4. inspect the PNGs against every acceptance-contract row;
5. record `PASS`, `NEEDS_REVISION`, or `BLOCKED`, plus evidence and the next
   revision;
6. repeat until all `MUST` rows pass.

When a review failure is directional (reference transfer, image strategy,
composition, hierarchy, or page architecture), return to that decision before
making local formatting edits. Use local edits for local defects such as
overflow, alignment, or contrast.

Do not patch only the exported PPTX when the change should be reproducible.

## 7. User confirmation

For delivery-grade work, confirm the page plan and visual direction before the
full build, then confirm the reviewed PNG result before final delivery. For a
quick draft, the first confirmation may be deferred, but PNG review is still
required before claiming the draft is usable.
