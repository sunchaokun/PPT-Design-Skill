# Design principles

These principles define the design judgment expected from the skill.

## Designer mindset

- Start with the audience: what must they understand, feel, or decide?
- Prefer restraint over decoration. Every element must earn its place.
- Treat the deck as a visual system, not a set of unrelated pages.
- Use deliberate contrast in scale, density, rhythm, and page role.
- Explain design choices in terms of communication intent.

## Content rules

- Use one main idea per page.
- Keep hook subtitles short; let CTA pages carry the fuller action statement.
- Use concrete claims and real data; mark examples and assumptions explicitly.
- Do not use filler verbs or generic claims such as “empower”, “leading”,
  “one-stop”, “ecosystem”, or “revolutionary” unless the user supplied them.
- Six or more bullets usually require two columns, cards, a table, or a
  diagram; never compress a long list into tiny text.
- Vary density intentionally, but keep the visual system locked.
- Technical topics benefit from code, architecture, specification, or process
  views; education benefits from exercises; evidence-heavy topics need source
  labels and captions.
- Quotes should be short fragments with attribution including name, role, and
  organization.

## Visual anti-patterns

- repeated card grids on every page;
- decorative gradients with no semantic purpose;
- tiny body text used to preserve too much content;
- stretched or irrelevant images;
- charts without a clear takeaway or readable labels;
- random colors and mid-deck theme changes;
- full-page screenshots used instead of editable content;
- fake precision in metrics;
- generic hero/problem/features/CTA structure for scientific, academic, or
  clinical material.

## Proposal rule

A proposal is structurally different only when page architecture, data form,
cover treatment, typography scale, or visual language changes. Palette-only
variants are not separate design proposals.

For delivery-grade work where the direction is not already fixed, prepare up
to three genuinely different directions before the full build. Each direction
must state its audience fit, page architecture, visual language, and tradeoff.
Do not generate three full decks merely to compare colors. If the user has
already approved a clear direction, proceed with one page plan.

## Quantified design guardrails

These are pre-flight defaults, not substitutes for judgment:

- body text should normally be at least 14 pt for projected decks; 11 pt is a
  lower bound for captions or dense reference material, not a target;
- keep one primary takeaway per page and avoid more than about 50 independent
  shapes unless the density is intentional and the result remains legible;
- use at least four meaningful type-size levels only when the hierarchy needs
  them; do not create scale variation as decoration;
- keep important content inside a consistent safe margin and never rely on
  animation, hover, or color alone to communicate meaning;
- dark themes should use near-black and near-white values, with contrast
  checked on the actual rendered PNG;
- vary page structure when variance is above the middle range, and follow a
  high-density page with a breathing page when the narrative permits it.

## Motion and accessibility

Motion must have a communication purpose. Use no motion or a simple fade for
scientific, academic, medical, government, print, and accessibility-sensitive
decks unless the user explicitly requests otherwise. For an unknown audience,
keep motion conservative. Every important message must remain understandable
in a static PDF/PNG export.

Check contrast, font fallback (including CJK), reading order, caption clarity,
projection readability, and whether charts remain interpretable without color
alone. The PNG review is the final evidence for these checks.

## Redesign protocol

When revising an existing deck, first record what must be preserved: page order,
navigation labels, logos, legal copy, data, and brand tokens. Then identify
patterns to keep, patterns to remove, and the smallest modernization lever
that solves the problem. Do not silently change content semantics or brand
meaning while changing visual structure. For a template-driven redesign, use
VI Build Mode and read `template-brand.md` before editing.

## Pre-flight checklist

Before rendering, check:

- content is real, sourced, or explicitly marked hypothetical;
- no page has overflow, overlap, broken image, or placeholder text;
- fonts, colors, margins, corner treatment, and image crop rules are locked;
- adjacent pages do not repeat the same structure without a narrative reason;
- cover, section transitions, dense data pages, and CTA/end pages have distinct
  roles;
- the deck remains understandable when exported as static PNGs.

## Generation-path distinction

`generate_ppt()` is the FreeStyle generation path. A topic query and a
structured `content` dictionary are two input forms for that same path. Use
Build Mode when the task requires explicit per-element placement, custom page
architecture, or a reproducible design system that the goal-based renderer
cannot express.
