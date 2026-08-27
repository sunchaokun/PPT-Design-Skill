# Acceptance contract — Vertical City Retrofit

## Must

- 14 editable 16:9 slides with a coherent architectural editorial system.
- At least four genuinely different page architectures.
- At least one editable architectural section, one systems diagram, one native
  chart, one phasing timeline, and one decision matrix.
- Source notes on evidence-led pages; illustrative assumptions clearly labeled.
- No repeated card-grid treatment across the deck.
- No rounded UI-like containers as the primary visual language.
- All body text readable at presentation scale.
- Final slide states the strategic decision and the evidence boundary.

## Must not

- Treat a generated architectural image as a real project photograph.
- Present scenario economics as measured project outcomes.
- Fill pages with decorative building imagery without an analytical role.

## QA

- Run `inspect_pptx.py` after generation.
- Render PPTX to PDF and PNG.
- Inspect every PNG at 1280 × 720.
- Revise any clipping, weak hierarchy, tiny text, empty containers, or repeated
  layout before reporting completion.
