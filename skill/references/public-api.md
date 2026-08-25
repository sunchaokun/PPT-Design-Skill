# Curated `pptx-designer` contract

Install the published package and use only public, documented imports. The
package requires Python 3.10+.

## Top-level pipeline

```python
from pptx_designer import Presentation, extract_design_dna, fetch_image, generate_ppt
```

- `generate_ppt(query, style=..., output=...)` is the quick path.
- `generate_ppt(content={...}, style=..., output=...)` is the structured path.
- `Presentation(template_path=None)` creates a 16:9 presentation.
- `extract_design_dna(path)` analyzes an existing presentation.
- `fetch_image(...)` is optional and may require image credentials.

## Design intelligence and theme atoms

The package includes reusable design data that the skill may consult during
direction design and implementation:

```python
from pptx_designer import PALETTES, STYLES, TYPOGRAPHY, recommend_styles
from pptx_designer.renderer.theme import ThemeComposer
from pptx_designer.search.adapters import search_color, search_style, search_typography
```

- Use `recommend_styles()` and the search functions before direction lock to
  generate candidate vocabulary and options.
- Use `ThemeComposer().compose(...)` after direction lock to pin explicit
  palette, font, decoration, layout, mood, and seed choices.
- For reproducible Build Mode work, copy the selected roles into an explicit
  `C` dictionary and typography configuration rather than relying on implicit
  random defaults.
- Treat library suggestions as candidates. They do not understand every user
  nuance, and a generic preset must not override a domain paradigm or brand
  constraint.

## Build Mode modules

```python
from pptx_designer.tools.cards import cta_slide, hero_slide, kpi_card, highlight_cards, section_divider
from pptx_designer.tools.charts import bar_chart, comparison_bars
from pptx_designer.tools.images import circle_image, cover_image
from pptx_designer.tools.layout import page_header, page_number, top_bar
from pptx_designer.tools.shapes import arrow, diamond, hexagon, oval, rect, rrect
from pptx_designer.tools.text import dramatic_text, gradient_text, multiline, text, vertical_text
```

All positions and dimensions are inches. Use named arguments such as
`left`, `top`, `width`, and `height`. `rrect` is the documented rounded
rectangle helper in the current package.

## Diagrams and SVG

Use diagram classes from `pptx_designer.diagrams` and call `.render(slide)`.
Use `svg_chart()` for supported editable SVG and inspect its `warnings`.
Catch `SVGCompileError` for invalid or unsafe SVG input.

## Reliable authoring rules

- Use a blank slide layout and explicit coordinates in Build Mode.
- Keep colors in a `C` dictionary or explicit theme tokens.
- Prefer native shapes, text, charts, and diagrams.
- Use `cover_image()` for image placement.
- Avoid private modules and invented signatures.
- Reopen and render the generated PPTX before reporting success.

For exact signatures and current availability, prefer the installed package's
maintained documentation and examples over memory. The skill should be
updated when the package's public contract changes.
