# Curated `pptx-designer` contract

Install the published package and use only public, documented imports. The
package requires Python 3.10+.

## Top-level pipeline

```python
from pptx_designer import (
    Presentation, extract_design_context, extract_design_dna, fetch_image, generate_ppt,
    merge_vi_design_context, validate_resolved_theme,
)
```

- `generate_ppt(query, style=..., output=...)` is the quick path.
- `generate_ppt(content={...}, style=..., output=...)` is the structured path.
- `Presentation(template_path=None, theme=..., strict_theme=False)` creates a
  16:9 presentation and can attach a theme context.
- `extract_design_dna(path)` analyzes an existing presentation.
- `extract_design_context(path)` returns the template context used by VI Build.
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
- Store the `ThemeComposer().compose(...)` result as the resolved theme. It is
  the only object accepted by `generate_ppt(theme=...)`; a Theme Lock is a
  separate skill-level design record and must not be passed directly.
- In ordinary Build Mode, use `Presentation(theme=resolved_theme,
  strict_theme=True)` and semantic roles. A local `C` override is permitted
  only for an explained page-level semantic need; derive it from the current
  resolved theme rather than a stale hard-coded token set.
- Treat library suggestions as candidates. They do not understand every user
  nuance, and a generic preset must not override a domain paradigm or brand
  constraint.

### Theme contract and VI composition

```python
from pptx_designer import (
    Presentation, extract_design_context, merge_vi_design_context,
    validate_resolved_theme,
)
from pptx_designer.renderer.theme import ThemeComposer

resolved_theme = ThemeComposer().compose(style="dark-tech", seed=17)
validate_resolved_theme(resolved_theme)
prs = Presentation(theme=resolved_theme, strict_theme=True)

template_context = extract_design_context(template_path)
page_context = {"page_role": "content"}  # Add only approved page-level fields.
# For a supplied template, locked template fields win over later inputs.
vi_context = merge_vi_design_context(template_context, resolved_theme, page_context)
if vi_context["diagnostics"]["conflicts"]:
    raise ValueError("Revise the theme/page context; it attempted to override a template lock")
```

When `theme=resolved_theme` is passed to `generate_ppt()`, do not also pass
`style`, palette atoms, or `style_seed`; those are discovery inputs and are
ignored for an already resolved theme.

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
Catch invalid or unsafe SVG input with the documented compiler error type:

```python
from pptx_designer import svg_chart
from pptx_designer.compiler import SVGCompileError
```

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
