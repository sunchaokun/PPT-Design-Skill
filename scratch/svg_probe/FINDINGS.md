# SVG → editable-pptx probe: findings

Date: 2026-08-13 · Scratch experiment (self-contained, no `src/` pollution)

## Question being answered

In Build mode we commit to **python-pptx, native editable shapes** (no raster images).
LLMs are unreliable at hand-computing geometry. Can an **SVG-subset → editable-pptx
compiler** give us SVG's declarative speed while keeping native editability?

## Verdict

**Feasible.** A ~600-line probe compiler built on existing primitives
(`freeform_builder`, `boolean_shapes`, `text_measurer`, `build_helpers`)
compiles real business-chart SVGs to **100% native editable shapes** (zero
`<p:pic>`) with **pixel-faithful geometry**, in **5–20 ms per chart**.

## Results (ground truth = Edge-rendered SVG PNG, result = LibreOffice-rendered pptx)

| case            | shapes | editable | features                       | ink-IoU | mean-diff | compile |
|-----------------|--------|----------|--------------------------------|---------|-----------|---------|
| pyramid         | 6      | ✓ 0 pics | gradient, polygon, text        | **0.98**| 1.9       | 13 ms   |
| venn_evenodd    | 4      | ✓ 0 pics | path(arc), text                | **0.83**| 5.2       | 9 ms    |
| funnel          | 8      | ✓ 0 pics | gradient, polygon, text        | **0.98**| 1.4       | 9 ms    |
| growth_curve    | 5      | ✓ 0 pics | clipPath, gradient, path(C), text | **0.91**| 2.6     | 13 ms   |
| matrix_bcg      | 13     | ✓ 0 pics | rect, circle, line, group, transform, text | **0.91** | 0.9 | 21 ms |
| unsupported     | 2      | ✓ 0 pics | image/filter/mask REFUSED      | —       | —         | 3 ms    |

## What works (compiler boundary)

- **Elements**: `rect`(→native RECTANGLE), `circle/ellipse`(→sampled polygon →
  custGeom), `polygon`, `polyline`, `line`, `path` (`M/L/H/V/C/S/Q/T/A/Z`,
  relative+absolute). Arc→cubic conversion verified (incl. the "full circle"
  `a r,r 0 1,0` idiom — this was a subtle sign/convention bug).
- **Boolean**: `fill-rule="evenodd"` and `clipPath` → Shapely union/intersection →
  `custGeom` (topology hardened with `buffer(0)` + `make_valid`).
- **Transforms**: `g` with `translate/scale/rotate/matrix` (nested), composed
  into a single affine applied to geometry.
- **Paints**: solid + `linearGradient` (→ `a:gradFill` with `a:lin` angle),
  `fill-opacity`/`stop-opacity`/`stroke`, `stroke-width`.
- **Text**: `text`→real textbox, CJK-capable, anchor-aware, auto-height via
  `text_measurer` (approximation — see boundary).
- **Editability**: every output is `<p:sp>` custGeom / MSO shape / `<p:txBody>`;
  verified zero `<p:pic>` in the slide XML.

## Boundaries / where it must REFUSE (not degrade)

- `image` element → **refused** (would become a picture = non-editable).
- `filter` / `mask` / `filter=` / `mask=` attrs → **refused** (decorative,
  raster-based; not silently approximated).
- `text` inside `clipPath` → refused (clipped text needs per-glyph mask).
- Text metric approximation: SVG baseline vs pptx vertical-anchor — small
  vertical drift (venn 0.83 vs 0.98), because we approximate baseline with
  `estimate_text_size`. Fix: baseline math + per-font metrics.

## Efficiency vs. LLM hand-drawing

- Compile is **5–21 ms/chart**; the bottleneck is not geometry but the LLM
  producing the SVG (LLM-native task).
- The "10,000 ways to draw a model" long-tail is handled by the **language**
  (paths/beziers/booleans/transforms), not by a finite template library.

## Repro

```
$env:PYTHONPATH="src"
python scratch/svg_probe/probe.py      # → out/REPORT.md + per-case .pptx/.png
```

Requires: python-pptx, shapely, lxml, Pillow, numpy, Edge (ground truth),
LibreOffice + poppler (result render, auto-detected by build_helpers.preview).

## Next steps (if productionized)

1. `src/ppt_pro_max/renderer/svg_compiler.py` — port probe logic, add
   `group_builder` (native `p:grpSp` for `<g>`), baseline-accurate text,
   `stroke-dasharray`, radial gradients, `<use>`.
2. Guard: any unsupported feature → raise (never silently degrade to picture).
3. Case library = SVG skeletons only; style injection happens at compile time
   (palette/typography → colors/fonts), keeping the 40k-combo design system.
