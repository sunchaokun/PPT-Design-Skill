# SVG Compiler Revision Plan

## Overview
Address the 12 issues identified in `docs/svg_compiler_audit.md`. Prioritized by impact and dependency chain.

---

## Phase 1: Critical Fixes (Foundation)

### 1.1 Fix `_resolve_svg_color` Type Contract ✅ DONE
- **File**: `src/ppt_pro_max/renderer/svg_compiler/_compiler.py`
- **Change**: Return sentinel `"none"` instead of `None`; or fix all downstream handlers to handle `Optional[str]` explicitly.
- **Decision**: Use sentinel `"none"` for `none` keyword; empty string remains empty.
- **Validation**: Run `tests/test_svg_paint.py` + add new test asserting `"none"` returns `"none"`.
- **Status**: COMPLETED — tests in `tests/test_svg_color_contract.py` added and passing.

### 1.2 Remove Redundant `_BASELINE_MAP` ✅ DONE
- **File**: `src/ppt_pro_max/renderer/svg_compiler/_text.py`
- **Change**: Delete the first definition (lines 42-51), keep the second (lines 94-103).
- **Validation**: Re-run `tests/test_svg_text.py`.
- **Status**: COMPLETED — no duplicate definitions remain.

---

## Phase 2: Feature Gaps (High Impact)

### 2.1 Wire Up Stroke Styles ✅ DONE
- **Files**: `_dash.py`, `_compiler.py`
- **Change**: 
  - In `_render_shape`, call `parse_stroke_style(el)` to extract `stroke-dasharray`, `stroke-linecap`, `stroke-linejoin`.
  - Pass result to native shape creation or apply via `apply_stroke_style()`.
  - Fixed `apply_stroke_style` to handle both `CT_Shape` and raw lxml elements.
- **Validation**: Add new test in `tests/test_svg_dash.py` covering end-to-end stroke style rendering.
- **Status**: COMPLETED — integration tests added and passing (282 total SVG tests).

### 2.2 Apply `gradientTransform` ✅ DONE
- **File**: `src/ppt_pro_max/renderer/svg_compiler/_paint.py`
- **Change**: Parse the `gradientTransform` attribute into an `Affine` matrix; multiply (x1,y1,x2,y2) before computing PPT angle/tile.
- **Validation**: Add unit test for rotated 45° gradient.
- **Status**: COMPLETED — `apply_gradient()` now applies `transform` to linear gradient (x1,y1,x2,y2) and radial gradient (cx,cy,r), with 3 new tests passing.

### 2.3 Support `<use>` Element ✅ DONE
- **File**: `_compiler.py:_render`
- **Change**: Maintain a `defs` dict during preprocessing; when encountering `<use href="#id">`, clone the referenced element and apply `(x, y)` offsets.
- **Validation**: Add test using a symbol library.
- **Status**: COMPLETED — supports `href`, `xlink:href`, and `SVG`-namespaced href; nested `<g>` references; 3 tests passing.

### 2.4 Fix `tspan` Layout
- **File**: `_text.py:_render_tspan`
- **Change**: 
  - Use `dx`/`dy` as cumulative offset within the paragraph.
  - If `tspan` has absolute `x`/`y`, set `space_before` on the paragraph to preserve visual spacing.
- **Validation**: Add test with multi-tspan text and verify Y positions.
- **Status**: COMPLETED — `_render_tspan_text` now computes `space_before` from absolute `y` differences; 5 tests passing.

### 2.4 Fix `tspan` Layout ✅ DONE
- **File**: `_text.py:_render_tspan`
- **Change**: 
  - Use `dx`/`dy` as cumulative offset within the paragraph.
  - If `tspan` has absolute `x`/`y`, set `space_before` on the paragraph to preserve visual spacing.
- **Validation**: Add test with multi-tspan text and verify Y positions.
- **Status**: COMPLETED — `_render_tspan_text` now computes `space_before` from absolute `y` differences; 5 tests passing.

---

## Phase 3: Robustness (Medium Impact)

### 3.1 Sanitizer Regex Hardening ✅ DONE
- **File**: `_sanitizer.py:_fix_self_closing`
- **Change**: Use `lxml` to detect tags with children vs. self-closing; only force-close truly empty tags.
- **Validation**: Add test with `<rect><title>x</title></rect>` to ensure structure preserved.
- **Status**: COMPLETED — replaced regex-based `_fix_self_closing` with lxml-based `_fix_self_closing_lxml`; 13 sanitizer tests passing.

### 3.2 Handle `rx`/`ry` Symmetry ✅ DONE
- **File**: `_compiler.py:_render_shape` (rect path)
- **Change**: If `rx` is missing but `ry` is present, set `rx = ry` before native rect creation.
- **Validation**: Add test for `<rect ry="10">` (no rx).
- **Status**: COMPLETED — compiler now checks both `rx` and `ry` to decide freeform vs native path.
- **File**: `_compiler.py:_render_shape` (rect path)
- **Change**: If `rx` is missing but `ry` is present, set `rx = ry` before native rect creation.
- **Validation**: Add test for `<rect ry="10">` (no rx).
- **Status**: COMPLETED — compiler now checks both `rx` and `ry` to decide freeform vs native path.

---

## Phase 4: Polish (Low Priority)

### 4.1 Add `scaling_mode` Parameter
- **File**: `_compiler.py:compile()`
- **Change**: Accept `scaling_mode: Literal["contain", "cover", "stretch"] = "contain"`.
- **Validation**: Add test for all 3 modes.
- **Status**: PENDING.

### 4.2 Optimize Overlap Detection
- **File**: `_compiler.py:_detect_text_overlaps`
- **Change**: If `n_shapes > 100`, use spatial bucketing (grid-based).
- **Validation**: Benchmark test on 200-shape SVG.
- **Status**: PENDING.

### 4.3 Standardize DPI
- **File**: `_text.py` and `text_measurer.py`
- **Change**: Use `96.0` consistently; document the choice in module docstring.
- **Validation**: Visual regression test on known font widths.
- **Status**: PENDING.

---

## Execution Order

1. **1.1 → 1.2** (foundation, low risk)
2. **3.2** (quick win, rect logic)
3. **2.1** (stroke styles — high user-visible value)
4. **2.2** (gradient transforms)
5. **2.3** (`<use>` — expands supported SVG subset)
6. **2.4** (tspan — text fidelity)
7. **3.1** (sanitizer hardening)
8. **4.1 → 4.2 → 4.3** (polish)

---

## Testing Strategy
- Run `pytest tests/test_svg_*.py tests/test_e2e_svg_integration.py -q` after each phase.
- Add new tests alongside each fix (no retroactive coverage gaps).
- Final full suite: `pytest tests/ -q --ignore=tests/test_group_audit.py --ignore=tests/test_image_fetcher.py --ignore=tests/test_pptx_capabilities.py --ignore=tests/test_xml_extraction.py --ignore=tests/test_analyze_template.py`.

---

## Out of Scope
- Full SVG 1.1 spec support (filters, masks, clip-paths beyond basic polygon) — already documented as unsupported; users should fall back to raster.
- True curved text along paths — would require `freeform_builder` integration beyond current scope.
