# SVG Compiler Module Audit Report

## 1. Overview
The SVG Compiler module is a specialized pipeline designed to transform SVG 1.1 subsets into native, editable PPTX shapes. It consists of 10 sub-modules handling sanitization, path parsing, affine transformations, paint resolution, text rendering, and theme integration.

- **Working Directory**: `src/ppt_pro_max/renderer/svg_compiler/`
- **Core Entry**: `SVGCompiler.compile(svg_text, slide, rect, C)`
- **Test Coverage**: 270 passed tests across 11 test files.

---

## 2. Critical Issues (Functional/Type Risks)

### 2.1 Type Contract Violation in `_resolve_svg_color`
- **Location**: `_compiler.py:108-109`
- **Issue**: The function signature `-> str` is violated when returning `None` for `"none"` or empty input.
- **Impact**: Downstream calls (e.g., `_add_freeform` or `_render_bool`) might receive `None` where a hex string is expected. While current usage often has `or "#FFFFFF"` fallbacks, this inconsistency leads to brittle code and potential rendering of shapes with missing fills when transparency was intended.
- **Recommendation**: Return a sentinel string (e.g., `"none"`) or fix all downstream handlers to explicitly handle `Optional[str]`.

### 2.2 Redundant Constant Definition: `_BASELINE_MAP`
- **Location**: `_text.py:42-51` and `_text.py:94-103`
- **Issue**: `_BASELINE_MAP` is defined twice with identical content.
- **Impact**: Code smell and maintenance risk. Future updates to baseline mapping might only change one instance, leading to inconsistent behavior.
- **Recommendation**: Remove the first definition (lines 42-51).

---

## 3. Moderate Issues (Rendering & Feature Gaps)

### 3.1 Unused `stroke-dasharray` and Line Style Implementation
- **Location**: `_dash.py` vs `_compiler.py`
- **Issue**: The functionality to parse and apply stroke styles (dashes, caps, joins) exists in `_dash.py` but is **never called** in `_compiler.py`.
- **Impact**: All SVG shapes are rendered with solid lines. Dashed lines, rounded caps, and specific joins in the SVG source are ignored.
- **Recommendation**: Update `_compiler.py:_render_shape` to call `parse_stroke_style(el)` and `apply_stroke_style(elem, style)`.

### 3.2 Incomplete `tspan` Layout Strategy
- **Location**: `_text.py:409-466`
- **Issue**: 
  - `dx`/`dy` offsets are parsed but never used for positioning.
  - Absolute `x`/`y` on `tspan` trigger a new paragraph but don't set the absolute position; instead, they rely on a fixed `line_h` increment.
- **Impact**: Multi-line text with non-uniform spacing (like labels in complex diagrams) will be misaligned compared to the original SVG design.
- **Recommendation**: Use `dx`/`dy` as coordinate offsets and prioritize `tspan.x/y` for paragraph positioning if present.

### 3.3 Missing `gradientTransform` Support
- **Location**: `_paint.py:98-122`
- **Issue**: `GradientDef` captures the `transform` attribute, but `apply_gradient` ignores it.
- **Impact**: Rotated or skewed gradients in SVG will appear as default horizontal/vertical linear gradients in PPT.
- **Recommendation**: Multiply the gradient coordinates (x1, y1, x2, y2) by the `grad.transform` matrix before mapping to PPT angles.

### 3.4 Unsupported `<use>` Element
- **Location**: `_compiler.py:422-435`
- **Issue**: The compiler does not handle `<use>` tags, which are common for re-using symbols or shapes defined in `<defs>`.
- **Impact**: LLM-generated SVGs that use a "library" approach (common in icon sets or complex charts) will fail to render the referenced components.
- **Recommendation**: Implement a lookup for the referenced ID and recursively walk the target element.

### 3.5 Sanitizer Regex Risk
- **Location**: `_sanitizer.py:59-68`
- **Issue**: `_fix_self_closing` uses regex to force-close tags like `<rect>`. If the tag contains children (e.g., `<rect><title>...</title></rect>`), the regex will break the XML structure.
- **Impact**: Loss of metadata or invalid XML if non-standard SVG is provided.
- **Recommendation**: Use `lxml` to handle tag closing instead of raw regex, or narrow the regex to only match tags without children.

---

## 4. Minor Issues (Performance/Optimization)

### 4.1 O(n²) Overlap Detection
- **Location**: `_compiler.py:660-703`
- **Issue**: Collision detection for text boxes uses nested loops comparing every pair of shapes.
- **Impact**: Performance degradation in SVGs with 100+ text elements (uncommon but possible in data-dense charts).
- **Recommendation**: Use a simple spatial grid if shape count exceeds a threshold.

### 4.2 Hardcoded 96 DPI for Text Measurement
- **Location**: `_text.py:187`
- **Issue**: Uses `px_per_inch = 96.0` for Pillow-based measurement.
- **Impact**: Slight sizing errors depending on the host OS display settings; PPT technically uses 72 DPI internally.
- **Recommendation**: Standardize on 72 DPI or 96 DPI across all modules.

### 4.3 `min` Scaling Assumption
- **Location**: `_compiler.py:236`
- **Issue**: Always uses `min(w/vw, h/vh)` for "contain" behavior.
- **Impact**: No option for "cover" or "stretch" scaling.
- **Recommendation**: Add a `scaling_mode` parameter to `compile()`.

### 4.4 Incomplete `rx`/`ry` Logic for Native Rects
- **Location**: `_compiler.py:493`
- **Issue**: Only checks for `rx`. If `ry` is present without `rx`, the SVG spec says `rx` should equal `ry`, but the code will use a native (non-rounded) rectangle.
- **Impact**: Incorrect rendering of rounded rectangles where only `ry` is defined.

---

## 5. Conclusion
The SVG Compiler is robust for standard LLM-generated diagrams (pyramids, venns, funnels) but requires refinement in **line style application**, **gradient transforms**, and **advanced text layout** to achieve professional-grade design fidelity. The critical type inconsistencies in color resolution should be addressed immediately to prevent runtime errors in edge cases.
