# build_helpers API Quick Reference

Import: `from ppt_pro_max.build_helpers import *`

Full SKILL.md: [`../SKILL.md`](../SKILL.md)

---

## Setup (every build.py starts here)

```python
from ppt_pro_max.build_helpers import *

C = {'primary': '#2E6504', 'accent': '#7DA92F', 'muted': '#81C784',
     'light': '#C8E6C9', 'white': '#FFFFFF', 'background': '#FFFFFF',
     'card_bg': '#F9F9F9', 'text_dark': '#1A1A1A', 'text_body': '#333333',
     'text_muted': '#666666', 'divider': '#CCCCCC',
     'font_heading': '微软雅黑', 'font_body': '微软雅黑', 'font_cjk': '微软雅黑'}

t = TYPOGRAPHY['mckinsey']
sp = SPACING['mckinsey']
prs = Presentation()
s = add_slide(prs)
```

---

## Classes

### Typography

```python
t = TYPOGRAPHY['mckinsey']  # access: t.hero, t.h1, t.h2, t.h3, t.body, t.caption, t.micro
```

| Key | hero | h1 | h2 | h3 | body | caption | micro |
|-----|------|----|----|----|----|---------|-------|
| `'mckinsey'` | 44 | 28 | 20 | 16 | 12 | 10 | 8 |
| `'cyberpunk'` | 48 | 28 | 18 | 14 | 11 | 9 | 7 |
| `'creative'` | 44 | 28 | 22 | 18 | 13 | 11 | 9 |
| `'professional'` | 44 | 28 | 20 | 16 | 12 | 10 | 8 |
| `'minimal'` | 40 | 24 | 18 | 14 | 11 | 9 | 7 |
| `'cjk_mckinsey'` | 44 | 30 | 22 | 18 | 14 | 12 | 10 |
| `'cjk_professional'` | 44 | 30 | 22 | 18 | 14 | 12 | 10 |
| `'cjk_creative'` | 44 | 30 | 24 | 20 | 15 | 13 | 11 |

### Spacing

```python
sp = SPACING['mckinsey']  # access: sp.page_margin, sp.section_gap, sp.card_gap, sp.card_padding, sp.line_height, sp.bar_gap
```

| Key | page_margin | section_gap | card_gap | card_padding | line_height | bar_gap |
|-----|-------------|-------------|----------|--------------|-------------|---------|
| `'mckinsey'` | 0.65 | 0.5 | 0.35 | 0.2 | 1.4 | 0.2 |
| `'cyberpunk'` | 0.8 | 0.6 | 0.4 | 0.25 | 1.3 | 0.25 |
| `'creative'` | 0.8 | 0.6 | 0.4 | 0.25 | 1.5 | 0.25 |
| `'professional'` | 0.65 | 0.5 | 0.35 | 0.2 | 1.4 | 0.2 |
| `'minimal'` | 1.0 | 0.6 | 0.5 | 0.3 | 1.5 | 0.3 |

---

## Page Structure

| Function | Signature | Purpose |
|----------|-----------|---------|
| `add_slide` | `add_slide(prs, layout_index=None)` | Add blank slide |
| `hero_slide` | `hero_slide(slide, title, subtitle='', C=None, typo=None, grouped=True)` | Cover/hero page |
| `cta_slide` | `cta_slide(slide, title, subtitle='', C=None, typo=None, grouped=True)` | Call-to-action page |
| `section_divider` | `section_divider(slide, number, title, C=None, typo=None, grouped=True)` | Section divider |
| `page_header` | `page_header(slide, title, subtitle='', C=None, left=0.65, width=None, typo=None, spacing=None)` | Title + subtitle + divider line |

---

## Text & Code

| Function | Signature | Purpose |
|----------|-----------|---------|
| `text` | `text(slide, left, top, width, height, txt, font_size=12, color='text_body', bold=False, align='left', font_name=None, C=None, anchor='top')` | Single-line text |
| `multiline` | `multiline(slide, left, top, width, height, lines, font_size=12, color='text_body', bold=False, align='left', font_name=None, C=None, line_spacing=None)` | Multi-line text |
| `gradient_text` | `gradient_text(slide, left, top, width, height, txt, preset='gold-shine', stops=None, font_size=44, bold=False, font_name=None, cjk_font=None, align='left')` | Gradient-filled text |
| `vertical_text` | `vertical_text(slide, left, top, width, height, txt, direction='ea', font_name='STKaiti', font_size=24, color='#000000', bold=False, align='center')` | Vertical text |
| `code_block` | `code_block(slide, left, top, width, height, lines, language='python', C=None, typo=None, grouped=True)` | Code block with language badge |
| `text_outline` | `text_outline(slide, left, top, width, height, txt, color='#FFFFFF', width_pt=1.5, font_size=44, bold=False, font_name=None, C=None, align='left')` | Outlined text |
| `text_shadow` | `text_shadow(slide, left, top, width, height, txt, blur_pt=8, distance_pt=3, direction_deg=90, color='#000000', alpha_pct=25, font_size=44, bold=False, font_name=None, C=None, align='left')` | Shadowed text |
| `text_glow` | `text_glow(slide, left, top, width, height, txt, color='#00FFFF', size_pt=8, alpha_pct=40, font_size=44, bold=False, font_name=None, C=None, align='left')` | Glowing text |

### Gradient Presets

`gold-shine`, `blue-deep`, `purple-neon`, `ink-wash`, `cyber-cyan`, `sunset`, `emerald`, `rose-gold`, `seal-red`, `steel`

---

## Shapes

### Basic

| Function | Signature | Purpose |
|----------|-----------|---------|
| `rect` | `rect(slide, left, top, width, height, fill, line=None, C=None)` | Rectangle |
| `rrect` | `rrect(slide, left, top, width, height, fill, line=None, C=None)` | Rounded rectangle |
| `oval` | `oval(slide, left, top, width, height, fill, line=None, C=None)` | Oval/circle |
| `shape` | `shape(slide, shape_type, left, top, width, height, fill, line=None, C=None)` | Any MSO_SHAPE (string or enum) |

### Polygons

| Function | Center-based | Purpose |
|----------|-------------|---------|
| `hexagon` | `hexagon(slide, cx, cy, size, fill, line=None, C=None)` | Hexagon |
| `pentagon` | `pentagon(slide, cx, cy, size, fill, line=None, C=None)` | Pentagon |
| `octagon` | `octagon(slide, cx, cy, size, fill, line=None, C=None)` | Octagon |
| `diamond` | `diamond(slide, cx, cy, size, fill, line=None, C=None)` | Diamond |

### Stars

| Function | Center-based | Purpose |
|----------|-------------|---------|
| `star5` | `star5(slide, cx, cy, size, fill, line=None, C=None)` | 5-point star |
| `star6` | `star6(slide, cx, cy, size, fill, line=None, C=None)` | 6-point star |
| `star8` | `star8(slide, cx, cy, size, fill, line=None, C=None)` | 8-point star |
| `star10` | `star10(slide, cx, cy, size, fill, line=None, C=None)` | 10-point star |
| `star12` | `star12(slide, cx, cy, size, fill, line=None, C=None)` | 12-point star |

### Special

| Function | Center-based | Purpose |
|----------|-------------|---------|
| `donut` | `donut(slide, cx, cy, size, fill, line=None, C=None)` | Donut ring |
| `heart` | `heart(slide, cx, cy, size, fill, line=None, C=None)` | Heart |
| `cross` | `cross(slide, cx, cy, size, fill, line=None, C=None)` | Cross |
| `moon` | `moon(slide, cx, cy, size, fill, line=None, C=None)` | Moon |
| `sun` | `sun(slide, cx, cy, size, fill, line=None, C=None)` | Sun |
| `block_arc` | `block_arc(slide, cx, cy, size, fill, line=None, C=None)` | Block arc |
| `gear` | `gear(slide, cx, cy, size, fill, line=None, C=None, teeth=6)` | Gear (6 or 9 teeth) |
| `tear` | `tear(slide, cx, cy, size, fill, line=None, C=None)` | Teardrop |

### Directional

| Function | Corner-based | Purpose |
|----------|-------------|---------|
| `arrow` | `arrow(slide, left, top, width, height, fill, line=None, C=None)` | Right arrow |
| `chevron` | `chevron(slide, left, top, width, height, fill, line=None, C=None)` | Chevron |
| `cloud` | `cloud(slide, left, top, width, height, fill, line=None, C=None)` | Cloud |
| `lightning` | `lightning(slide, left, top, width, height, fill, line=None, C=None)` | Lightning bolt |
| `funnel` | `funnel(slide, left, top, width, height, fill, line=None, C=None)` | Funnel |
| `wave` | `wave(slide, left, top, width, height, fill, line=None, C=None)` | Wave |

### Flowchart & Callout

| Function | Purpose |
|----------|---------|
| `callout(slide, left, top, width, height, fill, line=None, C=None, style='rect')` | Callout (style: rect/round/oval/cloud) |
| `flow_process(slide, left, top, width, height, ...)` | Process box |
| `flow_decision(slide, cx, cy, size, ...)` | Decision diamond |
| `flow_data(slide, left, top, width, height, ...)` | Data parallelogram |
| `flow_document(slide, left, top, width, height, ...)` | Document shape |
| `flow_connector(slide, cx, cy, size, ...)` | Connector circle |

### Extra

| Function | Purpose |
|----------|---------|
| `no_symbol(slide, cx, cy, size, ...)` | No/prohibition symbol |
| `plaque(slide, left, top, width, height, ...)` | Plaque |
| `frame(slide, left, top, width, height, ...)` | Frame |
| `cube(slide, left, top, width, height, ...)` | 3D cube |
| `bevel(slide, left, top, width, height, ...)` | Bevel |
| `folded_corner(slide, left, top, width, height, ...)` | Folded corner |
| `math_plus(slide, cx, cy, size, ...)` | Plus sign |
| `math_multiply(slide, cx, cy, size, ...)` | Multiply sign |

### shape() String Names

| Category | Names |
|----------|-------|
| Polygons | HEXAGON, PENTAGON, OCTAGON, DIAMOND, DECAGON, DODECAGON |
| Stars | STAR_4_POINT, STAR_5_POINT, STAR_6_POINT, STAR_8_POINT, STAR_10_POINT, STAR_12_POINT |
| Arrows | RIGHT_ARROW, LEFT_ARROW, UP_ARROW, DOWN_ARROW, BENT_ARROW, CHEVRON, NOTCHED_RIGHT_ARROW, U_TURN_ARROW, CIRCULAR_ARROW, QUAD_ARROW |
| Flowchart | FLOWCHART_PROCESS, FLOWCHART_DECISION, FLOWCHART_DATA, FLOWCHART_DOCUMENT, FLOWCHART_CONNECTOR, FLOWCHART_TERMINATOR |
| Callouts | RECTANGULAR_CALLOUT, ROUNDED_RECTANGULAR_CALLOUT, OVAL_CALLOUT, CLOUD_CALLOUT |
| Special | HEART, LIGHTNING_BOLT, CLOUD, MOON, SUN, CROSS, DONUT, FRAME, BEVEL, CUBE, WAVE, TEAR, FUNNEL, GEAR_6, GEAR_9, PLAQUE, FOLDED_CORNER, BLOCK_ARC, NO_SYMBOL |
| Math | MATH_PLUS, MATH_MINUS, MATH_MULTIPLY, MATH_DIVIDE, MATH_EQUAL |
| Ribbons | UP_RIBBON, DOWN_RIBBON, CURVED_UP_RIBBON, CURVED_DOWN_RIBBON |

---

## Boolean Shapes (require shapely)

| Function | Signature | Purpose |
|----------|-----------|---------|
| `spotlight` | `spotlight(slide, cx, cy, radius, alpha=70, color='#000000')` | Dark overlay with bright circular window |
| `bool_donut` | `bool_donut(slide, cx, cy, outer_r, inner_r, fill='#1D78FA', line=None, C=None)` | Custom donut with off-center hole |
| `bool_frame` | `bool_frame(slide, x, y, w, h, border, fill=None, line=None, C=None)` | Frame/border (outer minus inner) |
| `bool_clipped_card` | `bool_clipped_card(slide, x, y, w, h, clip_corners, clip_size=0.3, fill=None, line=None, C=None)` | Card with clipped corners |
| `bool_neon_tube` | `bool_neon_tube(slide, x, y, w, h, wall=0.06, fill=None, C=None)` | Hollow neon tube |
| `bool_star` | `bool_star(slide, cx, cy, r, points=5, inner_ratio=0.4, fill=None, line=None, C=None)` | Custom star with adjustable inner radius |
| `bool_cross` | `bool_cross(slide, cx, cy, w, h, bar_ratio=0.33, fill=None, line=None, C=None)` | Custom cross with adjustable bar thickness |

---

## Data & Charts

| Function | Signature | Purpose |
|----------|-----------|---------|
| `kpi_card` | `kpi_card(slide, left, top, width, height, number, label, trend='', trend_up=True, C=None, typo=None, grouped=True)` | Single metric highlight |
| `bar_chart` | `bar_chart(slide, left, top, data, max_width=5.0, bar_height=0.3, C=None, typo=None, spacing=None, grouped=True)` | Horizontal progress bars |
| `comparison_bars` | `comparison_bars(slide, left, top, metrics, max_width=4.0, C=None, typo=None, spacing=None, grouped=True)` | Before/after comparison |
| `donut_chart` | `donut_chart(slide, cx, cy, radius, inner_radius, sectors, C=None, typo=None, grouped=True, native=True)` | Donut/pie chart |
| `native_chart` | `native_chart(slide, left, top, width, height, chart_type, categories=None, series=None, style=None, C=None)` | Native PowerPoint chart |
| `highlight_cards` | `highlight_cards(slide, left, top, cards, total_width=12.0, C=None, typo=None, spacing=None, grouped=True)` | Multi-metric card row |

### native_chart Types

| Category | Types |
|----------|-------|
| Column | `bar`, `bar_stacked`, `bar_100`, `bar_3d` |
| Bar (horizontal) | `bar_horizontal`, `bar_horizontal_stacked`, `bar_horizontal_100` |
| Line | `line`, `line_markers`, `line_stacked`, `line_stacked_100` |
| Pie | `pie`, `pie_3d`, `pie_exploded` |
| Doughnut | `doughnut`, `doughnut_exploded` |
| Area | `area`, `area_stacked`, `area_stacked_100` |
| Scatter | `scatter`, `scatter_lines`, `scatter_smooth` |
| Radar | `radar`, `radar_markers` |
| Bubble | `bubble` |
| Stock | `stock_hlc`, `stock_ohlc` |

### native_chart style dict

| Key | Default | Description |
|-----|---------|-------------|
| `show_legend` | `True` | Show/hide legend |
| `legend_position` | `'bottom'` | `'bottom'`/`'top'`/`'left'`/`'right'` |
| `show_labels` | `False` | Show data labels |
| `show_value` | `True` | Show numeric value |
| `show_percentage` | `False` | Show percentage |
| `show_category_name` | `False` | Show category name |
| `label_font_size` | `9` | Data label font size (pt) |
| `label_position` | `'outside_end'` | `'center'`/`'inside_end'`/`'outside_end'`/`'best_fit'` |
| `number_format` | — | e.g. `'#,##0'`, `'0.0%'` |
| `color_scheme` | `'brand'` | `'brand'`/`'auto'`/`['#hex', ...]` |
| `title` | — | Chart title |
| `value_axis_title` | — | Y-axis title |
| `category_axis_title` | — | X-axis title |
| `gridlines` | `'major_y'` | `'none'`/`'major_y'`/`'major_x'`/`'major_xy'` |
| `tick_number_format` | — | Axis tick format |
| `chart_style` | — | 1-48 built-in style |

---

## Image Effects

| Function | Signature | Purpose |
|----------|-----------|---------|
| `cover_image` | `cover_image(slide, left, top, width, height, image_path)` | Cover-fit image (Pillow pre-crop) |
| `circle_image` | `circle_image(slide, cx, cy, radius, image_path, border_color=None)` | Circle-cropped image |
| `hex_image` | `hex_image(slide, cx, cy, size, image_path, border_color=None)` | Hexagon-cropped image |
| `star_image` | `star_image(slide, cx, cy, size, image_path, points=5, border_color=None)` | Star-cropped image |
| `diamond_image` | `diamond_image(slide, cx, cy, size, image_path, border_color=None)` | Diamond-cropped image |
| `heart_image` | `heart_image(slide, cx, cy, size, image_path, border_color=None)` | Heart-cropped image |
| `shape_image` | `shape_image(slide, shape_type, left, top, width, height, image_path, border_color=None)` | Any shape image crop |
| `soft_edge_image` | `soft_edge_image(slide, left, top, width, height, image_path, soft_radius=10)` | Soft-edge faded image |
| `duotone_image` | `duotone_image(slide, left, top, width, height, image_path, color1='#0000FF', color2='#FF0000')` | Duotone image |
| `artistic_image` | `artistic_image(slide, left, top, width, height, image_path, effect='watercolor_sponge', params=None)` | 22 artistic effects |
| `adjust_image` | `adjust_image(shape, brightness=0, contrast=0, saturation=100)` | Adjust brightness/contrast/saturation |

### Artistic Effects

`watercolor_sponge`, `pencil_grayscale`, `pencil_colored`, `mosaic_bubbles`, `film_grain`, `glow_diffused`, `blur`, `cutout`, `marker`, `paint_strokes`, `texturizer`, `light_screen`, `line_drawing`, `etching`, `plastic`, `glass`, `cement`, `chalk_smokey`, `crayon`, `halftone`, `photocopy`, `stamp`

---

## Shape Effects

| Function | Signature | Purpose |
|----------|-----------|---------|
| `add_shadow` | `add_shadow(shape, blur_pt=8, distance_pt=3, direction_deg=90, color='#000000', alpha_pct=25)` | Add shadow to any shape |
| `add_glow` | `add_glow(shape, color='#00FFFF', size_pt=8, alpha_pct=40)` | Add glow to any shape |
| `shape_3d` | `shape_3d(slide, left, top, width, height, depth=10.0, material='powder', extrusion_color='#000000', shape_type=MSO_SHAPE.RECTANGLE)` | 3D extrusion |
| `bevel_shape` | `bevel_shape(slide, left, top, width, height, top_w=4.0, top_h=2.0, material='powder', shape_type=MSO_SHAPE.RECTANGLE)` | Bevel effect |
| `pattern_fill` | `pattern_fill(slide, left, top, width, height, pattern_type, fg_color, bg_color, fg_alpha=None, shape_type=MSO_SHAPE.RECTANGLE)` | Pattern fill |
| `frosted_panel` | `frosted_panel(slide, left, top, width, height, tint='#FFFFFF', alpha=50, soft_edge=8)` | Frosted glass |

### Pattern Types

`cross`, `dark_downward_diagonal`, `dark_upward_diagonal`, `dark_horizontal`, `dark_vertical`, `small_checker`, `trellis`, `light_horizontal`, `light_vertical`, `light_downward_diagonal`, `light_upward_diagonal`, `narrow_horizontal`, `narrow_vertical`, `dashed_downward_diagonal`, `dashed_upward_diagonal`, `dashed_horizontal`, `dashed_vertical`, `small_confetti`, `large_confetti`, `zigzag`, `wave`, `diagonal_brick`, `horizontal_brick`, `weave`, `plaid`, `divot`, `dotted_grid`, `dotted_diamond`, `shingle`, `large_checker`, `large_grid`, `small_grid`, `solid_diamond`, `percent_5` through `percent_90`

---

## Decorations

| Function | Signature | Purpose |
|----------|-----------|---------|
| `top_bar` | `top_bar(slide, color, width=13.333, height=0.08, C=None)` | Top accent bar |
| `brush_divider` | `brush_divider(slide, left, top, width, color='#2C2C2C', thickness=0.08)` | Brush-stroke divider |
| `seal_stamp` | `seal_stamp(slide, left, top, size, txt, fill_hex='#C41E3A', font_name='STZhongsong', rotation=-15, style='zhu', border_width_pt=4.0)` | Chinese seal stamp |
| `neon_border` | `neon_border(slide, left, top, width, height, color='#8B5CF6', radius=0.1)` | Neon glowing border |
| `glass_panel` | `glass_panel(slide, left, top, width, height, tint='#FFFFFF', alpha=50, soft_edge=8)` | Glassmorphism panel |
| `grid_background` | `grid_background(slide, spacing=1.0, color='#E0E0E0', alpha=15)` | Subtle grid background |
| `ink_splash` | `ink_splash(slide, left, top, size, color='#2C2C2C', alpha=100)` | Ink splash decoration |

---

## Animation

| Function | Signature | Purpose |
|----------|-----------|---------|
| `slide_transition` | `slide_transition(slide, transition_type='fade', speed='medium', advance_on_click=True, advance_after_ms=None)` | Slide transition |
| `entrance_animation` | `entrance_animation(slide, shape_id, effect='fade_in', delay_ms=0, duration_ms=500, click_triggered=True)` | Entrance animation |
| `exit_animation` | `exit_animation(slide, shape_id, effect='fade_out', delay_ms=0, duration_ms=500, click_triggered=True)` | Exit animation |
| `emphasis_animation` | `emphasis_animation(slide, shape_id, effect='pulse', delay_ms=0, duration_ms=500, click_triggered=True)` | Emphasis animation |

### Transition Types

`fade`, `push`, `wipe`, `split`, `reveal`, `cover`, `clock`, `wheel`, `random`, `dissolve`, `newsflash`, `blinds`

### Entrance Effects

`fade_in`, `fly_in`, `zoom_in`, `grow_turn`, `wheel`, `spiral`, `swivel`, `bounce`, `float`, `rise_up`, `ascend`

### Exit Presets

`fade_out`, `fly_out`, `zoom_out`, `shrink_turn`, `spiral_out`, `swivel_out`, `bounce_out`, `descend`

### Emphasis Presets

`pulse`, `grow`, `shrink`, `spin`, `teeter`, `color_pulse`, `darken`, `lighten`

---

## Component Library

| Function | Signature | Purpose |
|----------|-----------|---------|
| `query_components` | `query_components(component_type=None, category=None, node_count=None, limit=10)` | Search 5,715 chart templates |

`component_type` maps to the DB `category` (all 6 types live under `type='group'`; pass `component_type='smartart'` + `category` to search SmartArt diagrams). Returns a list of dicts: `{id, type, category, variant, node_count, level_count, tags, xml_path, source}`.

### Component Types

| Type | Count | Description |
|------|-------|-------------|
| `infographic` | 4,237 | Infographic charts |
| `process` | 673 | Process diagrams |
| `hierarchy` | 566 | Hierarchy trees |
| `chart` | 133 | Data charts |
| `timeline` | 41 | Timeline diagrams |
| `swot` | 39 | SWOT analysis |
| `smartart` | 23 | SmartArt (cycle/process/pyramid) |

---

## Template Analysis (VI Build)

| Function | Signature | Purpose |
|----------|-----------|---------|
| `analyze_pptx` | `analyze_pptx(pptx_path)` | Extract design DNA from any PPTX |

Returns a dict with: `slides`, `num_slides`, `color_palette`, `actual_colors`, `actual_fonts`, `actual_font_sizes`, `font_scheme`, `cjk_font_scheme`, `brand_spec`, `has_logo`, `decorative_groups_count`, `slide_width_emu`, `slide_height_emu`, `source_path`.

---

## Accessibility

| Function | Signature | Purpose |
|----------|-----------|---------|
| `check_contrast` | `check_contrast(color1, color2, min_ratio=3.0)` | WCAG contrast ratio check. Returns `(ratio, ok)` |
| `contrast_text` | `contrast_text(bg_color, min_ratio=4.5)` | Auto-select white or dark text. Returns `'#FFFFFF'` or `'#1A1A1A'` |

---

## Template (VI Build only)

| Function | Signature | Purpose |
|----------|-----------|---------|
| `copy_decorations` | `copy_decorations(slide, template_slide, skip_long_text=True, skip_image=True)` | Copy decorations from template |
| `copy_logo` | `copy_logo(slide, template_slide, color_hints=None)` | Copy logo from template |

---

## Chinese Character Writing Grids

| Function | Signature | Purpose |
|----------|-----------|---------|
| `mizi_grid` | `mizi_grid(slide, left, top, size, char=None, border_color='#4CAF50', guide_color='#A0A0A0', ...)` | 米字格 (cross + diagonal) |
| `tian_grid` | `tian_grid(slide, left, top, size, char=None, border_color='#4CAF50', guide_color='#A0A0A0', ...)` | 田字格 (cross only) |
| `pinyin_grid` | `pinyin_grid(slide, left, top, width, pinyin=None, baseline_y=None, line_spacing=0.3, ...)` | 四线格/拼音格 |
| `hanzi_row` | `hanzi_row(slide, left, top, size, chars, grid_type='mizi', gap=0.3, ...)` | Row of character grids |
| `pinyin_hanzi_block` | `pinyin_hanzi_block(slide, left, top, size, items, gap=0.3, grid_type='mizi', ...)` | Pinyin + character paired block |

---

## Color Resolution

- Hex value: `'#2E6504'` → used directly
- Role name: `'primary'` → looks up `C['primary']`
- Missing role: returns `'#000000'` (never crashes)
