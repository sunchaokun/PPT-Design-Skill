# Shape Functions Reference

Complete reference for shape creation and image cropping functions in build_helpers.

Import: `from ppt_pro_max.build_helpers import *`

## Universal Shape Function

### `shape()` — Any MSO_SHAPE in One Call

```python
shape(slide, shape_type, left, top, width, height, fill, line=None, C=None)
```

- `shape_type`: `MSO_SHAPE` enum value **or** string name (case-insensitive)
- `fill`: hex color (`'#1D78FA'`) or C dict role name (`'primary'`)
- `line`: hex color or C dict role name; `None` = no line (default)

**Examples:**

```python
# Using MSO_SHAPE enum
from pptx.enum.shapes import MSO_SHAPE
shape(slide, MSO_SHAPE.HEXAGON, 2, 2, 3, 2.6, '#1D78FA')

# Using string name (case-insensitive)
shape(slide, 'hexagon', 2, 2, 3, 2.6, '#1D78FA')
shape(slide, 'STAR_5_POINT', 5, 3, 2, 2, '#FF5500')
shape(slide, 'DONUT', 8, 4, 3, 3, 'primary', C=C)

# With border
shape(slide, MSO_SHAPE.HEART, 10, 2, 2, 2, '#FF3366', line='#CC0044')
```

## Centered Shape Functions

These use center-based positioning (`cx`, `cy` + `size`). Best for symmetrical shapes.

| Function | Shape | Usage Example |
|----------|-------|---------------|
| `hexagon(slide, cx, cy, size, fill, line, C)` | Hexagon | `hexagon(s, 5, 3.5, 2, '#1D78FA')` |
| `pentagon(slide, cx, cy, size, fill, line, C)` | Pentagon | `pentagon(s, 5, 3.5, 2, 'primary', C=C)` |
| `octagon(slide, cx, cy, size, fill, line, C)` | Octagon | `octagon(s, 5, 3.5, 2, '#1D78FA')` |
| `diamond(slide, cx, cy, size, fill, line, C)` | Diamond | `diamond(s, 5, 3.5, 2, '#FF5500', line='#CC4400')` |
| `star5(slide, cx, cy, size, fill, line, C)` | 5-point star | `star5(s, 5, 3.5, 2, '#FFD700')` |
| `star6(slide, cx, cy, size, fill, line, C)` | 6-point star | `star6(s, 5, 3.5, 2, '#1D78FA')` |
| `star8(slide, cx, cy, size, fill, line, C)` | 8-point star | `star8(s, 5, 3.5, 2, '#1D78FA')` |
| `star10(slide, cx, cy, size, fill, line, C)` | 10-point star | `star10(s, 5, 3.5, 2, '#1D78FA')` |
| `star12(slide, cx, cy, size, fill, line, C)` | 12-point star | `star12(s, 5, 3.5, 2, '#1D78FA')` |
| `donut(slide, cx, cy, size, fill, line, C)` | Donut/ring | `donut(s, 5, 3.5, 2, '#1D78FA')` |
| `heart(slide, cx, cy, size, fill, line, C)` | Heart | `heart(s, 5, 3.5, 2, '#FF3366')` |
| `cross(slide, cx, cy, size, fill, line, C)` | Cross/plus | `cross(s, 5, 3.5, 2, '#1D78FA')` |
| `moon(slide, cx, cy, size, fill, line, C)` | Moon | `moon(s, 5, 3.5, 2, '#FFD700')` |
| `sun(slide, cx, cy, size, fill, line, C)` | Sun | `sun(s, 5, 3.5, 2, '#FFD700')` |
| `block_arc(slide, cx, cy, size, fill, line, C)` | Block arc | `block_arc(s, 5, 3.5, 2, '#1D78FA')` |
| `gear(slide, cx, cy, size, fill, line, C, teeth=6)` | Gear | `gear(s, 5, 3.5, 2, '#1D78FA', teeth=9)` |
| `tear(slide, cx, cy, size, fill, line, C)` | Teardrop | `tear(s, 5, 3.5, 2, '#1D78FA')` |
| `math_plus(slide, cx, cy, size, fill, line, C)` | Plus sign | `math_plus(s, 5, 3.5, 2, '#1D78FA')` |
| `math_multiply(slide, cx, cy, size, fill, line, C)` | Multiply sign | `math_multiply(s, 5, 3.5, 2, '#1D78FA')` |
| `no_symbol(slide, cx, cy, size, fill, line, C)` | No/prohibited | `no_symbol(s, 5, 3.5, 2, '#FF0000')` |
| `flow_decision(slide, cx, cy, size, fill, line, C)` | Decision diamond | `flow_decision(s, 5, 3.5, 2, '#1D78FA')` |
| `flow_connector(slide, cx, cy, size, fill, line, C)` | Connector circle | `flow_connector(s, 5, 3.5, 2, '#1D78FA')` |

## Corner-Based Shape Functions

These use corner-based positioning (`left`, `top` + `width`, `height`).

| Function | Shape | Usage Example |
|----------|-------|---------------|
| `triangle(slide, left, top, width, height, fill, line, C)` | Triangle | `triangle(s, 2, 2, 3, 2, '#1D78FA')` |
| `right_triangle(slide, left, top, width, height, fill, line, C)` | Right triangle | `right_triangle(s, 2, 2, 3, 2, '#1D78FA')` |
| `parallelogram(slide, left, top, width, height, fill, line, C)` | Parallelogram | `parallelogram(s, 2, 2, 4, 2, '#1D78FA')` |
| `trapezoid(slide, left, top, width, height, fill, line, C)` | Trapezoid | `trapezoid(s, 2, 2, 4, 2, '#1D78FA')` |
| `arrow(slide, left, top, width, height, fill, line, C)` | Right arrow | `arrow(s, 2, 2, 4, 2, '#1D78FA')` |
| `chevron(slide, left, top, width, height, fill, line, C)` | Chevron | `chevron(s, 2, 2, 4, 2, '#1D78FA')` |
| `cloud(slide, left, top, width, height, fill, line, C)` | Cloud | `cloud(s, 2, 2, 4, 3, '#1D78FA')` |
| `lightning(slide, left, top, width, height, fill, line, C)` | Lightning bolt | `lightning(s, 2, 2, 2, 3, '#FFD700')` |
| `funnel(slide, left, top, width, height, fill, line, C)` | Funnel | `funnel(s, 2, 2, 3, 4, '#1D78FA')` |
| `wave(slide, left, top, width, height, fill, line, C)` | Wave | `wave(s, 2, 2, 4, 2, '#1D78FA')` |
| `plaque(slide, left, top, width, height, fill, line, C)` | Plaque | `plaque(s, 2, 2, 4, 3, '#1D78FA')` |
| `frame(slide, left, top, width, height, fill, line, C)` | Frame | `frame(s, 2, 2, 4, 3, '#1D78FA')` |
| `cube(slide, left, top, width, height, fill, line, C)` | 3D cube | `cube(s, 2, 2, 4, 3, '#1D78FA')` |
| `bevel(slide, left, top, width, height, fill, line, C)` | Bevel | `bevel(s, 2, 2, 4, 3, '#1D78FA')` |
| `folded_corner(slide, left, top, width, height, fill, line, C)` | Folded corner | `folded_corner(s, 2, 2, 4, 3, '#1D78FA')` |
| `flow_process(slide, left, top, width, height, fill, line, C)` | Process box | `flow_process(s, 2, 2, 4, 2, '#1D78FA')` |
| `flow_data(slide, left, top, width, height, fill, line, C)` | Data parallelogram | `flow_data(s, 2, 2, 4, 2, '#1D78FA')` |
| `flow_document(slide, left, top, width, height, fill, line, C)` | Document | `flow_document(s, 2, 2, 4, 2, '#1D78FA')` |

## Callout Function

```python
callout(slide, left, top, width, height, fill, line=None, C=None, style='rect')
```

- `style`: `'rect'` (default) | `'round'` | `'oval'` | `'cloud'`

**Examples:**

```python
callout(s, 2, 2, 4, 3, '#1D78FA')                    # Rectangular callout
callout(s, 2, 2, 4, 3, '#1D78FA', style='round')     # Rounded callout
callout(s, 2, 2, 4, 3, '#1D78FA', style='cloud')     # Cloud thought bubble
callout(s, 2, 2, 4, 3, 'accent', line='primary', C=C) # With C dict colors
```

## Image Cropping Functions

Clip images into various shapes.

| Function | Shape | Usage Example |
|----------|-------|---------------|
| `circle_image(slide, cx, cy, radius, image_path, border_color)` | Circle | `circle_image(s, 5, 3.5, 1, 'photo.jpg')` |
| `hex_image(slide, cx, cy, size, image_path, border_color)` | Hexagon | `hex_image(s, 5, 3.5, 2, 'photo.jpg')` |
| `star_image(slide, cx, cy, size, image_path, points=5, border_color)` | Star | `star_image(s, 5, 3.5, 2, 'photo.jpg', points=6)` |
| `diamond_image(slide, cx, cy, size, image_path, border_color)` | Diamond | `diamond_image(s, 5, 3.5, 2, 'photo.jpg')` |
| `heart_image(slide, cx, cy, size, image_path, border_color)` | Heart | `heart_image(s, 5, 3.5, 2, 'photo.jpg')` |
| `shape_image(slide, shape_type, left, top, width, height, image_path, border_color)` | Any shape | `shape_image(s, MSO_SHAPE.OCTAGON, 2, 2, 3, 3, 'photo.jpg')` |

- `star_image()` `points`: 5 (default), 6, 8, 10, 12
- `shape_image()` `shape_type`: MSO_SHAPE enum or string name (same as `shape()`)

**Examples:**

```python
# Team member photos in hexagons
for i, member in enumerate(team):
    hex_image(s, 3 + i * 2.5, 4, 1.8, member['photo'])

# Star-shaped award badge
star_image(s, 5, 3.5, 3, 'award.png', points=5, border_color='#FFD700')

# Heart avatar
heart_image(s, 5, 3.5, 2, 'avatar.jpg', border_color='#FF3366')

# Any shape image
shape_image(s, 'OCTAGON', 2, 2, 3, 3, 'photo.jpg')
shape_image(s, MSO_SHAPE.DECAGON, 2, 2, 3, 3, 'photo.jpg', border_color='#333')
```

## Complete Design Examples

### Example 1: Feature Cards with Hexagon Icons

```python
from ppt_pro_max.build_helpers import *

C = {'primary': '#1D78FA', 'accent': '#FF5500', 'white': '#FFFFFF',
     'card_bg': '#F1F5F9', 'text_dark': '#1A1A1A', 'text_muted': '#666666'}
t = TYPOGRAPHY['mckinsey']

prs = Presentation()
s = add_slide(prs)
page_header(s, 'Core Features', 'What makes us different', C=C)

features = [
    ('AI Engine', 'Real-time ML inference', 'engine.png'),
    ('Live Dashboard', 'Instant metrics overview', 'dashboard.png'),
    ('Smart Integration', 'Connect 500+ tools', 'integration.png'),
]

for i, (title, desc, img) in enumerate(features):
    x = 1.0 + i * 4.0
    rrect(s, x, 1.8, 3.5, 4.0, C['card_bg'], line=C.get('light', '#DDDDDD'), C=C)
    hex_image(s, x + 1.75, 3.0, 1.5, img)
    text(s, x + 0.3, 4.2, 2.9, 0.4, title, font_size=t.h3,
         color=C['text_dark'], bold=True, C=C)
    text(s, x + 0.3, 4.7, 2.9, 0.8, desc, font_size=t.caption,
         color=C['text_muted'], C=C)

prs.save('features.pptx')
```

### Example 2: Process Flow with Chevrons and Flowchart Shapes

```python
prs = Presentation()
s = add_slide(prs)

steps = ['Input', 'Process', 'Decision', 'Output']
for i, step in enumerate(steps):
    x = 1.0 + i * 3.0
    if i < 3:
        chevron(s, x, 3.0, 2.8, 1.5, C['primary'], C=C)
    else:
        arrow(s, x, 3.0, 2.8, 1.5, C['accent'], C=C)
    text(s, x + 0.3, 3.3, 2.2, 0.5, step, font_size=14,
         color='#FFFFFF', bold=True, align='center')

flow_decision(s, 10, 5.5, 1.5, C['accent'], C=C)
text(s, 9.5, 5.3, 1.0, 0.3, 'OK?', font_size=12, color='#FFFFFF', bold=True, C=C)
```

### Example 3: Dashboard with Donut Charts and Star Ratings

```python
prs = Presentation()
s = add_slide(prs)

# KPI donuts
for i, (label, pct, color) in enumerate([('Revenue', '85%', '#1D78FA'),
                                           ('Retention', '97%', '#00B050'),
                                           ('Growth', '4.2x', '#FF5500')]):
    x = 1.5 + i * 4.0
    donut(s, x, 3.0, 2.0, color, C=C)
    text(s, x - 0.5, 2.6, 1.0, 0.5, pct, font_size=t.h2,
         color=color, bold=True, align='center', C=C)
    text(s, x - 0.8, 4.3, 1.6, 0.3, label, font_size=t.caption,
         color=C['text_muted'], align='center', C=C)

# Star rating
for i in range(5):
    fill = '#FFD700' if i < 4 else '#E0E0E0'
    star5(s, 9.0 + i * 0.7, 5.5, 0.6, fill)
```

### Example 4: Creative Slide with Mixed Shapes

```python
prs = Presentation()
s = add_slide(prs)

# Background accent shapes
lightning(s, 0, 0, 3, 4, '#1D78FA', C=C)
wave(s, 8, 5, 5, 2.5, '#FF5500', C=C)

# Content
heart(s, 6.5, 2.5, 3, '#FF3366', C=C)
text(s, 5, 2, 3, 1, 'Love It', font_size=36, color='#FFFFFF',
     bold=True, align='center', C=C)

# Callout bubble
callout(s, 3, 4, 4, 2, '#FFFFFF', line='#1D78FA', C=C, style='round')
text(s, 3.3, 4.3, 3.4, 1, 'Users love our product!', font_size=16,
     color=C['text_dark'], C=C)

# Moon and sun decoration
moon(s, 11, 1.5, 1.5, '#FFD700', C=C)
sun(s, 2, 6, 1.0, '#FFD700', C=C)
```

## MSO_SHAPE Quick Reference (for `shape()` and `shape_image()`)

Most useful shapes for PPT design:

| Category | Shapes |
|----------|--------|
| **Basic** | RECTANGLE, ROUNDED_RECTANGLE, OVAL, DIAMOND, ISOSCELES_TRIANGLE, RIGHT_TRIANGLE, PARALLELOGRAM, TRAPEZOID, HEXAGON, OCTAGON, PENTAGON, DONUT |
| **Stars** | STAR_4_POINT, STAR_5_POINT, STAR_6_POINT, STAR_7_POINT, STAR_8_POINT, STAR_10_POINT, STAR_12_POINT, STAR_16_POINT, STAR_24_POINT, STAR_32_POINT |
| **Arrows** | RIGHT_ARROW, LEFT_ARROW, UP_ARROW, DOWN_ARROW, LEFT_RIGHT_ARROW, UP_DOWN_ARROW, BENT_ARROW, CHEVRON, NOTCHED_RIGHT_ARROW, U_TURN_ARROW, CIRCULAR_ARROW, QUAD_ARROW, SWOOSH_ARROW |
| **Callouts** | RECTANGULAR_CALLOUT, ROUNDED_RECTANGULAR_CALLOUT, OVAL_CALLOUT, CLOUD_CALLOUT, LINE_CALLOUT_1..4 |
| **Flowchart** | FLOWCHART_PROCESS, FLOWCHART_DECISION, FLOWCHART_DATA, FLOWCHART_DOCUMENT, FLOWCHART_CONNECTOR, FLOWCHART_TERMINATOR, FLOWCHART_PREDEFINED_PROCESS |
| **Decorative** | HEART, LIGHTNING_BOLT, CLOUD, MOON, SUN, CROSS, FRAME, BEVEL, CUBE, WAVE, TEAR, FUNNEL, GEAR_6, GEAR_9, PLAQUE, FOLDED_CORNER, BLOCK_ARC, NO_SYMBOL |
| **Math** | MATH_PLUS, MATH_MINUS, MATH_MULTIPLY, MATH_DIVIDE, MATH_EQUAL, MATH_NOT_EQUAL |
| **Ribbons** | UP_RIBBON, DOWN_RIBBON, CURVED_UP_RIBBON, CURVED_DOWN_RIBBON |
| **Scrolls** | HORIZONTAL_SCROLL, VERTICAL_SCROLL |
| **Brackets** | LEFT_BRACE, RIGHT_BRACE, LEFT_BRACKET, RIGHT_BRACKET |
