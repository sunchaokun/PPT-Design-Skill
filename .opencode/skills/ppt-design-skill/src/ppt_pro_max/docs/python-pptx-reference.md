# python-pptx Capability Reference

Complete API reference for LLMs generating python-pptx code. Covers all shape types, formatting, charts, tables, media, and OOXML patterns.

## Quick Start

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
```

---

## 1. Shape Types (MSO_SHAPE — 170+ shapes)

### Basic Shapes
| Enum | Visual |
|------|--------|
| `RECTANGLE` | ▭ Rectangle |
| `ROUNDED_RECTANGLE` | ▢ Rounded rectangle |
| `OVAL` | ● Circle/ellipse |
| `DIAMOND` | ◆ Diamond/rhombus |
| `ISOSCELES_TRIANGLE` | △ Triangle |
| `RIGHT_TRIANGLE` | ◩ Right triangle |
| `PARALLELOGRAM` | ▱ Parallelogram |
| `TRAPEZOID` | ⏢ Trapezoid |
| `HEXAGON` | ⬡ Hexagon |
| `OCTAGON` | ⬡ Octagon |
| `PENTAGON` | ⬠ Pentagon |
| `DONUT` | ◎ Donut/ring |

### Arrows (15+)
| Enum | Visual |
|------|--------|
| `RIGHT_ARROW` | → Right arrow |
| `LEFT_ARROW` | ← Left arrow |
| `UP_ARROW` | ↑ Up arrow |
| `DOWN_ARROW` | ↓ Down arrow |
| `LEFT_RIGHT_ARROW` | ↔ Bidirectional |
| `UP_DOWN_ARROW` | ↕ Bidirectional vertical |
| `BENT_ARROW` | ↗ Bent arrow |
| `BENT_UP_ARROW` | ⤴ Bent up arrow |
| `CIRCULAR_ARROW` | ↻ Circular arrow |
| `U_TURN_ARROW` | ↩ U-turn |
| `CURVED_RIGHT_ARROW` | ➝ Curved right |
| `NOTCHED_RIGHT_ARROW` | ➤ Notched arrow |
| `CHEVRON` | » Chevron |
| `QUAD_ARROW` | ✦ Four-way arrow |
| `LEFT_RIGHT_UP_ARROW` | Three-way arrow |

### Callouts (12+)
| Enum | Visual |
|------|--------|
| `RECTANGULAR_CALLOUT` | Rectangular speech bubble |
| `ROUNDED_RECTANGULAR_CALLOUT` | Rounded speech bubble |
| `OVAL_CALLOUT` | Oval speech bubble |
| `CLOUD_CALLOUT` | Cloud thought bubble |
| `LINE_CALLOUT_1` | Line callout (no accent) |
| `LINE_CALLOUT_2` | Line callout (accent bar) |
| `LINE_CALLOUT_3` | Line callout (border) |

### Flowchart (24)
| Enum | Visual |
|------|--------|
| `FLOWCHART_PROCESS` | ▭ Process box |
| `FLOWCHART_DECISION` | ◇ Decision diamond |
| `FLOWCHART_DATA` | ▱ Data parallelogram |
| `FLOWCHART_DOCUMENT` | 📄 Document (wavy bottom) |
| `FLOWCHART_PREDEFINED_PROCESS` | Predefined process (double side) |
| `FLOWCHART_INTERNAL_STORAGE` | Internal storage |
| `FLOWCHART_CONNECTOR` | ◎ Connector circle |
| `FLOWCHART_OFFPAGE_CONNECTOR` | Off-page connector |
| `FLOWCHART_CARD` | Card shape |
| `FLOWCHART_PUNCHED_TAPE` | Punched tape |
| `FLOWCHART_MANUAL_INPUT` | Manual input (trapezoid top) |
| `FLOWCHART_MANUAL_OPERATION` | Manual operation (trapezoid) |
| `FLOWCHART_MERGE` | Merge (arrow in) |
| `FLOWCHART_EXTRACT` | Extract (arrow out) |
| `FLOWCHART_SORT` | Sort (funnel) |
| `FLOWCHART_COLLATE` | Collate |
| `FLOWCHART_DELAY` | ⏳ Delay (half oval) |
| `FLOWCHART_STORED_DATA` | Stored data |
| `FLOWCHART_SEQUENTIAL_ACCESS_STORAGE` | Sequential access (cylinder) |
| `FLOWCHART_DIRECT_ACCESS_STORAGE` | Direct access (disk) |
| `FLOWCHART_DISPLAY` | Display |
| `FLOWCHART_SUMMING_JUNCTION` | Summing junction |
| `FLOWCHART_OR` | OR gate |
| `FLOWCHART_TERMINATOR` | Terminator (rounded rect) |

### Stars & Decorative
| Enum | Visual |
|------|--------|
| `STAR_4_POINT` through `STAR_12_POINT` | ★ 4-12 point stars |
| `STAR_5_POINT` | ★ Classic 5-point star |
| `STAR_6_POINT` | ✡ 6-point star |
| `STAR_8_POINT` | ✴ 8-point star |
| `LIGHTNING_BOLT` | ⚡ Lightning bolt |
| `HEART` | ♥ Heart |
| `SUN` | ☀ Sun |
| `MOON` | ☾ Moon |
| `CLOUD` | ☁ Cloud |
| `SMILEY_FACE` | 😊 Smiley |
| `NO_SYMBOL` | 🚫 No/prohibited |

### Math & Symbols
| Enum | Visual |
|------|--------|
| `MATH_PLUS` | ➕ Plus |
| `MATH_MINUS` | ➖ Minus |
| `MATH_MULTIPLY` | ✖ Multiply |
| `MATH_DIVIDE` | ➗ Divide |
| `MATH_EQUAL` | = Equal |
| `MATH_NOT_EQUAL` | ≠ Not equal |

### 3D & Special
| Enum | Visual |
|------|--------|
| `CUBE` | 🧊 3D cube |
| `CAN` | 🥫 3D cylinder |
| `BEVEL` | Beveled rectangle |
| `FOLDED_CORNER` | Page with folded corner |
| `FRAME` | Picture frame |
| `GEAR_6` / `GEAR_9` | ⚙ Gear shapes |
| `FUNNEL` | 🔻 Funnel |
| `CROSS` | ✚ Cross |
| `EXPLOSION1` / `EXPLOSION2` | 💥 Explosion shapes |

### Brackets & Braces
| Enum | Visual |
|------|--------|
| `LEFT_BRACE` / `RIGHT_BRACE` | { } Braces |
| `LEFT_BRACKET` / `RIGHT_BRACKET` | [ ] Brackets |
| `DOUBLE_BRACE` / `DOUBLE_BRACKET` | Double braces/brackets |

### Ribbons & Scrolls
| Enum | Visual |
|------|--------|
| `UP_RIBBON` / `DOWN_RIBBON` | Ribbon shapes |
| `CURVED_UP_RIBBON` / `CURVED_DOWN_RIBBON` | Curved ribbons |
| `VERTICAL_SCROLL` / `HORIZONTAL_SCROLL` | 📜 Scroll shapes |

### Snip & Round Variants
| Enum | Visual |
|------|--------|
| `SNIP_1_RECTANGLE` | 1-corner snip |
| `SNIP_2_SAME_RECTANGLE` | 2-corner snip (same side) |
| `SNIP_2_DIAG_RECTANGLE` | 2-corner snip (diagonal) |
| `SNIP_ROUND_RECTANGLE` | 1 snip + 1 round corner |
| `ROUND_1_RECTANGLE` | 1-corner round |
| `ROUND_2_SAME_RECTANGLE` | 2-corner round (same side) |
| `ROUND_2_DIAG_RECTANGLE` | 2-corner round (diagonal) |

### Usage
```python
shape = slide.shapes.add_shape(MSO_SHAPE.FLOWCHART_DECISION, left, top, width, height)
shape = slide.shapes.add_shape(MSO_SHAPE.CLOUD_CALLOUT, Inches(2), Inches(1), Inches(4), Inches(3))
```

---

## 2. Shape Creation Methods

| Method | Description | Example |
|--------|-------------|---------|
| `add_shape(type, left, top, width, height)` | Auto-shape | `add_shape(MSO_SHAPE.STAR_5_POINT, ...)` |
| `add_textbox(left, top, width, height)` | Text box | `add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))` |
| `add_picture(path, left, top, width=None, height=None)` | Image | `add_picture("photo.png", Inches(0), Inches(0))` |
| `add_chart(chart_type, left, top, width, height, data)` | Chart | `add_chart(XL_CHART_TYPE.PIE, ...)` |
| `add_table(rows, cols, left, top, width, height)` | Table | `add_table(3, 4, Inches(1), Inches(2), Inches(8), Inches(3))` |
| `add_connector(type, begin_x, begin_y, end_x, end_y)` | Connector | `add_connector(MSO_CONNECTOR_TYPE.ELBOW, ...)` |
| `add_group_shape()` | Group container | `add_group_shape()` |
| `add_movie(path, left, top, width, height, poster_frame, mime_type)` | Video | `add_movie("demo.mp4", ...)` |
| `add_ole_object(ole_file, left, top, width, height, prog_id, icon_file)` | Embedded object | `add_ole_object("data.xlsx", ...)` |
| `build_freeform(start_x, start_y, scale=Emu(914400))` | Custom path | `build_freeform(Inches(1), Inches(1))` |

---

## 3. Fill Types

### Solid Fill
```python
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)
```

### Gradient Fill (High-Level API)
```python
shape.fill.gradient()
shape.fill.gradient_stops[0].color.rgb = RGBColor(0xFF, 0x00, 0x00)
shape.fill.gradient_stops[0].position = 0.0
shape.fill.gradient_stops[1].color.rgb = RGBColor(0x00, 0x00, 0xFF)
shape.fill.gradient_stops[1].position = 1.0
```

### Pattern Fill (48 patterns)
```python
from pptx.enum.dml import MSO_PATTERN_TYPE
shape.fill.patterned()
shape.fill.pattern = MSO_PATTERN_TYPE.DARK_DOWNWARD_DIAGONAL
shape.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
shape.fill.back_color.rgb = RGBColor(0x00, 0x00, 0x00)
```

**Available patterns:** `CROSS`, `DARK_DOWNWARD_DIAGONAL`, `DARK_UPWARD_DIAGONAL`, `DARK_HORIZONTAL`, `DARK_VERTICAL`, `DIAGONAL_BRICK`, `HORIZONTAL_BRICK`, `LARGE_CHECKER_BOARD`, `LARGE_CONFETTI`, `LARGE_GRID`, `LIGHT_DOWNWARD_DIAGONAL`, `LIGHT_UPWARD_DIAGONAL`, `LIGHT_HORIZONTAL`, `LIGHT_VERTICAL`, `NARROW_HORIZONTAL`, `NARROW_VERTICAL`, `OUTLINED_DIAGONAL`, `PLAID`, `SHINGLE`, `SMALL_CHECKER_BOARD`, `SMALL_CONFETTI`, `SMALL_GRID`, `SOLID_DIAGONAL`, `SPHERE`, `TRELLIS`, `WAVE`, `WEAVE`, `ZIG_ZAG`, and more.

### No Fill (Transparent)
```python
shape.fill.background()
```

### Picture/Texture Fill (via XML)
```python
from pptx.oxml.ns import qn
from lxml import etree
spPr = shape._element.spPr
blipFill = etree.SubElement(spPr, qn("a:blipFill"))
blip = etree.SubElement(blipFill, qn("a:blip"))
blip.set(qn("r:embed"), rId)  # rId from image relationship
```

---

## 4. Line/Border

### Basic Line
```python
shape.line.color.rgb = RGBColor(0x00, 0x00, 0xFF)
shape.line.width = Pt(2)
```

### No Line
```python
shape.line.fill.background()
```

### Dash Styles (8 types, python-pptx 1.0.2 actual values)
```python
from pptx.enum.dml import MSO_LINE_DASH_STYLE
shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH
shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH_DOT
shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH_DOT_DOT
shape.line.dash_style = MSO_LINE_DASH_STYLE.LONG_DASH
shape.line.dash_style = MSO_LINE_DASH_STYLE.LONG_DASH_DOT
shape.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT
shape.line.dash_style = MSO_LINE_DASH_STYLE.SOLID
shape.line.dash_style = MSO_LINE_DASH_STYLE.SQUARE_DOT
```

### Gradient Line (via XML)
```python
from pptx.oxml.ns import qn
from lxml import etree
ln = shape._element.spPr.find(qn("a:ln"))
if ln is None:
    ln = etree.SubElement(shape._element.spPr, qn("a:ln"))
gradFill = etree.SubElement(ln, qn("a:gradFill"))
# Add gradient stops...
```

---

## 5. Text & Font

### Text Frame
```python
tf = shape.text_frame
tf.word_wrap = True
tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # Shrink text to fit
tf.margin_left = Inches(0.1)
tf.margin_right = Inches(0.1)
tf.margin_top = Inches(0.05)
tf.margin_bottom = Inches(0.05)
tf.vertical_anchor = MSO_ANCHOR.MIDDLE  # TOP / MIDDLE / BOTTOM
```

### Paragraph
```python
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT       # LEFT / CENTER / RIGHT / JUSTIFY / DISTRIBUTE
p.space_before = Pt(6)
p.space_after = Pt(6)
p.line_spacing = Pt(24)           # Fixed spacing
p.level = 0                       # Indent level (0-8)

# Add new paragraph
p2 = tf.add_paragraph()
```

**PP_ALIGN values:** `LEFT`, `CENTER`, `RIGHT`, `JUSTIFY`, `DISTRIBUTE`, `JUSTIFY_LOW`, `THAI_DISTRIBUTE`, `CENTER`

### Run & Font
```python
run = p.add_run()
run.text = "Hello World"
font = run.font
font.name = "Arial"
font.size = Pt(18)
font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
font.bold = True
font.italic = True
font.underline = True                    # Boolean or MSO_UNDERLINE enum
font.strikethrough = True
font.subscript = True
font.superscript = True
```

### Underline Styles (17 types)
```python
from pptx.enum.text import MSO_UNDERLINE
font.underline = MSO_UNDERLINE.SINGLE_LINE
font.underline = MSO_UNDERLINE.DOUBLE_LINE
font.underline = MSO_UNDERLINE.WAVY_LINE
font.underline = MSO_UNDERLINE.DASH_LINE
font.underline = MSO_UNDERLINE.DOTTED_LINE
font.underline = MSO_UNDERLINE.DASH_DOT_LINE
font.underline = MSO_UNDERLINE.DASH_DOT_DOT_LINE
font.underline = MSO_UNDERLINE.THICK_LINE
font.underline = MSO_UNDERLINE.WAVY_DOUBLE_LINE
# ... and more
```

### Theme Colors
```python
from pptx.enum.dml import MSO_THEME_COLOR
font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
font.color.theme_color = MSO_THEME_COLOR.DARK_1
font.color.theme_color = MSO_THEME_COLOR.LIGHT_1
font.color.theme_color = MSO_THEME_COLOR.TEXT_1
font.color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
font.color.theme_color = MSO_THEME_COLOR.HYPERLINK
font.color.brightness = 0.5  # Tint (positive) or shade (negative)
```

**MSO_THEME_COLOR values:** `ACCENT_1` through `ACCENT_6`, `DARK_1`, `DARK_2`, `LIGHT_1`, `LIGHT_2`, `BACKGROUND_1`, `BACKGROUND_2`, `TEXT_1`, `TEXT_2`, `HYPERLINK`, `FOLLOWED_HYPERLINK`

### CJK Font Companion (via XML)
```python
from pptx.oxml.ns import qn
rPr = run._element.find(qn("a:rPr"))
if rPr is None:
    rPr = run._element.get_or_add_rPr()
rPr.set(qn("a:ea"), "Microsoft YaHei")   # East Asian font
rPr.set(qn("a:cs"), "Arial")             # Complex script font
```

---

## 6. Tables

### Create Table
```python
table_shape = slide.shapes.add_table(rows=3, cols=4, left=Inches(1), top=Inches(2), width=Inches(8), height=Inches(3))
table = table_shape.table
```

### Cell Access & Formatting
```python
cell = table.cell(row=0, col=0)
cell.text = "Header"
cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Cell fill
cell.fill.solid()
cell.fill.fore_color.rgb = RGBColor(0x00, 0x70, 0xC0)

# Cell margins
cell.margin_left = Inches(0.1)
cell.margin_right = Inches(0.1)
cell.margin_top = Inches(0.05)
cell.margin_bottom = Inches(0.05)
cell.vertical_anchor = MSO_ANCHOR.MIDDLE
```

### Table Styling
```python
table.first_row = True    # Header row formatting
table.first_col = True    # First column formatting
table.last_row = True     # Total/footer row
table.last_col = True     # Last column formatting
table.horz_banding = True # Alternating row colors
table.vert_banding = True # Alternating column colors
```

### Column & Row Sizing
```python
table.columns[0].width = Inches(2.5)
table.rows[0].height = Inches(0.8)
```

### Iterate Cells
```python
for row in table.rows:
    for cell in row.cells:
        cell.text = "..."
```

---

## 7. Charts

### Chart Types (73 types)
```python
from pptx.enum.chart import XL_CHART_TYPE
```

**Common types:**
| Category | Types |
|----------|-------|
| Column | `COLUMN_CLUSTERED`, `COLUMN_STACKED`, `COLUMN_STACKED_100`, `COLUMN_3D_CLUSTERED`, `COLUMN_3D_STACKED`, `COLUMN_3D_STACKED_100`, `COLUMN_3D` |
| Bar | `BAR_CLUSTERED`, `BAR_STACKED`, `BAR_STACKED_100`, `BAR_3D_CLUSTERED`, `BAR_3D_STACKED`, `BAR_3D_STACKED_100` |
| Line | `LINE`, `LINE_MARKERS`, `LINE_STACKED`, `LINE_STACKED_100`, `LINE_MARKERS_STACKED`, `LINE_MARKERS_STACKED_100`, `LINE_3D` |
| Area | `AREA`, `AREA_STACKED`, `AREA_STACKED_100`, `AREA_3D`, `AREA_3D_STACKED`, `AREA_3D_STACKED_100` |
| Pie | `PIE`, `PIE_3D`, `PIE_OF_PIE`, `BAR_OF_PIE`, `DOUGHNUT`, `DOUGHNUT_EXPLODED`, `PIE_EXPLODED`, `PIE_3D_EXPLODED` |
| Scatter | `XY_SCATTER`, `XY_SCATTER_LINES`, `XY_SCATTER_LINES_NO_MARKERS`, `XY_SCATTER_SMOOTH`, `XY_SCATTER_SMOOTH_NO_MARKERS` (**NOTE**: python-pptx 1.0.2 has a bug with XY_SCATTER using CategoryChartData — use `XyChartData` or `LINE_MARKERS` as workaround) |
| Radar | `RADAR`, `RADAR_FILLED`, `RADAR_MARKERS` |
| Stock | `STOCK_HLC`, `STOCK_OHLC`, `STOCK_VHLC`, `STOCK_VOHLC` |
| Surface | `SURFACE_3D`, `SURFACE_WIREFRAME_3D`, `SURFACE_CONTOUR`, `SURFACE_CONTOUR_WIREFRAME` |
| Bubble | `BUBBLE`, `BUBBLE_3D_EFFECT` |
| Cone/Cylinder/Pyramid | `CONE_COL_CLUSTERED`, `CONE_COL_STACKED`, `CYLINDER_COL_CLUSTERED`, `PYRAMID_COL_CLUSTERED`, etc. |

### Create Chart
```python
from pptx.chart.data import CategoryChartData, XyChartData, BubbleChartData

# Category chart (bar, column, line, pie, etc.)
chart_data = CategoryChartData()
chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
chart_data.add_series("Revenue", (30, 45, 60, 75))
chart_data.add_series("Cost", (20, 30, 35, 40))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(2), Inches(8), Inches(4),
    chart_data
)
chart = chart_frame.chart
```

### Chart Styling
```python
# Legend
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False

# Axis
value_axis = chart.value_axis
value_axis.has_title = True
value_axis.axis_title.text_frame.paragraphs[0].text = "Revenue ($M)"
value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
value_axis.tick_labels.number_format = '#,##0'

category_axis = chart.category_axis
category_axis.has_tick_labels = True

# Series styling
series = chart.series[0]
series.format.fill.solid()
series.format.fill.fore_color.rgb = RGBColor(0x00, 0x70, 0xC0)
series.smooth = True  # Smooth line for line charts
```

### XY/Scatter Chart
```python
chart_data = XyChartData()
series = chart_data.add_series("Data")
series.add_data_point(1, 10)
series.add_data_point(2, 25)
series.add_data_point(3, 18)
```

---

## 8. Connectors

```python
from pptx.enum.shapes import MSO_CONNECTOR_TYPE

# Straight connector
connector = slide.shapes.add_connector(
    MSO_CONNECTOR_TYPE.STRAIGHT,
    Inches(1), Inches(1),  # begin_x, begin_y
    Inches(5), Inches(3)   # end_x, end_y
)

# Elbow (right-angle) connector
connector = slide.shapes.add_connector(
    MSO_CONNECTOR_TYPE.ELBOW,
    Inches(1), Inches(1), Inches(5), Inches(3)
)

# Curved connector
connector = slide.shapes.add_connector(
    MSO_CONNECTOR_TYPE.CURVE,
    Inches(1), Inches(1), Inches(5), Inches(3)
)

# Style the connector line
connector.line.color.rgb = RGBColor(0x00, 0x00, 0xFF)
connector.line.width = Pt(2)
connector.line.dash_style = MSO_LINE_DASH_STYLE.DASH
```

---

## 9. Group Shapes

```python
# Create group
group = slide.shapes.add_group_shape()

# Add shapes to group
child = group.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)

# Access children
for shape in group.shapes:
    print(shape.shape_type)
```

---

## 10. Freeform Paths

```python
# Official API (python-pptx 1.0.2)
freeform = slide.shapes.build_freeform(Inches(1), Inches(1))
freeform.add_line_segments([
    (Inches(5), Inches(1)),
    (Inches(5), Inches(5)),
    (Inches(1), Inches(5)),
    (Inches(1), Inches(1)),  # Close back to start
])
shape = freeform.convert_to_shape()

# Move pen without drawing
freeform = slide.shapes.build_freeform(Inches(0), Inches(0))
freeform.move_to(Inches(1), Inches(1))
freeform.add_line_segments([(Inches(5), Inches(3))])
shape = freeform.convert_to_shape()
```

---

## 11. Hyperlinks & Actions

### Shape Click Action
```python
from pptx.enum.action import PP_ACTION

# Hyperlink on shape
shape.click_action.hyperlink.address = "https://example.com"

# Navigate to slide
shape.click_action.action = PP_ACTION.NEXT_SLIDE
shape.click_action.action = PP_ACTION.LAST_SLIDE
shape.click_action.action = PP_ACTION.FIRST_SLIDE

# Run program
shape.click_action.action = PP_ACTION.RUN_PROGRAM
```

### Hyperlink on Text Run
```python
run = p.add_run()
run.text = "Click here"
run.hyperlink.address = "https://example.com"
```

### Hover Action
```python
shape.hover_action.hyperlink.address = "https://example.com"
```

**PP_ACTION values:** `HYPERLINK`, `NEXT_SLIDE`, `PREVIOUS_SLIDE`, `FIRST_SLIDE`, `LAST_SLIDE`, `END_SHOW`, `NAMED_SLIDE`, `RUN_MACRO`, `RUN_PROGRAM`, `OPEN_FILE`, `OLE_VERB`, `NONE`

---

## 12. Media (Video/Audio)

```python
# Embed video
movie = slide.shapes.add_movie(
    movie_file="demo.mp4",
    left=Inches(1), top=Inches(2),
    width=Inches(6), height=Inches(4),
    poster_frame_image="poster.png",  # Optional poster frame
    mime_type="video/mp4"
)
```

---

## 13. OLE Objects (Embedded Files)

```python
# Embed Excel file
ole = slide.shapes.add_ole_object(
    ole_file="data.xlsx",
    left=Inches(2), top=Inches(2),
    width=Inches(4), height=Inches(3),
    prog_id="Excel.Sheet.12",
    icon_file="excel_icon.png"  # Optional custom icon
)
```

**Common PROG_IDs:** `Excel.Sheet.12`, `Word.Document.12`, `PowerPoint.Show.12`

---

## 14. Picture Cropping

```python
picture = slide.shapes.add_picture("photo.jpg", Inches(0), Inches(0), Inches(10), Inches(7))

# Crop via python-pptx API
picture.crop_left = 0.1    # 10% crop from left
picture.crop_right = 0.05  # 5% crop from right
picture.crop_top = 0.0
picture.crop_bottom = 0.15 # 15% crop from bottom
```

---

## 15. Slide Background

```python
# Solid background
slide.background.fill.solid()
slide.background.fill.fore_color.rgb = RGBColor(0x00, 0x00, 0x00)

# Gradient background (via XML)
from pptx.oxml.ns import qn
from lxml import etree
bgPr = slide.background._element
# ... add gradient fill XML
```

---

## 16. Slide Notes

```python
# Add speaker notes
if not slide.has_notes_slide:
    notes_slide = slide.notes_slide
else:
    notes_slide = slide.notes_slide
notes_slide.notes_text_frame.text = "Speaker notes for this slide"
```

---

## 17. Core Properties (Metadata)

```python
prs.core_properties.author = "John Doe"
prs.core_properties.title = "Q4 Report"
prs.core_properties.subject = "Financial Analysis"
prs.core_properties.keywords = "finance, Q4, report"
prs.core_properties.comments = "Internal use only"
prs.core_properties.category = "Financial Reports"
prs.core_properties.created = datetime(2024, 1, 1)
prs.core_properties.modified = datetime.now()
```

---

## 18. Shape Properties

### Position & Size
```python
shape.left = Inches(1)
shape.top = Inches(2)
shape.width = Inches(5)
shape.height = Inches(3)
shape.rotation = 45.0  # Degrees
```

### Read Properties
```python
left = shape.left       # EMU
top = shape.top         # EMU
w = shape.width         # EMU
h = shape.height        # EMU
rot = shape.rotation    # Degrees
sid = shape.shape_id    # Integer ID
stype = shape.shape_type  # MSO_SHAPE_TYPE enum
has_tf = shape.has_text_frame  # Boolean
```

### Placeholder Detection
```python
if shape.is_placeholder:
    ph_type = shape.placeholder_format.type  # e.g., TITLE, BODY, PICTURE
    ph_idx = shape.placeholder_format.idx    # Integer index
```

---

## 19. Effects (via XML)

### Shadow
```python
from pptx.oxml.ns import qn
from lxml import etree

spPr = shape._element.spPr
effectLst = spPr.find(qn("a:effectLst"))
if effectLst is None:
    effectLst = etree.SubElement(spPr, qn("a:effectLst"))
outerShdw = etree.SubElement(effectLst, qn("a:outerShdw"))
outerShdw.set("blurRad", "76200")     # 6pt blur (EMU)
outerShdw.set("dist", "38100")        # 3pt distance
outerShdw.set("dir", "5400000")       # 90° direction (60000ths of degree)
outerShdw.set("algn", "tl")           # Alignment
srgbClr = etree.SubElement(outerShdw, qn("a:srgbClr"))
srgbClr.set("val", "000000")
alpha = etree.SubElement(srgbClr, qn("a:alpha"))
alpha.set("val", "25000")             # 25% opacity (0-100000)
```

### Glow
```python
glow = etree.SubElement(effectLst, qn("a:glow"))
glow.set("rad", "76200")  # 6pt radius
srgbClr = etree.SubElement(glow, qn("a:srgbClr"))
srgbClr.set("val", "00BFFF")
alpha = etree.SubElement(srgbClr, qn("a:alpha"))
alpha.set("val", "40000")  # 40% opacity
```

### Soft Edge
```python
softEdge = etree.SubElement(effectLst, qn("a:softEdge"))
softEdge.set("rad", "50800")  # 4pt radius
```

### Reflection
```python
reflection = etree.SubElement(effectLst, qn("a:reflection"))
reflection.set("blurRad", "6350")
reflection.set("stA", "50000")   # Start alpha 50%
reflection.set("stPos", "0")     # Start position
reflection.set("endA", "0")      # End alpha 0%
reflection.set("endPos", "85000") # End position 85%
reflection.set("dist", "38100")
reflection.set("dir", "5400000")
```

### 3D Effects (via XML)
```python
# 3D shape properties
sp3d = etree.SubElement(spPr, qn("a:sp3d"))
sp3d.set("z", "254000")  # Extrusion depth

# Bevel
bevel = etree.SubElement(sp3d, qn("a:bevelT"))
bevel.set("w", "63500")   # Width
bevel.set("h", "25400")   # Height

# Material
prstMaterial = etree.SubElement(sp3d, qn("a:prstMaterial"))
prstMaterial.set("val", "powder")  # matte, powder, metal, plastic, etc.

# 3D scene (lighting)
scene3d = etree.SubElement(spPr, qn("a:scene3d"))
camera = etree.SubElement(scene3d, qn("a:camera"))
camera.set("prst", "perspectiveFront")
lightRig = etree.SubElement(scene3d, qn("a:lightRig"))
lightRig.set("rig", "threePt")
lightRig.set("dir", "t")
```

---

## 20. Transitions & Animations (via XML)

### Slide Transition
```python
from pptx.oxml.ns import qn
from lxml import etree

# Add transition
transition = etree.SubElement(slide._element, qn("p:transition"))
transition.set("spd", "med")  # spd: "slow", "med", "fast"

# Fade transition
fade = etree.SubElement(transition, qn("p:fade"))
fade.set("thruBlk", "0")

# Push transition
push = etree.SubElement(transition, qn("p:push"))
push.set("dir", "l")  # l=left, r=right, u=up, d=down

# Wipe transition
wipe = etree.SubElement(transition, qn("p:wipe"))
wipe.set("dir", "r")

# Other: p:cover, p:dissolve, p:wheel, p:wedge, p:split, p:blinds, p:checker, p:comb
```

### Entrance Animation
```python
# Add animation sequence
timing = etree.SubElement(slide._element, qn("p:timing"))
tnLst = etree.SubElement(timing, qn("p:tnLst"))
par = etree.SubElement(tnLst, qn("p:par"))
# ... complex animation XML (see animation.py for examples)
```

---

## 21. Turbo-Add Mode (Performance)

For slides with 100+ shapes, enable turbo mode for faster shape addition:

```python
slide.shapes.turbo_add_enabled = True
# Add shapes...
slide.shapes.turbo_add_enabled = False  # Disable when done
```

---

## 22. OOXML Quick Reference

### Common Namespace Prefixes
| Prefix | URI | Purpose |
|--------|-----|---------|
| `p:` | `http://schemas.openxmlformats.org/presentationml/2006/main` | Presentation |
| `a:` | `http://schemas.openxmlformats.org/drawingml/2006/main` | Drawing |
| `r:` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationships |

### Key Elements
| Element | Parent | Purpose |
|---------|--------|---------|
| `p:sp` | `p:spTree` | Shape |
| `p:spPr` | `p:sp` | Shape properties |
| `a:solidFill` | `p:spPr` | Solid fill |
| `a:gradFill` | `p:spPr` | Gradient fill |
| `a:noFill` | `p:spPr` | No fill |
| `a:ln` | `p:spPr` | Line/border |
| `a:effectLst` | `p:spPr` | Effects container |
| `a:outerShdw` | `a:effectLst` | Drop shadow |
| `a:glow` | `a:effectLst` | Glow effect |
| `a:sp3d` | `p:spPr` | 3D shape properties |
| `p:txBody` | `p:sp` | Text body |
| `a:r` | `a:p` | Text run |
| `a:rPr` | `a:r` | Run properties |
| `a:latin` | `a:rPr` | Latin font |
| `a:ea` | `a:rPr` | East Asian font |
| `a:cs` | `a:rPr` | Complex script font |
| `a:alpha` | `a:srgbClr` | Alpha/transparency (val = pct × 1000) |
| `p:transition` | `p:sld` | Slide transition |
| `p:timing` | `p:sld` | Animation timing |
| `p:grpSp` | `p:spTree` | Group shape |
| `p:cSld` | `p:sld` | Slide content |
| `p:spTree` | `p:cSld` | Shape tree |

### Alpha Values
```python
# OOXML alpha: value = percentage × 1000
# 100% = 100000
# 50%  = 50000
# 25%  = 25000
# 10%  = 10000
alpha.set("val", "50000")  # 50% opacity
```

### EMU Conversions
```python
# English Metric Units (EMU)
# 1 inch = 914400 EMU
# 1 pt   = 12700 EMU
# 1 cm   = 360000 EMU
# 1 mm   = 36000 EMU

Inches(1)   # → 914400
Pt(12)      # → 152400
Emu(914400) # → 914400
```

---

## 23. Common Patterns

### Cover-Fit Image (no stretch)
```python
from PIL import Image as PILImage

# Pre-crop with Pillow, then add_picture
img = PILImage.open(image_path)
target_ratio = width / height
img_ratio = img.width / img.height
if img_ratio > target_ratio:
    new_w = int(img.height * target_ratio)
    left = (img.width - new_w) // 2
    img = img.crop((left, 0, left + new_w, img.height))
else:
    new_h = int(img.width / target_ratio)
    top = (img.height - new_h) // 2
    img = img.crop((0, top, img.width, top + new_h))
img.save(cropped_path)
slide.shapes.add_picture(cropped_path, left, top, width, height)
```

### Rounded Rectangle with Custom Radius
```python
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
# Adjust corner radius via XML
from pptx.oxml.ns import qn
spPr = shape._element.spPr
prstGeom = spPr.find(qn("a:prstGeom"))
avLst = prstGeom.find(qn("a:avLst"))
if avLst is None:
    from lxml import etree
    avLst = etree.SubElement(prstGeom, qn("a:avLst"))
# Clear existing
for gd in avLst.findall(qn("a:gd")):
    avLst.remove(gd)
from lxml import etree
gd = etree.SubElement(avLst, qn("a:gd"))
gd.set("name", "adj")
gd.set("fmla", "val 16667")  # 16667 = ~17% radius (0-50000)
```

### Gradient Overlay on Shape
```python
from pptx.oxml.ns import qn
from lxml import etree

spPr = shape._element.spPr
gradFill = etree.SubElement(spPr, qn("a:gradFill"))
gsLst = etree.SubElement(gradFill, qn("a:gsLst"))

# Stop 1: transparent
gs1 = etree.SubElement(gsLst, qn("a:gs"))
gs1.set("pos", "0")
srgb1 = etree.SubElement(gs1, qn("a:srgbClr"))
srgb1.set("val", "000000")
alpha1 = etree.SubElement(srgb1, qn("a:alpha"))
alpha1.set("val", "0")  # 0% opacity

# Stop 2: dark
gs2 = etree.SubElement(gsLst, qn("a:gs"))
gs2.set("pos", "100000")
srgb2 = etree.SubElement(gs2, qn("a:srgbClr"))
srgb2.set("val", "000000")
alpha2 = etree.SubElement(srgb2, qn("a:alpha"))
alpha2.set("val", "80000")  # 80% opacity

# Linear gradient (top to bottom)
lin = etree.SubElement(gradFill, qn("a:lin"))
lin.set("ang", "5400000")  # 90° (top to bottom)
lin.set("scaled", "1")
```
