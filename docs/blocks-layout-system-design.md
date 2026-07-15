# Blocks Layout System Design

> Version: 1.0 | Date: 2026-07-15 | Status: Design Phase

---

## 1. Problem Statement

### 1.1 Current Architecture Defect

`PrecisionRenderer.render_slide()` uses a **mutually exclusive elif chain** to dispatch content rendering:

```python
# precision_renderer.py:483-506
if cards:          → horizontal card row (always horizontal, fixed 4.5" height)
elif component:    → component library injection
elif diagram:      → diagram engine (fixed region 0.9,1.5,7.0,5.0)
elif code:         → code block (fixed #1E293B bg, 11.533" wide)
elif exercise:     → badge + steps
elif bullets:      → bullet list (left-aligned, 6+ → double column)
```

**Consequences:**

1. **Elements are mutually exclusive** — a page cannot have both cards AND diagram, both bullets AND code
2. **Positions are hardcoded** — image always at (8.3, 1.2), cards always horizontal, title always at (margin_left, 0.5)
3. **No layout composition** — impossible to create "3 hexagons on left + bullets on right" or "cards on top + flowchart on bottom"
4. **All pages look identical** — changing style only changes colors, not structure
5. **Section divider bug** — `render_section_divider()` creates a second slide, leaving the first blank

### 1.2 Impact

- 3 style proposals with identical layout, only colors differ
- Users perceive the system as "color-swapping garbage"
- Build Script mode is the only way to get differentiated layouts, but loses brand compliance, version management, CJK pairing, and 28 design upgrades

### 1.3 Goal

Transform `render_slide()` from a **single-choice dispatcher** to a **composable blocks system** where:

1. A page can contain **multiple blocks** (cards + diagram, bullets + image, hexagons + text, etc.)
2. Each block has an explicit **region** (x, y, w, h) — either user-specified or auto-computed
3. **Backward compatible** — old content.json format (cards, bullets, diagram as top-level keys) still works
4. **Brand compliance preserved** — all blocks read from BrandSpec, CJK pairing, design upgrades still apply
5. **Section divider bug fixed** — no more blank slides

---

## 2. Architecture

### 2.1 Core Concept: Block

A block is the smallest renderable unit. Each block has:

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `type` | string | Yes | Block type: `cards`, `bullets`, `diagram`, `code`, `exercise`, `image`, `hexagons`, `ovals`, `donuts`, `metrics`, `badge`, `gradient_line`, `table_chart` |
| `region` | string or dict | No | Layout region. String = preset name, dict = `{x, y, w, h}` in inches. Auto-computed if omitted |
| `data` | dict | Yes* | Block-specific data (*some blocks infer data from page-level fields) |
| `style` | dict | No | Block-level style overrides (font_size, color_role, corner_radius, etc.) |

### 2.2 Region System

#### Preset Regions

| Preset | x | y | w | h | Use Case |
|--------|---|---|---|---|----------|
| `full` | 0.9 | 1.6 | 11.533 | 5.4 | Single block fills content area |
| `left` | 0.9 | 1.6 | 5.566 | 5.4 | Left half |
| `right` | 6.866 | 1.6 | 5.566 | 5.4 | Right half |
| `left-2-3` | 0.9 | 1.6 | 7.422 | 5.4 | Left 2/3 |
| `right-1-3` | 8.722 | 1.6 | 3.711 | 5.4 | Right 1/3 |
| `left-1-3` | 0.9 | 1.6 | 3.711 | 5.4 | Left 1/3 |
| `right-2-3` | 5.011 | 1.6 | 7.422 | 5.4 | Right 2/3 |
| `top` | 0.9 | 1.6 | 11.533 | 2.5 | Top band |
| `bottom` | 0.9 | 4.5 | 11.533 | 2.5 | Bottom band |
| `top-left` | 0.9 | 1.6 | 5.566 | 2.5 | Top-left quadrant |
| `top-right` | 6.866 | 1.6 | 5.566 | 2.5 | Top-right quadrant |
| `bottom-left` | 0.9 | 4.5 | 5.566 | 2.5 | Bottom-left quadrant |
| `bottom-right` | 6.866 | 4.5 | 5.566 | 2.5 | Bottom-right quadrant |
| `center` | 2.5 | 2.0 | 8.333 | 4.6 | Centered block |
| `sidebar` | 0.0 | 0.0 | 3.0 | 7.5 | Full-height sidebar (overrides brand background strip) |

> **Geometry basis**: Slide 13.333"×7.5", margin_left=0.9, margin_right=0.9, content_top=1.6 (after title 0.5+0.6h+0.5gap), margin_bottom=0.5. Content area = 11.533"×5.4". Inter-region gap = 0.4".

> **sidebar 注意**: sidebar region 从 x=0 开始，会覆盖 brand background 的左侧 accent strip。当页面使用 sidebar block 时，`apply_brand_background()` 应跳过该页面的 left strip。

#### Custom Region

```json
{"x": 1.2, "y": 2.0, "w": 5.0, "h": 4.5}
```

#### Auto-Compute (region omitted)

When `region` is omitted, the system auto-computes based on:
1. Number of blocks on the page
2. Block types (some blocks prefer horizontal, some vertical)
3. Available content area (after title + gradient line)

Auto-compute rules:

| Block Count | Layout | Regions |
|-------------|--------|---------|
| 1 | Full | `full` |
| 2 | Side-by-side | Block 0 → `left`, Block 1 → `right` |
| 3 | Left 2/3 + Right 1/3 stacked | Block 0 → `left-2-3`, Block 1 → `top-right`, Block 2 → `bottom-right` |
| 4 | 2×2 grid | Block 0 → `top-left`, Block 1 → `top-right`, Block 2 → `bottom-left`, Block 3 → `bottom-right` |

### 2.3 Block Types

#### 2.3.1 `cards` Block

Renders a row or grid of cards.

```json
{
  "type": "cards",
  "region": "left-2-3",
  "data": {
    "items": [
      {"title": "FreeStyle", "text": "一句话生成，40,000+风格"},
      {"title": "Enterprise", "text": "品牌合规，版本管理"},
      {"title": "Build Script", "text": "逐页精确控制"}
    ],
    "layout": "horizontal",
    "featured_index": 0
  },
  "style": {"corner_radius": "lg", "shadow": true}
}
```

| Data Field | Type | Default | Description |
|------------|------|---------|-------------|
| `items` | array | required | Card items with `title` and `text` |
| `layout` | string | `"horizontal"` | `"horizontal"`, `"vertical"`, `"grid-2x2"` |
| `featured_index` | int | `0` | Which card gets featured treatment (-1 = none) |

**New capabilities vs old**: vertical layout, grid-2x2, featured_index control, custom region

#### 2.3.2 `bullets` Block

Renders a bullet list, single or multi-column.

```json
{
  "type": "bullets",
  "region": "right-1-3",
  "data": {
    "items": ["叙事规划", "设计决策", "内容生成", "PPT渲染"],
    "columns": 1,
    "bullet_style": "dash"
  },
  "style": {"font_size": 14, "color_role": "foreground"}
}
```

| Data Field | Type | Default | Description |
|------------|------|---------|-------------|
| `items` | array | required | Bullet text items |
| `columns` | int | auto | Number of columns (auto = 1 if <6, 2 if >=6) |
| `bullet_style` | string | `"dot"` | `"dot"`, `"dash"`, `"number"`, `"none"` |

**New capabilities vs old**: custom columns, bullet style, custom region, can coexist with other blocks

#### 2.3.3 `diagram` Block

Renders a diagram via DiagramEngine.

```json
{
  "type": "diagram",
  "region": "left",
  "data": {
    "diagram_type": "flowchart",
    "diagram_data": {
      "nodes": [{"id": "1", "label": "叙事规划"}, {"id": "2", "label": "设计决策"}],
      "connectors": [{"from": "1", "to": "2"}]
    }
  }
}
```

| Data Field | Type | Default | Description |
|------------|------|---------|-------------|
| `diagram_type` | string | required | flowchart/funnel/timeline/swot/matrix/cycle/table/hierarchy/pyramid/venn |
| `diagram_data` | dict | required | Type-specific data structure |

**New capabilities vs old**: custom region, can coexist with other blocks

#### 2.3.4 `code` Block

Renders a code block with language badge.

```json
{
  "type": "code",
  "region": "full",
  "data": {
    "language": "python",
    "source": "from ppt_pro_max import generate_ppt\nresult = generate_ppt('AI pitch')"
  },
  "style": {"bg_color": "#1E293B", "font": "Consolas"}
}
```

**New capabilities vs old**: custom region, custom bg_color (brand-aware), can coexist with other blocks

#### 2.3.5 `exercise` Block

Renders exercise badge + instructions + steps.

```json
{
  "type": "exercise",
  "region": "full",
  "data": {
    "instructions": "安装并生成你的第一个AI演示文稿",
    "duration": "5 min",
    "steps": ["git clone并运行install.py", "执行ppt-design命令", "打开.pptx查看效果"]
  }
}
```

**New capabilities vs old**: custom region, can coexist with other blocks

#### 2.3.6 `image` Block

Renders a masked image in a rounded-rect frame.

```json
{
  "type": "image",
  "region": "right-1-3",
  "data": {
    "path": "images/hero.png",
    "mask": true,
    "cover": false
  },
  "style": {"corner_radius": "lg", "shadow": true}
}
```

| Data Field | Type | Default | Description |
|------------|------|---------|-------------|
| `path` | string | required | Image file path |
| `mask` | bool | `true` | Rounded-rect frame with padding |
| `cover` | bool | `false` | Full-bleed cover fit (for hero backgrounds) |

**New capabilities vs old**: custom position (not always right-side), can be primary content block

#### 2.3.7 `hexagons` Block (NEW)

Renders a honeycomb grid of hexagons with labels.

```json
{
  "type": "hexagons",
  "region": "left",
  "data": {
    "items": [
      {"label": "叙事", "color_role": "primary"},
      {"label": "设计", "color_role": "accent"},
      {"label": "内容", "color_role": "primary"},
      {"label": "渲染", "color_role": "accent"}
    ],
    "layout": "honeycomb"
  }
}
```

| Data Field | Type | Default | Description |
|------------|------|---------|-------------|
| `items` | array | required | Hexagon items with `label` and optional `color_role` |
| `layout` | string | `"honeycomb"` | `"honeycomb"`, `"row"`, `"grid-2x2"` |

#### 2.3.8 `ovals` Block (NEW)

Renders circular elements with labels, useful for cycle/process visualization.

```json
{
  "type": "ovals",
  "region": "left-2-3",
  "data": {
    "items": [
      {"label": "01", "subtitle": "规划", "color_role": "primary"},
      {"label": "02", "subtitle": "决策", "color_role": "accent"},
      {"label": "03", "subtitle": "生成", "color_role": "primary"},
      {"label": "04", "subtitle": "渲染", "color_role": "accent"}
    ],
    "layout": "horizontal",
    "show_connectors": true
  }
}
```

#### 2.3.9 `donuts` Block (NEW)

Renders donut/ring shapes with center labels, useful for KPI/metrics.

```json
{
  "type": "donuts",
  "region": "right-1-3",
  "data": {
    "items": [
      {"label": "40K+", "subtitle": "风格组合"},
      {"label": "28", "subtitle": "设计升级"},
      {"label": "10", "subtitle": "图形引擎"}
    ],
    "layout": "vertical"
  }
}
```

#### 2.3.10 `metrics` Block (NEW)

Renders large number + label pairs, useful for data/traction pages.

```json
{
  "type": "metrics",
  "region": "top",
  "data": {
    "items": [
      {"value": "40,000+", "label": "风格组合", "color_role": "primary"},
      {"value": "28", "label": "设计升级", "color_role": "accent"},
      {"value": "10", "label": "图形引擎", "color_role": "primary"},
      {"value": "824", "label": "测试通过", "color_role": "accent"}
    ],
    "layout": "horizontal"
  }
}
```

#### 2.3.11 `badge` Block

Renders a standalone badge/tag.

```json
{
  "type": "badge",
  "region": {"x": 0.9, "y": 1.2, "w": 1.5, "h": 0.35},
  "data": {
    "text": "NEW IN V0.7",
    "variant": "solid"
  }
}
```

#### 2.3.12 `gradient_line` Block

Renders a gradient accent line.

```json
{
  "type": "gradient_line",
  "region": {"x": 0.9, "y": 1.15, "w": 3.0, "h": 0.04},
  "data": {
    "color_role": "accent"
  }
}
```

#### 2.3.13 `table_chart` Block

Renders a data table with alternating row colors.

```json
{
  "type": "table_chart",
  "region": "full",
  "data": {
    "headers": ["引擎", "类型", "模型"],
    "rows": [
      ["Seedream", "AI生成", "doubao-seedream-5-0"],
      ["GPT Image", "AI生成", "gpt-image-1"],
      ["DALL-E 3", "AI生成", "dall-e-3"]
    ]
  }
}
```

### 2.4 content.json Format

#### New Format (blocks)

```json
{
  "meta": {"title": "PPT Design Skill"},
  "slides": [
    {
      "goal": "hook",
      "title": "PPT Design Skill",
      "subtitle": "AI驱动的专业演示文稿引擎",
      "blocks": [
        {"type": "metrics", "region": "bottom", "data": {
          "items": [
            {"value": "40,000+", "label": "风格组合"},
            {"value": "28", "label": "设计升级"},
            {"value": "10", "label": "图形引擎"}
          ]
        }}
      ]
    },
    {
      "goal": "content",
      "title": "4阶段智能流水线",
      "blocks": [
        {"type": "hexagons", "region": "left-2-3", "data": {
          "items": [
            {"label": "叙事规划"}, {"label": "设计决策"},
            {"label": "内容生成"}, {"label": "PPT渲染"}
          ]
        }},
        {"type": "bullets", "region": "right-1-3", "data": {
          "items": ["策略选择+情绪弧线", "40,000+风格自动匹配", "PAS/FAB/AIDA文案", "python-pptx直出+QA门禁"]
        }}
      ]
    },
    {
      "goal": "content",
      "title": "三模式引擎",
      "blocks": [
        {"type": "cards", "region": "full", "data": {
          "items": [
            {"title": "FreeStyle", "text": "一句话生成"},
            {"title": "Enterprise", "text": "品牌合规"},
            {"title": "Build Script", "text": "精确交付"}
          ]
        }}
      ]
    }
  ]
}
```

#### Old Format (backward compatible)

```json
{
  "goal": "features",
  "title": "三模式引擎",
  "cards": [
    {"title": "FreeStyle", "text": "一句话生成"},
    {"title": "Enterprise", "text": "品牌合规"},
    {"title": "Build Script", "text": "精确交付"}
  ]
}
```

**Backward compatibility rule**: If `blocks` key is present, use blocks system. If absent, fall back to old elif chain dispatch. Zero breaking change.

### 2.5 render_slide() New Dispatch Logic

```python
def render_slide(self, prs, page, component_lib=None,
                 layout_variant=None, page_index=0, total_pages=0):
    goal = page.get("goal", "content")
    title = page.get("title", "")
    blocks = page.get("blocks")

    # ── Section divider (BUG FIXED) ──
    if goal == "section":
        slide = self.add_slide(prs)
        section_num = page.get("section_number", page_index + 1)
        section_sub = page.get("subtitle") or ""
        self.render_section_divider(slide, section_num, title, section_sub)
        return slide

    # ── Hero slides (hook/cta) ──
    is_hero = goal in ("hook", "cta")
    slide = self.add_slide(prs)

    if is_hero:
        self._render_hero(slide, page)
        # Hero can also have blocks (e.g., metrics at bottom)
        if blocks:
            self._render_blocks(slide, blocks, is_hero=True)
        return slide

    # ── Content slides ──
    self.apply_brand_background(slide, prs, goal, page_index, total_pages)

    if title:
        self._render_title_block(slide, page, layout_variant)

    if blocks:
        # NEW PATH: composable blocks
        self._render_blocks(slide, blocks)
    else:
        # OLD PATH: backward compatible elif chain
        self._render_legacy_content(slide, page, component_lib)

    if total_pages > 0:
        self.add_progress_bar(slide, page_index + 1, total_pages)

    return slide
```

### 2.6 BlockRenderer

New class that renders individual blocks onto a slide within a given region.

```python
class BlockRenderer:
    """Renders individual content blocks within specified regions."""

    REGION_PRESETS = {
        "full":          {"x": 0.9,   "y": 1.6,  "w": 11.533, "h": 5.4},
        "left":          {"x": 0.9,   "y": 1.6,  "w": 5.566,  "h": 5.4},
        "right":         {"x": 6.866, "y": 1.6,  "w": 5.566,  "h": 5.4},
        "left-2-3":      {"x": 0.9,   "y": 1.6,  "w": 7.422,  "h": 5.4},
        "right-1-3":     {"x": 8.722, "y": 1.6,  "w": 3.711,  "h": 5.4},
        "left-1-3":      {"x": 0.9,   "y": 1.6,  "w": 3.711,  "h": 5.4},
        "right-2-3":     {"x": 5.011, "y": 1.6,  "w": 7.422,  "h": 5.4},
        "top":           {"x": 0.9,   "y": 1.6,  "w": 11.533, "h": 2.5},
        "bottom":        {"x": 0.9,   "y": 4.5,  "w": 11.533, "h": 2.5},
        "top-left":      {"x": 0.9,   "y": 1.6,  "w": 5.566,  "h": 2.5},
        "top-right":     {"x": 6.866, "y": 1.6,  "w": 5.566,  "h": 2.5},
        "bottom-left":   {"x": 0.9,   "y": 4.5,  "w": 5.566,  "h": 2.5},
        "bottom-right":  {"x": 6.866, "y": 4.5,  "w": 5.566,  "h": 2.5},
        "center":        {"x": 2.5,   "y": 2.0,  "w": 8.333,  "h": 4.6},
        "sidebar":       {"x": 0.0,   "y": 0.0,  "w": 3.0,    "h": 7.5},
    }

    def __init__(self, precision_renderer: PrecisionRenderer):
        self._pr = precision_renderer

    def render(self, slide, blocks: list[dict], is_hero: bool = False) -> None:
        for i, block in enumerate(blocks):
            region = self._resolve_region(block, i, len(blocks), is_hero)
            self._render_block(slide, block, region)

    def _resolve_region(self, block: dict, block_index: int,
                        total_blocks: int, is_hero: bool) -> dict:
        region_spec = block.get("region")
        if isinstance(region_spec, dict):
            return region_spec
        if isinstance(region_spec, str):
            return self.REGION_PRESETS.get(region_spec,
                                           self.REGION_PRESETS["full"])
        # Auto-compute
        return self._auto_region(block_index, total_blocks, is_hero)

    def _auto_region(self, block_index: int, total_blocks: int,
                     is_hero: bool) -> dict:
        if total_blocks == 1:
            return self.REGION_PRESETS["full"]
        elif total_blocks == 2:
            return self.REGION_PRESETS["left" if block_index == 0 else "right"]
        elif total_blocks == 3:
            if block_index == 0:
                return self.REGION_PRESETS["left-2-3"]
            return self.REGION_PRESETS["top-right" if block_index == 1 else "bottom-right"]
        elif total_blocks == 4:
            quadrant = ["top-left", "top-right", "bottom-left", "bottom-right"]
            return self.REGION_PRESETS[quadrant[block_index]]
        return self.REGION_PRESETS["full"]

    def _render_block(self, slide, block: dict, region: dict) -> None:
        block_type = block.get("type", "bullets")
        data = block.get("data", {})
        style = block.get("style", {})
        handler = getattr(self, f"_render_{block_type}_block", None)
        if handler:
            handler(slide, data, region, style)
        else:
            self._render_bullets_block(slide, data, region, style)
```

### 2.7 Section Divider Bug Fix

**Root cause**: `render_section_divider()` calls `self.add_slide(prs)` internally, creating a second slide. The first slide (already created by `render_slide()`) remains empty.

**Fix**: Refactor `render_section_divider()` to accept an existing slide instead of creating a new one.

```python
# BEFORE (buggy):
def render_section_divider(self, prs, section_number, section_title, section_subtitle=""):
    slide = self.add_slide(prs)  # Creates SECOND slide!
    # ... render content onto this slide ...
    return slide

# AFTER (fixed):
def render_section_divider(self, slide, section_number, section_title, section_subtitle=""):
    # Render directly onto the provided slide
    primary_hex = self._c("primary", "#2563EB")
    if self._is_dark():
        bg = self._lighten(primary_hex, 120)
    else:
        bg = self._lighten(primary_hex, 80)
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = self._rgb(bg)
    self.add_text(slide, f"{section_number:02d}", 2.0, 1.5, 9.333, 2.0,
                  size=72, color_role="primary", bold=True)
    self.add_text(slide, section_title, 2.0, 3.5, 9.333, 1.0,
                  size=40, color_role="foreground", bold=True)
    self.add_rect(slide, 2.0, 4.6, 3.0, 0.03, fill_role="accent", gradient=True)
    if section_subtitle:
        self.add_text(slide, section_subtitle, 2.0, 4.8, 9.333, 0.5,
                      font=self._font_b(), size=18, color_role="muted-foreground")
```

And in `render_slide()`:

```python
# BEFORE (buggy):
if is_section:
    section_num = page.get("section_number", page_index + 1)
    section_sub = subtitle or ""
    self.render_section_divider(prs, section_num, title, section_sub)
    return slide  # Returns the EMPTY first slide!

# AFTER (fixed):
if is_section:
    section_num = page.get("section_number", page_index + 1)
    section_sub = subtitle or ""
    self.render_section_divider(slide, section_num, title, section_sub)
    return slide  # Returns the slide with content
```

---

## 3. Design Examples: 3 Differentiated Proposals

With the blocks system, the same content can be rendered in 3 **structurally different** layouts:

### Proposal A: Dark Cyberpunk — Hexagon Flow

```json
{
  "goal": "content",
  "title": "4阶段智能流水线",
  "blocks": [
    {"type": "hexagons", "region": "left-2-3", "data": {
      "items": [
        {"label": "叙事规划", "color_role": "primary"},
        {"label": "设计决策", "color_role": "accent"},
        {"label": "内容生成", "color_role": "primary"},
        {"label": "PPT渲染", "color_role": "accent"}
      ]
    }},
    {"type": "bullets", "region": "right-1-3", "data": {
      "items": ["策略选择+情绪弧线", "40,000+风格匹配", "PAS/FAB/AIDA文案", "python-pptx直出+QA"]
    }}
  ]
}
```

**Visual**: 4 hexagons in honeycomb on left 2/3, bullet list on right 1/3. Impossible with old elif chain.

### Proposal B: Professional Tech — Sidebar + Pyramid

```json
{
  "goal": "content",
  "title": "三模式引擎",
  "blocks": [
    {"type": "bullets", "region": "sidebar", "data": {
      "items": ["01  FreeStyle", "02  Enterprise", "03  Build Script"],
      "bullet_style": "none"
    }, "style": {"font_size": 16, "color_role": "foreground"}},
    {"type": "diagram", "region": "right-2-3", "data": {
      "diagram_type": "pyramid",
      "diagram_data": {
        "levels": [
          {"label": "Build Script", "text": "精确交付"},
          {"label": "Enterprise", "text": "品牌合规"},
          {"label": "FreeStyle", "text": "快速探索"}
        ]
      }
    }}
  ]
}
```

**Visual**: Left sidebar with numbered items, pyramid diagram on right 2/3. Impossible with old elif chain.

### Proposal C: Elegant Dark — Centered Cards + Donut Metrics

```json
{
  "goal": "content",
  "title": "核心数据",
  "blocks": [
    {"type": "donuts", "region": "top", "data": {
      "items": [
        {"label": "40K+", "subtitle": "风格"},
        {"label": "28", "subtitle": "升级"},
        {"label": "10", "subtitle": "图形"},
        {"label": "824", "subtitle": "测试"}
      ],
      "layout": "horizontal"
    }},
    {"type": "cards", "region": "bottom", "data": {
      "items": [
        {"title": "Tier 1", "text": "OKLCH·阴影·渐变·品牌条"},
        {"title": "Tier 2", "text": "CJK·边距·徽章·分隔页"},
        {"title": "Tier 3", "text": "噪点·进度条·圆角·遮罩"}
      ],
      "layout": "horizontal"
    }}
  ]
}
```

**Visual**: 4 donut metrics across top, 3 cards across bottom. Impossible with old elif chain.

---

## 4. Implementation Plan

### Phase 1: Bug Fix + BlockRenderer Foundation

| Task | Files | Priority |
|------|-------|----------|
| Fix section divider bug | `precision_renderer.py` | P0 |
| Create `BlockRenderer` class with region presets | New file `block_renderer.py` | P0 |
| Implement core block types: `cards`, `bullets`, `diagram`, `code`, `exercise`, `image` | `block_renderer.py` | P0 |
| Modify `render_slide()` to check `blocks` key first | `precision_renderer.py` | P0 |
| Backward compatibility: old format → `_render_legacy_content()` | `precision_renderer.py` | P0 |

### Phase 2: New Block Types

| Task | Files | Priority |
|------|-------|----------|
| `hexagons` block | `block_renderer.py` | P1 |
| `ovals` block with connectors | `block_renderer.py` | P1 |
| `donuts` block | `block_renderer.py` | P1 |
| `metrics` block (big numbers) | `block_renderer.py` | P1 |
| `badge` block | `block_renderer.py` | P1 |
| `gradient_line` block | `block_renderer.py` | P1 |
| `table_chart` block | `block_renderer.py` | P1 |

### Phase 3: Auto-Region Computation

| Task | Files | Priority |
|------|-------|----------|
| Auto-region for 2-block layouts (left/right split) | `block_renderer.py` | P2 |
| Auto-region for 3-block layouts | `block_renderer.py` | P2 |
| Auto-region for 4-block layouts (2×2 grid) | `block_renderer.py` | P2 |
| Content-aware region sizing (measure text before placement) | `block_renderer.py` | P2 |

### Phase 4: Integration + Testing

| Task | Files | Priority |
|------|-------|----------|
| Pipeline passes `blocks` from content.json to render_slide | `pipeline.py` | P1 |
| Content parser supports `blocks` key | `content_parser.py` | P1 |
| Update SKILL.md with blocks format | `SKILL.md` | P2 |
| Unit tests for each block type | `tests/test_block_renderer.py` | P1 |
| Integration tests: 3 differentiated proposals | `tests/test_blocks_integration.py` | P1 |
| Regression: all 824 existing tests still pass | Full test suite | P0 |

### Phase 5: SKILL.md Update

| Task | Priority |
|------|----------|
| Default recommendation: use `blocks` in content.json | P2 |
| Document all 13 block types with examples | P2 |
| Document region presets | P2 |
| Update 5-step workflow: Step 2 proposals use different block compositions | P2 |

---

## 5. Risk Analysis

| Risk | Mitigation |
|------|------------|
| Breaking existing content.json files | `blocks` key is opt-in; old format falls back to `_render_legacy_content()` |
| Block overlap when regions conflict | Region validation: warn if blocks overlap, auto-adjust if possible |
| New block types lack design quality | Each new block reuses existing PrecisionRenderer primitives (add_hexagon, add_oval, etc.) |
| Auto-region produces bad layouts | Auto-region is a convenience; users can always specify explicit regions |
| Performance regression | BlockRenderer adds negligible overhead (just region computation + dispatch) |
| Test regression | Phase 4 mandates all 824 existing tests pass before merge |

---

## 6. Success Criteria

1. **3 proposals with different layouts** — same content, 3 structurally different visual outputs
2. **Zero breaking changes** — all existing content.json files work without modification
3. **Section divider bug fixed** — no more blank slides
4. **All 824 tests pass** — no regression
5. **Brand compliance preserved** — all blocks read from BrandSpec, CJK pairing works
6. **New block types functional** — hexagons, ovals, donuts, metrics render correctly
