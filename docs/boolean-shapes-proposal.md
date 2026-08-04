# Boolean Shape Operations — Design Proposal v3

## Core Problem

LLM 只用矩形，不是因为文档不够，而是因为 **build_helpers.py 只暴露了 3 个形状函数**：
- `rect()` → RECTANGLE
- `rrect()` → ROUNDED_RECTANGLE  
- `oval()` → OVAL

python-pptx 有 170+ 形状，但 LLM 写 `from ppt_pro_max.build_helpers import *` 后
只能用这 3 个。它不会主动去写 `slide.shapes.add_shape(MSO_SHAPE.HEXAGON, ...)`。

## Solution: Two Changes

### Change 1: 补全形状函数（解决"LLM 只用矩形"的根本问题）

#### 1a. 通用 `shape()` 函数

```python
def shape(slide, shape_type, left, top, width, height, fill, line=None, C=None):
    """Add any MSO_SHAPE. Same API as rect()/oval() but for all 170+ shapes.
    
    shape_type: MSO_SHAPE enum value or string name
    Examples:
        shape(slide, MSO_SHAPE.HEXAGON, 2, 2, 3, 2.6, '#1D78FA')
        shape(slide, 'STAR_5_POINT', 5, 3, 2, 2, C['accent'])
        shape(slide, MSO_SHAPE.DONUT, 1, 1, 3, 3, '#FF5500')
    """
```

这一行就让 LLM 能用所有 170+ 形状。

#### 1b. 常用形状便捷函数（和 rect/oval 同级）

```python
# Polygons
def hexagon(slide, cx, cy, size, fill, line=None, C=None): ...
def pentagon(slide, cx, cy, size, fill, line=None, C=None): ...
def octagon(slide, cx, cy, size, fill, line=None, C=None): ...
def diamond(slide, cx, cy, size, fill, line=None, C=None): ...

# Triangles
def triangle(slide, left, top, width, height, fill, line=None, C=None): ...

# Stars (5/6/8/10/12 point)
def star5(slide, cx, cy, size, fill, line=None, C=None): ...
def star6(slide, cx, cy, size, fill, line=None, C=None): ...
def star8(slide, cx, cy, size, fill, line=None, C=None): ...

# Special shapes
def donut(slide, cx, cy, size, fill, line=None, C=None): ...
def heart(slide, cx, cy, size, fill, line=None, C=None): ...
def cross(slide, cx, cy, size, fill, line=None, C=None): ...
def arrow(slide, left, top, width, height, fill, line=None, C=None): ...
def chevron(slide, left, top, width, height, fill, line=None, C=None): ...
def cloud(slide, left, top, width, height, fill, line=None, C=None): ...
def lightning(slide, left, top, width, height, fill, line=None, C=None): ...
def gear(slide, cx, cy, size, fill, line=None, C=None, teeth=6): ...
def funnel(slide, left, top, width, height, fill, line=None, C=None): ...

# Callouts
def callout(slide, left, top, width, height, fill, line=None, C=None, style='rect'): ...

# Flowchart
def flow_process(slide, left, top, width, height, fill, line=None, C=None): ...
def flow_decision(slide, cx, cy, size, fill, line=None, C=None): ...
def flow_data(slide, left, top, width, height, fill, line=None, C=None): ...
```

约 25 个便捷函数，覆盖 PPT 设计中最常用的形状。

#### 1c. 形状图片裁剪（扩展 circle_image）

```python
def hex_image(slide, cx, cy, size, image_path, border_color=None): ...
def star_image(slide, cx, cy, size, image_path, points=5, border_color=None): ...
def diamond_image(slide, cx, cy, size, image_path, border_color=None): ...
def heart_image(slide, cx, cy, size, image_path, border_color=None): ...
def shape_image(slide, shape_type, left, top, width, height, image_path, border_color=None): ...
```

### Change 2: 布尔形状函数（解决"预设形状不够用"的问题）

当 MSO_SHAPE 的 170+ 形状仍然不够时（比如需要偏心孔甜甜圈、切角卡片、聚光灯遮罩），
布尔运算提供自定义形状能力。

```python
# Overlay effects
def spotlight(slide, cx, cy, radius, alpha=70, color='#000000'): ...

# Boolean-derived shapes (需要 Shapely，有 fallback)
def bool_donut(slide, cx, cy, outer_r, inner_r, fill, line=None, C=None): ...
def bool_frame(slide, x, y, w, h, border, fill=None, line=None, C=None): ...
def bool_star(slide, cx, cy, r, points=5, inner_ratio=0.4, fill=None, line=None, C=None): ...
def bool_cross(slide, cx, cy, w, h, bar_ratio=0.33, fill=None, line=None, C=None): ...
def bool_clipped_card(slide, x, y, w, h, clip_corners, clip_size=0.3, fill=None, line=None, C=None): ...
def bool_neon_tube(slide, x, y, w, h, wall=0.06, fill=None, C=None): ...

# Boolean primitives (LLM 进阶用)
def bool_subtract(a, b): ...
def bool_union(*shapes): ...
def bool_intersect(a, b): ...
def poly_rect(x, y, w, h): ...
def poly_circle(cx, cy, r): ...
def poly_rounded_rect(x, y, w, h, radius): ...
def poly_star(cx, cy, r, points, inner_ratio): ...
def bool_shape(geometry, slide, x, y, w, h, fill=None, line=None, C=None): ...
def bool_image(geometry, slide, x, y, w, h, image_path, border_color=None): ...
```

## LLM 形状使用路径

```
Level 1 (最简单):  rect(), oval(), rrect(), hexagon(), star5(), donut(), ...
                   → 和现在一样，直接调函数

Level 2 (通用):    shape(slide, MSO_SHAPE.XXX, x, y, w, h, fill)
                   → 170+ 形状，一个函数全覆盖

Level 3 (布尔便捷): spotlight(), bool_donut(), bool_frame(), bool_clipped_card()
                   → 预设形状做不到的效果

Level 4 (布尔原语): poly_rect() + bool_subtract() + bool_shape()
                   → 完全自定义形状
```

## 参考文档策略

### 新建: `src/ppt_pro_max/docs/shapes-reference.md`

包含：
1. 所有便捷函数的签名和示例
2. `shape()` 函数 + MSO_SHAPE 完整分类表（从 python-pptx-reference.md 精简）
3. 布尔形状函数的签名和示例
4. 10+ 完整示例代码（不同场景用不同形状）

### SKILL.md 变更

在 "Functions — Shapes" 部分扩展为：

```markdown
### Functions — Shapes

| Function | Purpose | Key Params |
|----------|---------|------------|
| `rect(slide, left, top, width, height, fill, line, C)` | Rectangle | fill/line: role name or hex |
| `rrect(slide, left, top, width, height, fill, line, C)` | Rounded rectangle | Same as rect |
| `oval(slide, left, top, width, height, fill, line, C)` | Ellipse | Same as rect |
| `shape(slide, shape_type, left, top, width, height, fill, line, C)` | **Any MSO_SHAPE** | shape_type: enum or string name |
| `hexagon(slide, cx, cy, size, fill, line, C)` | Hexagon | Center-based positioning |
| `pentagon(slide, cx, cy, size, fill, line, C)` | Pentagon | Center-based |
| `diamond(slide, cx, cy, size, fill, line, C)` | Diamond | Center-based |
| `triangle(slide, left, top, width, height, fill, line, C)` | Triangle | Corner-based |
| `star5(slide, cx, cy, size, fill, line, C)` | 5-point star | Center-based |
| `donut(slide, cx, cy, size, fill, line, C)` | Donut/ring | Center-based |
| `heart(slide, cx, cy, size, fill, line, C)` | Heart | Center-based |
| `cross(slide, cx, cy, size, fill, line, C)` | Cross/plus | Center-based |
| `arrow(slide, left, top, width, height, fill, line, C)` | Right arrow | Corner-based |
| `chevron(slide, left, top, width, height, fill, line, C)` | Chevron | Corner-based |
| `cloud(slide, left, top, width, height, fill, line, C)` | Cloud shape | Corner-based |
| `gear(slide, cx, cy, size, fill, line, C, teeth)` | Gear | teeth: 6 or 9 |
| `funnel(slide, left, top, width, height, fill, line, C)` | Funnel | Corner-based |
| `callout(slide, left, top, width, height, fill, line, C, style)` | Callout bubble | style: 'rect'/'round'/'cloud' |

### Functions — Shape Image Cropping

| Function | Purpose | Key Params |
|----------|---------|------------|
| `circle_image(slide, cx, cy, radius, image_path, border_color)` | Circle-cropped image | Center + radius |
| `hex_image(slide, cx, cy, size, image_path, border_color)` | Hexagon-cropped image | Center + size |
| `star_image(slide, cx, cy, size, image_path, points, border_color)` | Star-cropped image | points: 5/6/8 |
| `diamond_image(slide, cx, cy, size, image_path, border_color)` | Diamond-cropped image | Center + size |
| `shape_image(slide, shape_type, left, top, width, height, image_path, border_color)` | **Any shape** image crop | shape_type: MSO_SHAPE or string |

### Functions — Boolean Shapes

Advanced shapes via boolean operations. See **[`shapes-reference.md`](src/ppt_pro_max/docs/shapes-reference.md)** for full API and examples.

| Function | Purpose | Key Params |
|----------|---------|------------|
| `spotlight(slide, cx, cy, radius, alpha, color)` | Dark overlay + bright window | alpha=70 |
| `bool_donut(slide, cx, cy, outer_r, inner_r, fill, line, C)` | Donut with custom hole | Off-center capable |
| `bool_frame(slide, x, y, w, h, border, fill, line, C)` | Frame/border shape | border: inches |
| `bool_clipped_card(slide, x, y, w, h, clip_corners, clip_size, fill, line, C)` | Card with clipped corners | clip_corners: ['tl','tr','bl','br'] |
| `bool_neon_tube(slide, x, y, w, h, wall, fill, C)` | Hollow neon tube | wall: inches |
```

约 35 行新增，SKILL.md 从 1396 → ~1431 行，可接受。

## Implementation Phases

### Phase 1: 补全基础形状函数（解决根本问题）

| Item | Description |
|------|-------------|
| `shape()` 通用函数 | 一行代码用任何 MSO_SHAPE |
| 25 个便捷函数 | hexagon/pentagon/diamond/triangle/star5/donut/heart/cross/arrow/chevron/cloud/lightning/gear/funnel/callout/flow_process/flow_decision/flow_data/... |
| 5 个图片裁剪函数 | hex_image/star_image/diamond_image/heart_image/shape_image |
| shapes-reference.md | 完整形状参考 + 示例代码 |
| SKILL.md 更新 | Functions — Shapes 表格扩展 |
| 测试 | 每个函数至少 1 个测试 |

### Phase 2: 布尔形状函数（锦上添花）

| Item | Description |
|------|-------------|
| boolean_shapes.py | Shapely → custGeom 内部模块 |
| 6 个布尔便捷函数 | spotlight/bool_donut/bool_frame/bool_clipped_card/bool_neon_tube/bool_star |
| 布尔原语 | bool_subtract/union/intersect + poly_* + bool_shape/bool_image |
| fallback | 每个函数无 Shapely 时降级到预设形状 |
| shapes-reference.md 扩展 | 布尔形状示例 |
| 测试 | 27 个已有 + 新增 |

### Phase 3: 文字轮廓提取（可延后）

| Item | Description |
|------|-------------|
| polygon_from_text() | fontTools → Shapely polygon |
| CJK 支持 | 中文/日文/韩文字符轮廓 |
| 集成 | spotlight 文字亮区 / seal_stamp 雕刻文字 |

## Impact Analysis

### Before (现状)

LLM 可用形状函数: **3 个** (rect, rrect, oval)
LLM 实际使用: 矩形为主，偶尔 oval

### After (Phase 1 完成后)

LLM 可用形状函数: **30+ 个** (rect/rrect/oval + shape() + 25 便捷函数)
LLM 可用图片裁剪: **6 个** (circle_image + hex/star/diamond/heart/shape_image)
LLM 可用 MSO_SHAPE: **170+** (通过 shape() 通用函数)

### After (Phase 2 完成后)

LLM 可用布尔形状: **6 个便捷 + 4 个原语**
LLM 自定义形状: **无限** (通过布尔原语组合)

## Key Insight

**补全基础形状函数（Phase 1）比布尔运算（Phase 2）更重要。**

原因：
- Phase 1 解决"LLM 只用矩形"的根本问题 — 从 3 个函数扩展到 170+ 形状
- Phase 2 解决"170+ 形状仍然不够"的进阶问题 — 布尔运算创造新形状
- 大多数 PPT 设计用 Phase 1 的形状就够了
- Phase 2 是差异化竞争力，但不是刚需

建议先做 Phase 1，验证 LLM 确实开始使用更多形状后，再做 Phase 2。
