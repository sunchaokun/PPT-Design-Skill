# SVGCompiler Development Plan

> Branch: `feat/svg-compiler` | Base: `main` | Date: 2026-08-14
> Status: **PLANNING** — no code changes to `src/` until Phase 1 tests pass

---

## 0. Problem Statement & Probe Foundation

### 核心痛点

PPT-Design-Skill 的 Build 模式承诺"100% native editable shapes, zero raster"。但现有
DiagramEngine 的 10 种图表类型（pyramid/funnel/venn/cycle/swot/...）全部用 MSO 原生形状
（RECTANGLE/OVAL）堆叠，无法表达：

- 梯形/三角形等非矩形几何（pyramid 实际是堆叠矩形，不是真金字塔）
- 渐变填充的任意多边形（funnel 的梯形段无法用 RECTANGLE 实现）
- 布尔运算形状（Venn 交集区域、clipPath 裁剪）
- 任意 Bezier 曲线（增长曲线、仪表盘弧线）

**核心矛盾**：LLM 擅长写声明式 SVG（path/bezier/gradient），但不擅长手算 python-pptx 的
EMU 坐标。需要一个编译层把 SVG 的声明式几何翻译成 python-pptx 的原生形状。

### 探针验证基础

**探针源码**：`scratch/svg_probe/probe.py` (795 行单文件) + `scratch/svg_probe/FINDINGS.md` (80 行报告)
**验证结论**：完全可行。基于现有 primitives（`freeform_builder`, `boolean_shapes`, `text_measurer`）可在 5–20 ms 内将任意复杂 SVG 编译为 100% 可编辑原生形状，零 `<p:pic>`。

**6 个验证用例指标**：

| 用例 | 形状数 | 可编辑性 | 特性支持 | ink-IoU | 编译耗时 |
|------|--------|----------|----------|---------|---------|
| pyramid | 6 | ✓ | gradient, polygon, text | 0.98 | 13 ms |
| venn_evenodd | 4 | ✓ | path(arc), text | 0.83 | 9 ms |
| funnel | 8 | ✓ | gradient, polygon, text | 0.98 | 9 ms |
| growth_curve | 5 | ✓ | clipPath, gradient, path(C), text | 0.91 | 13 ms |
| matrix_bcg | 13 | ✓ | rect/circle/line/group/text | 0.91 | 21 ms |
| unsupported | 2 | ✓ | image/filter/mask REFUSED | — | 3 ms |

**探针修复的边界情况**：
- **Arc full-circle 规避**：对 `a r,r 0 1,0` 这种终点与起点重合的极小距离弧，强制走 4 段 Cubic Bezier 保证渲染完美。
- **布尔运算拓扑硬化**：利用 Shapely `buffer(0)` 和 `make_valid` 解决 LLM 自交多边形崩溃问题。
- **渐变布尔两步法**：`boolean_shapes.py` 不支持渐变，编译时先以白色填充构建 geometry，再委托 `GradientFill.apply()` 替换填充。

---

## 1. Architecture: Three-Tier Routing

```
User Request (diagram_type / chart / svg_diagram)
       │
       ▼
  PrecisionRenderer.render_slide()
       │
       ├─ page["chart"] ──────────► ChartBuilder (Tier 1)
       │   条件: 结构化数据 (categories × series × values)
       │   产物: <c:chart> (Excel-backed, 数据可编辑)
       │   例: 柱状图、折线图、饼图、散点图
       │
       ├─ page["diagram_type"] ───► DiagramEngine (Tier 2)
       │   条件: 标准布局 + 简单几何 (rect/oval/connector)
       │   产物: MSO shapes + connectors
       │   例: 简单流程图、基础时间线、2-set Venn
       │
       └─ page["svg_diagram"] ───► SVGCompiler (Tier 3)  ← NEW
           条件: 自定义几何 / 布尔运算 / 长尾场景
           产物: custGeom + textbox + group (100% editable)
           例: 梯形金字塔、BCG 矩阵、仪表盘、增长曲线
```

### build.py + SVGCompiler 协同工作流

SVGCompiler 不是独立的工作流，而是 `build_helpers` 中的一个**组件级画笔**。两者统一在同一个 `Presentation` 对象中输出。

```python
# build_deepseek_bp.py (示例工作流)
from ppt_pro_max.build_helpers import *

C = {'primary': '#1D78FA', 'on_primary': '#FFFFFF', 'accent': '#FF5500', ...}
t = TYPOGRAPHY['mckinsey']

prs = Presentation(template_path)
s = add_slide(prs)
page_header(s, '战略演进', '三阶段路径', C=C, typo=t)

# 左侧：传统 Python 代码绘制侧边栏装饰 (Tier 2)
rect(s, 0.3, 1.5, 0.1, 5.5, fill='muted', C=C)
text(s, 0.6, 1.8, 2.0, 0.5, '演进阶段', font_size=t.h3, bold=True, color='primary', C=C)
multiline(s, 0.6, 2.4, 2.0, 3.6, [...], font_size=14, color='text_body', C=C)

# 右侧：SVG 编译复杂图表 (Tier 3)
evolution_svg = """<svg viewBox="0 0 400 300">
  <polygon points="200,30 380,290 20,290" fill="var(--primary)"/>
  <polygon points="200,120 320,290 80,290" fill="var(--accent)"/>
  <text x="200" y="85" text-anchor="middle" fill="var(--on-primary)">战略愿景</text>
</svg>"""
svg_chart(s, evolution_svg, x=3.5, y=1.5, w=6.3, h=5.5, C=C)

prs.save('output.pptx')
```

**关键设计**：`svg_chart(slide, svg_text, x, y, w, h, C=C)` 是 `build_helpers` 中的标准组件函数，与 `kpi_card()`、`native_chart()` 完全平级。同一份 `C` 字典同时驱动所有形状的颜色，无需修改 SVG 字符串。

### Routing Decision Table

| 场景 | Tier | 理由 |
|------|------|------|
| 柱状图/折线图/饼图 (有数据表) | 1 | ChartBuilder 提供 Excel 数据绑定 |
| 简单流程图 (3-6 步, 线性) | 2 | MSO 形状最简单，编辑性最好 |
| 基础 Venn (2-set, 无交集标注) | 2 | 两个 OVAL 足够 |
| 金字塔 (真梯形, 渐变) | 3 | 需要 polygon + gradientFill |
| 漏斗 (真梯形, 渐变) | 3 | 需要 polygon + gradientFill |
| BCG 矩阵 (虚线网格+气泡) | 3 | 需要 dash + circle + group |
| 增长曲线 (Bezier + clipPath) | 3 | 需要 cubicBezTo + clipPath |
| 仪表盘 (弧形 + 指针) | 3 | 需要 arc + boolean subtract |
| Venn (3-set, 交集标注) | 3 | 需要 evenodd boolean |
| 瀑布图 | 3 | 需要 polygon + 渐变 |
| 雷达图 (有数据) | 1 | ChartBuilder 原生支持 |
| 雷达图 (装饰性, 无数据) | 3 | 自定义多边形更灵活 |

### 关键原则：路由层级不在 LLM 决策

**反模式**：让 LLM 在生成 `pages.json` 时选择 Tier（"这段用 Tier 3"）。

**正确做法**：
- build.py 的 narrative generator 决定章节结构（"本章有 8 页：hook + 3 个 evolution + data + cta"）
- SVGCompiler 的存在是让 build.py 后续能"想画什么画什么"——LLM 不需要知道 Tier，只需要在 build.py 里把任意 SVG 片段交给 `svg_chart()`
- 路由决策完全由 build.py 的代码逻辑决定，不需要 LLM 干预

### Integration Point (PrecisionRenderer)

**行号来源**：`precision_renderer.py:686`

```python
# CURRENT:
            elif diagram_type and diagram_data:
                self._render_diagram_on_slide(slide, diagram_type, diagram_data, cx, cy, cw, ch)

# PLANNED (after merge):
            elif diagram_type and diagram_data:
                if page.get("svg_diagram"):          # ← new key
                    self._render_svg_diagram_on_slide(slide, page.get("svg_diagram"), cx, cy, cw, ch)
                else:
                    self._render_diagram_on_slide(slide, diagram_type, diagram_data, cx, cy, cw, ch)
```

**隔离保证**：`svg_diagram` 是新 key，现有 `diagram_type`/`diagram_data` 路径完全不受影响。
在测试通过前，SVGCompiler 只通过 `build_helpers` 的独立 API 调用，不接入 render_slide 调度。

---

## 2. Color Strategy: build_helpers 层统一配色

### 核心原则：配色不在编译器内部闭门造车

SVGCompiler 不持有自己的配色逻辑。所有颜色解析通过 `build_helpers._resolve_color` 和
`visual_effects` 模块完成，确保 SVG 编译出的形状与同一页面上其他形状使用完全相同的色值。

### 配色注入路径

```
build.py 定义 C 字典
    │
    ├─ rect(s, ..., fill='primary', C=C)  ──► _resolve_color('primary', C) ──► '#1D78FA'
    │
    └─ svg_chart(s, svg, ..., C=C)  ──► SVGCompiler(C=C)
                                            │
                                            ├─ fill="var(--primary)"  ──► C['primary'] ──► '#1D78FA'
                                            ├─ fill="accent"          ──► C['accent']  ──► '#FF5500'
                                            ├─ fill="currentColor"    ──► C['text_dark']
                                            └─ fill="#FF0000"         ──► 直接使用（硬编码色）
```

**结果**：用户在 build.py 开头修改 `C['primary'] = '#NewColor'`，所有 SVG 图表自动同步，
无需修改 SVG 字符串。导出的 PPT 中，右键形状 → 设置形状格式，看到的填充色值与页面其他形状完全一致。

### 颜色解析优先级（`_resolve_svg_color` 内部逻辑）

| 优先级 | SVG 输入 | 解析行为 | 示例 |
|--------|---------|---------|------|
| 1 | `var(--name)` | 查找 `C[name]`，未找到则查找 `C['palette_' + name]`，再未找到用 fallback | `var(--primary)` → `C['primary']` |
| 2 | C key 直传 | 如果值是 C 字典中的 key，直接解析 | `fill="accent"` → `C['accent']` |
| 3 | `currentColor` | 映射为 `C['text_dark']`（SVG 规范中 currentColor 继承自 color 属性，PPT 中最接近的语义是正文色） | `fill="currentColor"` → `C['text_dark']` |
| 4 | `hex` | 直接使用，不经过 C 字典 | `fill="#FF0000"` → `#FF0000` |
| 5 | `rgb()`/`rgba()` | 转换为 hex + alpha 通道 | `fill="rgb(255,0,0)"` → `#FF0000` |
| 6 | `hsl()`/`hsla()` | 转换为 hex + alpha 通道 | `fill="hsl(0,100%,50%)"` → `#FF0000` |
| 7 | SVG named color | 查找 CSS 4.1 named-color 表（141 色），转换为 hex | `fill="red"` → `#FF0000` |
| 8 | `url(#gradientId)` | 查找 `<defs>` 中的渐变定义，委托 `GradientFill` | `fill="url(#g1)"` → gradientFill |

### 拒绝策略（绝不静默降级）

| 输入 | 行为 | 理由 |
|------|------|------|
| `meshgradient` | `SVGCompileError` | PPT 无原生网格渐变，强制转换会视觉崩坏 |
| `lab()` / `lch()` / `oklch()` | `SVGCompileError` | 非 sRGB 色彩空间，PPT 不支持 |
| `color-mix()` | `SVGCompileError` | CSS 相对色函数，PPT 无对应 |
| 未知 named color | `SVGCompileError` | 避免静默使用黑色替代 |
| `var(--unknown)` 且 C 中无对应 key | `SVGCompileError` | 避免静默使用 fallback 替代 |

### Alpha 通道处理

PPT 原生支持 `srgbClr` + `alpha` 子节点，编译器将所有带透明度的颜色拆分为：

```xml
<!-- SVG: fill="#1D78FA" fill-opacity="0.45" -->
<a:srgbClr val="1D78FA">
  <a:alpha val="45000"/>  <!-- 0.45 × 100000 -->
</a:srgbClr>
```

这确保透明度在 PPT 中**可编辑**（右键 → 设置形状格式 → 透明度滑块），而非烘焙为不可编辑的位图。

### 渐变中的颜色解析

渐变 `<stop>` 的 `stop-color` 同样走 `_resolve_svg_color`：

```xml
<!-- SVG skeleton (LLM 生成) -->
<linearGradient id="g1" x1="0" y1="0" x2="1" y2="1">
  <stop offset="0" stop-color="var(--primary)"/>
  <stop offset="1" stop-color="var(--secondary)"/>
</linearGradient>

<!-- 编译时解析为 -->
<linearGradient id="g1" x1="0" y1="0" x2="1" y2="1">
  <stop offset="0" stop-color="#1D78FA"/>
  <stop offset="1" stop-color="#64748B"/>
</linearGradient>
```

### Font 注入

| SVG 输入 | 解析行为 |
|---------|---------|
| `font-family="var(--heading)"` | → `C['font_heading']` |
| `font-family="var(--body)"` | → `C['font_body']` |
| `font-family="Arial"` (硬编码) | 保持不变（品牌指定字体） |
| `font-family="sans-serif"` (generic) | → `C['font_body']`，无则 fallback Calibri |

---

## 3. Module Structure

### New Files (all under `src/ppt_pro_max/renderer/svg_compiler/`)

```
svg_compiler/
├── __init__.py              # public API: compile(), compile_to_slide()
├── _affine.py               # Affine transform class + parse_transform()
├── _path.py                 # SVG path parser: parse_path(), arc_to_cubics(), to_beziers()
├── _paint.py                # Paint resolution: solid/gradient/radial, style attr parsing
├── _text.py                 # Text rendering: baseline-aware placement, <tspan>
├── _sanitizer.py            # LLM SVG sanitization: fix unclosed tags, style→attrs, etc.
├── _compiler.py             # Core SVGCompiler class (walk/compile/render)
└── _dash.py                 # stroke-dasharray → a:prstDash / a:custDash mapping
```

### Why a Package, Not a Single File

1. **probe.py 是 795 行的单文件**，生产化后预计 1200-1500 行，拆分是必要的
2. **各子模块可独立测试**：path parser、affine transform、paint resolver 都有独立的测试面
3. **渐进式开发**：Phase 1 只需 `_affine` + `_path` + `_compiler`，其他模块后续追加

### Dependencies (No New Packages)

| Module | Depends On | Notes |
|--------|-----------|-------|
| `_affine` | stdlib only (math, re) | Pure math |
| `_path` | stdlib only (math, re) | Pure math |
| `_paint` | `visual_effects.py` (GradientFill) | Delegate gradient XML |
| `_text` | `text_measurer.py`, Pillow (Phase 2) | Phase 1 uses estimate_text_size |
| `_sanitizer` | lxml (etree) | XML fixup |
| `_compiler` | `freeform_builder.py`, `boolean_shapes.py`, `group_builder.py` | Core |
| `_dash` | lxml (etree) | OOXML dash XML |

| `_compiler` | `freeform_builder.py`, `boolean_shapes.py`, `group_builder.py` | Core |
| `_dash` | lxml (etree) | OOXML dash XML |

**关键原则**：复用现有生产模块，不重新实现。probe 中重复的 `_apply_gradient`、
`_apply_alpha`、`normalize_hex` 全部替换为 `visual_effects.py` 和 `build_helpers.py`
中的对应函数。

### Existing Primitives Reuse

| Probe Code | Production Equivalent | Action |
|-----------|----------------------|--------|
| `Affine` class | None (probe-only) | Promote to `_affine.py` |
| `parse_transform` | None | Promote to `_affine.py` |
| `parse_path` / `to_beziers` | None | Promote to `_path.py` |
| `arc_to_cubics` | None | Promote to `_path.py` |
| `flatten_bezier` | N/A | **Remove**: use `cubic_bezier_to` instead |
| `normalize_hex` | `_resolve_color` in `build_helpers.py` | Delegate |
| `_apply_gradient` | `GradientFill.apply()` in `visual_effects.py` | Delegate |
| `_apply_alpha` | `set_solid_fill_with_alpha()` in `visual_effects.py` | Delegate |
| `_add_native` | `slide.shapes.add_shape()` | Use directly |
| `_add_freeform` | `FreeformBuilder.build()` | Use directly |
| `_add_boolean` | `bool_shape()` in `boolean_shapes.py` | **Caution**: no gradient support (see below) |
| `estimate_text_size` | `text_measurer.estimate_text_size` | Phase 1: use directly; Phase 2: replace with Pillow |
| Group shapes | `group_builder.py:33` `GroupBuilder.build_from_shapes()` | Use directly |

---

## 4. Development Phases

### Phase 1: Core Compiler (MVP)

**目标**：将 probe 的编译能力移植到 `src/`，达到与 probe 等价的质量。

**产出**：
- `svg_compiler/_affine.py` — 从 probe 移植，加类型注解和 docstring
- `svg_compiler/_path.py` — 从 probe 移植，加错误处理
- `svg_compiler/_compiler.py` — 核心编译器，对接 FreeformBuilder/boolean_shapes
- `svg_compiler/_sanitizer.py` — 基础 XML 修复与结构化规范（提供 Phase 1 的错误隔离）
- `svg_compiler/__init__.py` — 公开 API
- `tests/test_svg_compiler.py` — 6 个 probe case 的回归测试

**关键改进（vs probe）**：
1. **cubic_bezier_to 替代 flatten_bezier**：对于原生 SVG 编译路径（非 boolean 路径），直接使用 `FreeformBuilder.cubic_bezier_to()`，圆形只需 4 个 cubic 段（4 个编辑点），编辑体验从 64 点变为正常。
2. **委托 visual_effects**：probe 自己写 `_apply_gradient` 和 `_apply_alpha`，生产版委托给 `GradientFill.apply()` 和 `set_solid_fill_with_alpha()`。**注意**：`boolean_shapes.py` 内置的 `bool_shape()` 只支持 `solidFill`（无渐变）。编译时沿用两步走策略：用白色填充创建 boolean 形状，再调用 `GradientFill.apply()` 替换填充。这是 `build_helpers.py:bool_shape` 签名限制导致的已知约束。
3. **错误处理与 Sanitizer 提前**：为应对 LLM 生成的不完整 XML，在 Phase 1 引入基础 `_sanitizer.py` 确保 XML 解析不崩溃。
4. **编译输出控制**：`SVGCompiler` 返回结构化的结果集，对于不支持的特性抛出 `SVGCompileError`。

**验收标准**：
- [ ] 6 个 probe case 全部编译成功（pyramid/venn/funnel/growth_curve/matrix_bcg/unsupported）
- [ ] 产物 XML 零 `<p:pic>`
- [ ] pyramid/funnel ink-IoU ≥ 0.95
- [ ] 编译圆形的编辑顶点数 ≤ 8 个（4 个 cubic 段）
- [ ] 编译时间 < 30ms/chart
- [ ] `pytest tests/test_svg_compiler.py` 全绿
- [ ] `ruff check src/ppt_pro_max/renderer/svg_compiler/` 零错误

### Phase 2: Commercial Quality

**目标**：修复 probe 的已知质量问题，达到商用交付标准。

**产出**：
- `svg_compiler/_text.py` — Baseline-aware text rendering
- `svg_compiler/_paint.py` — Paint resolution + radial gradient
- `svg_compiler/_dash.py` — stroke-dasharray support
- `tests/test_svg_text.py`, `tests/test_svg_paint.py`, `tests/test_svg_dash.py`

**关键改进**：

| Gap | Fix | Impact |
|-----|-----|--------|
| Text baseline 偏移 (IoU 0.83) | 尝试使用 Pillow `ImageFont.getbbox()` 获取精确的 ascent，转换 SVG `y` -> textbox top；CI headless 环境无字体时，优雅降级为 `estimate_text_size` 估算 | IoU → 0.95+ |
| 无 Group 输出 | 调用 `slide.shapes.add_group_shape(shapes)` 把属于同一个 `<g>` 的原生子形状打包成 PPT 组合 | 图表可整体移动/缩放 |
| 无虚线 | 将 `stroke-dasharray` 转换并应用到 PPT 的 `a:ln/a:prstDash` 中（支持常见虚线预设） | 网格线/参考线正确渲染 |
| 无径向渐变 | `radialGradient` 转换为 OOXML `a:path="circle"` 渐变（通过 `GradientFill(gradient_type="path", fill_to_rect={...})`） | 饼图/聚光灯效果 |
| `<tspan>` 丢失 | 逐 `<tspan>` 渲染，支持 `dx`/`dy` 偏移 | 多行文本/上下标 |
| `font-weight`/`font-style` 忽略 | 映射到 `run.font.bold` / `run.font.italic` | 文字粗细正确 |

**验收标准**：
- [ ] venn_evenodd ink-IoU ≥ 0.93（baseline 修复后）
- [ ] matrix_bcg 虚线正确渲染
- [ ] 所有 `<g>` 输出为 `<p:grpSp>` 组合形状
- [ ] `pytest tests/` 全绿（含新增测试）
- [ ] `ruff check` 零错误

### Phase 3: Style System Integration & build_helpers

**目标**：SVGCompiler 接入 40,000+ 风格组合系统及 `build_helpers` 模块。

**产出**：
- `svg_compiler/_theme.py` — 语义色 token → palette 映射
- PrecisionRenderer 集成（`svg_diagram` key dispatch）
- `tests/test_svg_theme.py`, `tests/test_svg_integration.py`

**关键设计**：

```python
# SVG skeleton (LLM 生成，不含具体颜色)
<svg viewBox="0 0 400 300">
  <polygon points="..." fill="var(--primary)"/>
  <text x="200" y="85" fill="var(--on-primary)">战略愿景</text>
</svg>

# Compiler 接收 C (Context) dict
compiler = SVGCompiler(C={
    "palette_primary": "#1D78FA",
    "palette_on_primary": "#FFFFFF",
    "font_heading": "Georgia",
    "font_body": "Calibri",
    ...
})
compiler.compile(svg_text, slide, rect=(3.5, 0.8, 6.3, 5.6))
```

**Token 映射表**：

| SVG Token | C Key | Fallback |
|-----------|-------|----------|
| `var(--primary)` | `palette_primary` | `#1D78FA` |
| `var(--secondary)` | `palette_secondary` | `#64748B` |
| `var(--accent)` | `palette_accent` | `#FF5500` |
| `var(--on-primary)` | `palette_on_primary` | `#FFFFFF` |
| `var(--background)` | `palette_background` | `#FFFFFF` |
| `var(--foreground)` | `palette_foreground` | `#1E293B` |
| `var(--muted)` | `palette_muted` | `#F1F5F9` |
| `var(--border)` | `palette_border` | `#E2E8F0` |

**Font 注入**：
- SVG `font-family="var(--heading)"` → `C["font_heading"]`
- SVG `font-family="var(--body)"` → `C["font_body"]`
- SVG `font-family="Arial"` (硬编码) → 保持不变（品牌指定字体）

**Mood Effects**：
- `theme_composer.compose()` 返回 `text_effect_preset` 和 `image_effect`
- 编译后对形状批量应用 mood 对应的 shadow/glow/3D 效果

**验收标准**：
- [ ] 同一 SVG skeleton + 3 种不同 C → 3 种视觉风格（颜色/字体/效果不同）
- [ ] 现有 1,377 个测试不受影响
- [ ] `pytest tests/ -q` 全绿
- [ ] E2E: FreeStyle 模式生成含 SVG 图表的 PPT，风格正确

### Phase 4: Merge to Main

**前置条件**：
1. Phase 1-3 全部验收通过
2. `pytest tests/ -q` 全绿（含原有 1,377 测试 + 新增测试）
3. `ruff check src/` 零错误
4. E2E 测试：FreeStyle + Enterprise + Build 三种模式各生成一个含 SVG 图表的 PPT
5. 人工审计：在 PowerPoint 中打开产物，确认可编辑性

**合并步骤**：
1. `git checkout main && git merge feat/svg-compiler --no-ff`
2. 合并后运行完整回归测试
3. 更新 AGENTS.md 和 ARCHITECTURE.md
4. Tag: `v0.18.0`

---

## 5. Isolation Guarantees

### Code Isolation

| 保证项 | 机制 |
|--------|------|
| 新代码不修改现有文件 | Phase 1-2 只新增 `svg_compiler/` 包，不修改 `precision_renderer.py` 等 |
| 新代码不修改现有测试 | 新测试文件独立：`test_svg_compiler.py` 等 |
| 新代码不引入新依赖 | 只用已有的 python-pptx/shapely/lxml/Pillow |
| 新代码不影响现有流程 | `svg_diagram` 是新 key，现有 `diagram_type` 路径零影响 |

### Branch Strategy

```
main ──────────────────────────────────────────────►
  │
  └─ feat/svg-compiler ────────────────────────────►
       │
       ├─ svg-p1 commit: _affine + _path + _compiler + _sanitizer + tests
       ├─ svg-p2 commit: _text + _dash + _paint + tests
       ├─ svg-p3 commit: _theme + PrecisionRenderer integration + tests
       └─ Phase 4: merge PR → main
```

- 每个 Phase 完成后打 tag: `svg-p1`, `svg-p2`, `svg-p3`
- 每个 Phase 的 commit 必须包含对应测试
- 合并前必须通过完整回归测试

### Rollback Plan

如果合并后发现问题：
1. `git revert` 合并 commit
2. `svg_compiler/` 包整体删除，零残留（因为只新增文件，不修改现有文件）
3. `precision_renderer.py` 中的 `svg_diagram` dispatch 代码一并 revert

---

## 5. Test Strategy

### Unit Tests (per module)

| Test File | Covers | Cases |
|-----------|--------|-------|
| `test_svg_affine.py` | `_affine.py` | compose, apply, parse_transform (translate/scale/rotate/matrix/nested) |
| `test_svg_path.py` | `_path.py` | parse_path (M/L/H/V/C/S/Q/T/A/Z), arc_to_cubics, to_beziers, edge cases |
| `test_svg_sanitizer.py` | `_sanitizer.py` | LLM quirks: unclosed tags, style→attrs, missing viewBox |
| `test_svg_compiler.py` | `_compiler.py` | 6 probe cases + editability audit + compile timing + cubic edit-point count |
| `test_svg_text.py` | `_text.py` | baseline placement, CJK, text-anchor, <tspan>, font-weight/style |
| `test_svg_paint.py` | `_paint.py` | solid/linear/radial gradient, fill-opacity, var(--primary) token resolution |
| `test_svg_dash.py` | `_dash.py` | stroke-dasharray → prstDash mapping |
| `test_svg_theme.py` | `_theme.py` | var(--primary) → palette resolution, font injection, 3-style output |
| `test_svg_integration.py` | E2E | SVGCompiler + PrecisionRenderer + theme_composer |

### Regression Guard

```python
# tests/test_svg_compiler.py — core regression
PROBE_CASES = {
    "pyramid": {...}, "venn_evenodd": {...}, "funnel": {...},
    "growth_curve": {...}, "matrix_bcg": {...}, "unsupported": {...},
}

@pytest.mark.parametrize("name,svg", PROBE_CASES.items())
def test_compile_no_pictures(name, svg):
    """Every compiled output must have zero <p:pic>."""
    ...

@pytest.mark.parametrize("name,svg,expected_iou", [
    ("pyramid", ..., 0.95),
    ("funnel", ..., 0.95),
    ("growth_curve", ..., 0.90),
    ("matrix_bcg", ..., 0.90),
])
def test_geometry_fidelity(name, svg, expected_iou):
    """Ink-IoU must meet threshold."""
    ...

def test_unsupported_features_refused():
    """image/filter/mask must raise SVGCompileError, not silently degrade."""
    ...
```

### Existing Test Guard

每次 commit 前运行：
```bash
python -m pytest tests/ -q --ignore=tests/test_group_audit.py \
  --ignore=tests/test_image_fetcher.py --ignore=tests/test_pptx_capabilities.py \
  --ignore=tests/test_xml_extraction.py --ignore=tests/test_analyze_template.py
```

确保 1,377 个现有测试不受影响。

---

## 6. API Design

### Public API (svg_compiler/__init__.py)

```python
class SVGCompiler:
    """Compile SVG subset to native editable pptx shapes.

    Usage:
        compiler = SVGCompiler()                          # Phase 1-2
        compiler = SVGCompiler(C=context_dict)            # Phase 3
        result = compiler.compile(svg_text, slide, rect)
        # result = SVGResult(shapes=[...], warnings=[...], features={...})

    Guarantees:
        - Zero <p:pic> in output (never silently degrades to picture)
        - All shapes are native editable (custGeom / MSO / textbox)
        - Unsupported features raise SVGCompileError
    """

    def __init__(self, C: dict | None = None):
        self.C = C or {}

    def compile(self, svg_text: str, slide, rect: tuple[float, float, float, float],
                vb: tuple[float, float, float, float] | None = None) -> SVGResult:
        """Compile SVG to native shapes on slide.

        Args:
            svg_text: SVG source string
            slide: python-pptx Slide object
            rect: (x, y, w, h) in inches — target placement on slide
            vb: (x, y, w, h) — SVG viewBox override (auto-detected if None)

        Returns:
            SVGResult with shapes list, warnings, and features set

        Raises:
            SVGCompileError: if SVG contains unsupported features that would
                require degrading to picture (image/filter/mask)
        """

@dataclass
class SVGResult:
    shapes: list          # list of shape elements created
    warnings: list[str]   # non-fatal issues (skipped elements, approximations)
    features: set[str]    # SVG features detected during compilation
    compile_ms: float     # compilation time in milliseconds

class SVGCompileError(Exception):
    """Raised when SVG contains features that cannot compile to editable shapes."""
```

### build_helpers Integration (Phase 3)

```python
# build_helpers.py — new convenience function
def svg_chart(slide, svg_text, x, y, w, h, C=None):
    """Render an SVG diagram as native editable shapes.

    Example:
        svg_chart(slide, pyramid_svg, 3.5, 0.8, 6.3, 5.6, C=C)
    """
    from ppt_pro_max.renderer.svg_compiler import SVGCompiler
    compiler = SVGCompiler(C=C)
    return compiler.compile(svg_text, slide, (x, y, w, h))
```

---

## 7. Risk Register

| Risk | Probability | Impact | Mitigation |
|------|------------|--------|------------|
| Shapely import fails (no GEOS) | Low | High | `HAS_SHAPELY` guard; boolean features gracefully disabled |
| Pillow not installed | Low | Medium | Phase 1 uses `text_measurer.estimate_text_size`; Phase 2 requires Pillow |
| LLM generates invalid SVG | High | Medium | `_sanitizer.py` fixes common issues; `SVGCompileError` for unfixable |
| PowerPoint renders custGeom differently | Medium | Low | Pixel-compare tests catch regressions |
| Merge conflicts with main | Low | Medium | Feature branch; frequent rebase on main |
| Performance regression (many shapes) | Low | Medium | Shape count budget; group consolidation |

---

## 8. Success Metrics

| Metric | Target | Measurement |
|--------|--------|-------------|
| Editability | 0 `<p:pic>` per chart | XML audit in test |
| Geometry fidelity | ink-IoU ≥ 0.90 (all cases) | Pixel compare vs ground truth |
| Compile speed | < 50ms per chart | time.perf_counter in test |
| Test coverage | ≥ 80% lines in svg_compiler/ | pytest-cov |
| Existing tests | 0 regressions | Full test suite green |
| Lint | 0 errors | ruff check |

---

## 9. Decision Log

| # | Decision | Rationale | Date |
|---|----------|-----------|------|
| D1 | SVGCompiler 作为 `build_helpers` 的**组件级画笔**而非独立流水线 | 配色统一：build.py 的 `C` 字典同时驱动 SVG 形状和其他原生形状，避免视觉断层；零集成摩擦 | 2026-08-15 |
| D2 | LLM 不参与 Tier 路由决策 | build.py 决定章节结构 → 调用 `svg_chart()`，路由在代码层完成。LLM 不需要知道 Tier 概念 | 2026-08-15 |
| D3 | 颜色解析在 `build_helpers._resolve_color` 层完成 | 避免 SVGCompiler 与 `_resolve_color` 重复实现；统一所有形状（MSO/Connector/CustGeom）的色值来源 | 2026-08-15 |
| D4 | 拒绝静默降级（unsupported → raise） | "PPT 看起来对"但可编辑性破损 = 核心承诺违约。宁可编译失败让人修 SVG，也不能输出 `<p:pic>` | 2026-08-13 |
| D5 | Boolean shapes 必须用 solid 填充后再委托 `GradientFill` | `bool_shape()` 签名限制导致。已验证两步走视��无差 | 2026-08-13 |
| D6 | 圆形用 4 段 Cubic Bezier 而非 64 段直线 | 视觉无差，编辑顶点数从 64 降到 4，大幅改善用户体验 | 2026-08-13 |
| D7 | `PrecisionRenderer` 新增 `svg_diagram` key，但旧 `diagram_type`/`diagram_data` 路径**完全保留** | 零现有代码修改；现有调用零迁移成本；新功能仅作为可选分支 | 2026-08-14 |
| D8 | SVG skeleton 由 LLM 生成，不直接给最终用户改 | LLM 写 path/bezier 比手算 EMU 可靠 100 倍；用户改的是 `C` 字典和文本 | 2026-08-15 |

---

## 10. Open Questions

- [ ] **OQ1**: `<use>` 元素的处理优先级？Phase 1 直接拒绝（raise），Phase 2 优化为 inline 展开
- [ ] **OQ2**: `text` + `<animate>` 的 fallback？LLM 极少生成动画 SVG，但需决定是 raise 还是忽略 animate 节点
- [ ] **OQ3**: 多 `<svg>` 根节点的合并策略？当前 sanitizer 只接受单 SVG 根，多 SVG 抛错
- [ ] **OQ4**: `viewBox` 缺失时如何计算 fallback？当前用 0,0,viewBox.width 推断，但页面放置 rect 不可见时怎么办

---

## 11. References

- **Probe source**: `scratch/svg_probe/probe.py`
- **Probe findings**: `scratch/svg_probe/FINDINGS.md`
- **Probe report**: `scratch/svg_probe/out/REPORT.md`
- **Existing primitives**:
  - `src/ppt_pro_max/build_helpers.py:89` — `_resolve_color`
  - `src/ppt_pro_max/build_helpers.py:2173` — `set_theme_colors`
  - `src/ppt_pro_max/renderer/visual_effects.py` — `GradientFill`, `set_solid_fill_with_alpha`
  - `src/ppt_pro_max/renderer/freeform_builder.py` — `FreeformBuilder.cubic_bezier_to`
  - `src/ppt_pro_max/renderer/boolean_shapes.py` — `bool_shape`
  - `src/ppt_pro_max/renderer/text_measurer.py` — `estimate_text_size`
  - `src/ppt_pro_max/renderer/theme_composer.py:724` — `compose()`
- **Integration target**: `src/ppt_pro_max/enterprise/precision_renderer.py:686`
- **python-pptx reference**: `src/ppt_pro_max/docs/python-pptx-reference.md`
