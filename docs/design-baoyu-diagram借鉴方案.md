# baoyu-diagram 借鉴方案

## 背景

baoyu-diagram 是一个基于 SVG 手绘的图表生成 skill，通过模板化组件 + 分层绘制规则生成复杂图形。本文档分析其核心机制，并提出对 PPT Design Skill 的可借鉴方案。

---

## baoyu-diagram 核心机制

### 1. 支持的图形类型（8 种）

| 类型 | 用途 | 核心 SVG 元素 |
|------|------|---------------|
| Architecture | 系统架构 | 分组矩形 + 连接箭头 + 区域边界 |
| Flowchart | 决策流程 | 菱形判断 + 圆角矩形 + 方向流 |
| Sequence | 时序交互 | 垂直虚线生命线 + 水平消息箭头 + 激活条 |
| Structural | 类图/ER图/组织图 | 多层隔间矩形 + 继承/组合关系线 |
| Mind Map | 思维导图 | 中心节点 + 贝塞尔曲线分支 |
| Timeline | 时间线 | 轴线 + 事件标记 |
| State Machine | 状态机 | 圆角状态节点 + 转换箭头 |
| Data Flow | 数据流 | 处理气泡 + 数据存储 |

### 2. 暗色主题分层系统

所有图形使用统一暗色主题 `#0f172a` 背景 + 语义化颜色：

| 角色 | 颜色 | 用途 |
|------|------|------|
| Primary | cyan `#22d3ee` | 前端/用户面 |
| Secondary | emerald `#34d399` | 后端/服务 |
| Tertiary | violet `#a78bfa` | 数据库/存储 |
| Accent | amber `#fbbf24` | 基础设施 |
| Alert | rose `#fb7185` | 安全/错误 |
| Connector | orange `#fb923c` | 总线/队列 |

### 3. 遮罩层技巧（SVG 特有）

半透明组件会透出下方箭头，解决方案是双层绘制：

```svg
<!-- 第1层：不透明遮罩 -->
<rect x="100" y="100" width="160" height="60" rx="6" fill="#0f172a"/>
<!-- 第2层：半透明视觉层 -->
<rect x="100" y="100" width="160" height="60" rx="6" fill="rgba(8,51,68,0.4)" stroke="#22d3ee"/>
```

### 4. SVG 结构层级（z-order）

严格按顺序绘制保证正确遮挡：

1. 背景 + 网格
2. 区域边界（虚线框）
3. 连接箭头和线
4. 不透明遮罩矩形
5. 组件框（半透明填充 + 描边）
6. 文字标签
7. 图例
8. 标题

### 5. 间距规范

- 组件高度：50-70px
- 组件最小间距：40px（垂直），30px（水平）
- viewBox 边距：30px

---

## 借鉴方案

### 方案 1：SVG 编译器半透明穿透修复

**优先级：** P2

**范围：** `src/ppt_pro_max/renderer/svg_compiler/_compiler.py`

**问题：** 当 SVG 包含半透明组件叠加时，PPTX 中底层内容会穿透显示。这是因为 SVG 和 PPTX 的渲染模型不同：
- SVG：半透明元素会与下方元素混合
- PPTX：每个形状是独立的，半透明不会与下方形状混合

**当前代码行为：**
`_walk()` 方法已经按 SVG 文档顺序直接渲染每个元素（`for child in el:` 遍历），z-order 本身是正确的。问题在于 PPTX 的渲染模型不支持 SVG 的透明度混合。

**解决方案：**

**方案 A（推荐）：遮罩层自动添加**

在 `_render_shape()` 中，当检测到形状有半透明填充（`alpha < 100`）时，自动在该形状下方添加一个不透明的遮罩层：

```python
def _render_shape(self, el, tag, tf, clip_stack):
    # ... 现有代码 ...
    
    # 检测半透明
    fkind, fval, fa = self._paint(el, "fill")
    if fa < 100 and fkind == "solid":
        # 添加不透明遮罩层（使用相同形状，但填充背景色）
        self._add_mask_layer(local, ix0, iy0, iw, ih)
    
    # 继续渲染原始形状
    elem = self._add_freeform(...)
```

**方案 B：SVG 预处理**

在编译前扫描 SVG，将半透明元素拆分为"不透明遮罩 + 半透明视觉层"：

```python
def _preprocess_svg(self, root):
    """将半透明元素拆分为遮罩层+视觉层"""
    for el in root.iter():
        alpha = self._get_opacity(el)
        if alpha < 1.0:
            # 克隆元素，设置为不透明
            mask_el = copy.deepcopy(el)
            mask_el.set("opacity", "1")
            # 插入到原元素之前
            el.getparent().insert(list(el.getparent()).index(el), mask_el)
```

**工作量：** 约 30-50 行代码

---

### 方案 2：间距检查规则

**优先级：** P2

**范围：** `src/ppt_pro_max/build_qa.py`

**问题：** BuildQA 目前没有检查元素间距，可能导致布局拥挤。

**当前检查项（确认无间距检查）：**
- `_check_residual_placeholders` (fatal)
- `_check_blank_page` (fatal)
- `_check_broken_image_ref` (fatal)
- `_check_font_too_small` (warning)
- `_check_text_overflow` (warning)
- `_check_title_duplicate` (warning)
- `_check_color_break` (review)
- `_check_font_mismatch` (review)
- `_check_element_out_of_bounds` (fatal/warning/review)
- `_check_image_stretched` (warning)
- `_check_low_contrast_text` (fatal)
- `_check_page_count` (fatal)
- `_check_toc_mismatch` (warning)

**方案：** 新增 `_check_spacing()` 方法，遍历 slide.shapes 计算相邻形状间距，低于阈值时发出 review 级别警告。

**单位计算：**
- 30px @ 96 DPI = 30/96 inches = 0.3125 inches = 285750 EMU ✓

**阈值（EMU 单位）：**

| 检查项 | 阈值 | 严重级别 |
|--------|------|----------|
| 最小间距 | 30px = 285750 EMU | review |
| viewBox 边距 | 30px = 285750 EMU | review |
| 组件高度异常 | < 50px 或 > 120px | review |

**实现要点：**

```python
def _check_spacing(self, slide, slide_idx: int) -> list[CheckItem]:
    """检查元素间距"""
    issues = []
    shapes = list(slide.shapes)
    min_gap = 285750  # 30px @ 96 DPI = 285750 EMU
    
    for i, s1 in enumerate(shapes):
        for s2 in shapes[i+1:]:
            gap = self._calc_gap(s1, s2)
            if 0 < gap < min_gap:
                issues.append(CheckItem(
                    check_id="spacing_tight",
                    severity="review",
                    message=f"Shapes too close: {gap/914400:.2f}in",
                    shape_id=s1.shape_id,
                ))
    return issues
```

**工作量：** 约 60-80 行代码

---

### 方案 3：语义化颜色角色查询（已取消）

**优先级：** ~~P3~~ → 取消

**原因：** PPT设计与UI设计的颜色系统不同

| UI设计 | PPT设计 |
|--------|---------|
| 按技术角色分配（frontend=cyan, backend=emerald） | 按内容风格分配（调色板系统） |
| 颜色有技术含义 | 颜色有视觉含义 |

PPT的技术图表应该使用PPT的25个调色板系统，而不是UI的颜色角色。颜色选择应该基于：
- 内容风格（专业、科技、温暖等）
- 视觉效果（对比度、可读性）
- 品牌一致性

**结论：** 此方案不适用于PPT设计，已取消。

---

### 方案 4：图表类型扩展（可选）

**优先级：** P4

**方案 A（推荐）：** 使用 `svg_chart()` 实现

- LLM 生成 SVG → `svg_chart()` 编译为 PPTX
- `svg_chart()` 能编译任意合法 SVG，因此可以支持任何图表类型
- 无需新增原生布局

**方案 B：** 新建原生布局

- 每种图表类型需要独立的 layout 算法
- 工作量：600-1200 行代码
- 需要新增 goal type + precision_renderer dispatch

**建议：** 优先使用方案 A，仅在性能或编辑性要求高时考虑方案 B。

**可编辑性说明：**
- 数据图表（饼图、柱状图、折线图）→ 使用 `native_chart()`（python-pptx 原生，可编辑数据）
- 结构图表（架构图、流程图、时序图）→ 使用 `svg_chart()`（各元素可编辑为 Freeform）

**注意事项：**
- `svg_chart()` 本身不做图表语义理解，只是编译 SVG
- 复杂图表（如时序图的生命线、状态机的转换箭头）需要 LLM 生成正确的 SVG 结构
- 某些 SVG 特性不被支持（如 `<image>`、`<filter>`、`<style>` CSS 类）

---

## 高级PPT设计技法分析

### 案例分析

| 案例 | 核心技法 | 现有能力覆盖率 |
|------|----------|----------------|
| 金色流体风格 | PNG素材叠加 + 渐变文字 + 图表 | ~80% |
| 绿色城市风格 | 渐变背景 + 风景照融合 + 自然装饰 | ~70% |
| 黑白时尚风格 | 极端字号对比 + 几何装饰 + 人物照片 | ~65% |

### 核心短板（P0）

| 短板 | 影响 | 补齐方案 | 实际难度 |
|------|------|----------|----------|
| 无"底部渐变蒙版"函数 | 风景照与背景融合 | 渐变矩形叠层（`GradientFill` 已支持 alpha stops） | 低 |
| 无"虚线矩形框"atom | 几何装饰 | 扩展 `rect()` 添加 `dash` 参数 | 低 |
| 缺少"装饰素材库" | 流体/自然装饰 | 需澄清定义（代码生成 or 预制文件） | 中-高 |

### 补充短板（P1）

| 短板 | 影响 | 补齐方案 | 实际难度 |
|------|------|----------|----------|
| 缺少"极大/极小对比排版"atom | 文字排版 | 组合 `text()` + `text()` | 低 |
| 缺少"半透明叠层"atom | 文字可读性 | `frosted_panel()` 已近似实现 | 低 |
| 缺少"页码装饰"atom | 页码设计 | 提取 `_render_footer()` 为独立 atom | 低 |

---

## 实施计划

| 优先级 | 项目 | 工作量 | 难度 | 状态 |
|--------|------|--------|------|------|
| P0 | 底部渐变蒙版函数 | 30-50行 | 低 | ✅ 完成 |
| P0 | 虚线矩形框 atom | 20-30行 | 低 | ✅ 完成 |
| P0 | 装饰素材库 | 可变 | 中-高 | 待定（需澄清定义） |
| P1 | 极大/极小对比排版 atom | 30-50行 | 低 | ✅ 完成 |
| P1 | 半透明叠层 atom | 15-25行 | 低 | 已有 frosted_panel() |
| P1 | 页码装饰组件 | 30-50行 | 低 | ✅ 完成 |
| P1 | 间距检查规则 | 60-80行 | 中 | ✅ 完成 |

---

## 参考资源

- baoyu-diagram skill（全局安装）: `C:\Users\Administrator\.config\opencode\skills\baoyu-diagram\`
- baoyu-diagram skill（项目安装）: `C:\Users\Administrator\.codex\skills\baoyu-diagram\`
- baoyu-diagram references:
  - `references/architecture.md` - 架构图布局算法
  - `references/flowchart.md` - 流程图形状词汇
  - `references/sequence.md` - 时序图核心元素
  - `references/structural.md` - 结构图关系线样式

---

## 修订记录

| 日期 | 修订内容 | 原因 |
|------|---------|------|
| 2026-08-19 | 初稿 | 基于 baoyu-diagram 分析 |
| 2026-08-19 | 方案1：z-order排序 → 半透明穿透修复 | 当前代码已按文档顺序渲染，z-order正确；真正问题是PPTX渲染模型不支持SVG透明度混合 |
| 2026-08-19 | 方案3：SKILL.md硬编码 → ui-ux模块查询 | 配色应由ui-ux模块负责，通过search_color()和get_design_system()查询颜色建议 |
| 2026-08-19 | 修正文件路径 | 列出全局和项目两个安装路径 |
| 2026-08-19 | 方案4：补充svg_chart()说明 | svg_chart()可编译任意SVG，不做图表语义理解 |
| 2026-08-19 | 难度评估修正 | 验证OOXML不支持图片渐变透明；overlay_panel()已近似实现；装饰素材库定义需澄清 |
| 2026-08-19 | 新增高级PPT设计技法分析 | 分析三个案例的设计思路，找出核心短板 |
| 2026-08-19 | 修正渐变蒙版方案 | 使用渐变矩形叠层（GradientFill已支持alpha stops），难度降为"低" |
| 2026-08-19 | 实现P0方案 | dashed_rect(), gradient_mask_image() |
| 2026-08-19 | 实现P1方案 | dramatic_text(), page_number(), 间距检查规则 |