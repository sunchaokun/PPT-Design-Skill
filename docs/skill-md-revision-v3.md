# SKILL.md 修订方案 v3 — 不破坏基础原则

## 核心原则

**系统已经跑通，效果不错。修订是锦上添花，不是搞崩系统。**

这意味着：
1. **content.json 格式是 LLM 写内容的唯一标准** — 不能删、不能改、不能让设计约束与之矛盾
2. **build_helpers.py 是 Build 模式的核心** — SKILL.md 必须包含其 API 参考，否则 LLM 不知道怎么写 build.py
3. **现有 5 步工作流是经过验证的** — 只能在步骤内增加约束，不能改变流程结构
4. **Content Design Rules 是 LLM 写 content.json 的操作指南** — 不能被设计约束覆盖或替代

## v2 方案的 3 个破坏性问题

### 问题 A：build_helpers.py 完全缺失

v2 方案的 SKILL.md 结构中**没有 build_helpers.py 的 API 参考**。

但 Build 模式是系统三种模式之一（FreeStyle / Build / VI Build），且 usage-guide.md 明确说"最终交付用 Build Script 或 VI Build"。LLM 在 Build 模式下需要知道：

- `add_slide(prs)` — 添加空白页
- `page_header(slide, title, subtitle, C)` — 页面标题+分割线
- `kpi_card(slide, left, top, width, height, number, label, trend, C)` — KPI 卡片
- `bar_chart(slide, left, top, data, C)` — 条形图
- `highlight_cards(slide, left, top, cards, C)` — 高亮卡片组
- `copy_decorations(slide, template_slide)` — 复制装饰
- `copy_logo(slide, template_slide, color_hints)` — 复制 LOGO
- `text()` / `multiline()` / `rect()` / `rrect()` / `oval()` — 基础形状
- C 字典的角色名约定（primary/accent/muted/text_dark/text_body/text_muted 等）

**如果不写这些，LLM 在 Build 模式下无法工作。**

### 问题 B：content.json 格式不可压缩

v2 方案保留了 content.json 格式（35 行），这是对的。但 v1 压缩评估曾建议"Python API 压缩到 5 行"——如果过度压缩 `generate_ppt()` 签名，LLM 不知道 `content_file` 参数怎么传，也不知道 `fetch_images=True` 的效果。

**content.json 格式 + generate_ppt() 签名 + build_helpers API 是 SKILL.md 的三大不可压缩核心。**

### 问题 C：Design Constraints 可能与 content.json goal type 冲突

v2 方案的 Dial Action Map 建议 LLM 写 `layout_variant:"sidebar-left"` 和 `animation:"fade-in"` 到 content.json，但当前 content.json 格式定义中**没有这两个字段**。

如果 LLM 写了渲染器不认识的字段，要么被忽略（无害但 Dial 仍是空声明），要么报错（破坏系统）。

**必须先确认渲染器是否接受这些字段，再写入 SKILL.md。**

## 修正方案

### 修正 1：SKILL.md 必须包含 build_helpers API 参考

在 Component Library 之后、Key Constraints 之前，新增：

```markdown
## Build Helpers API (for Build/VI Build mode)

LLM writes `build.py` scripts using these functions. Import: `from ppt_pro_max.build_helpers import *`

### Color Dictionary (C)
```python
C = {
    'primary': '#2E6504', 'accent': '#7DA92F', 'muted': '#81C784',
    'light': '#C8E6C9', 'white': '#FFFFFF', 'background': '#FFFFFF',
    'card_bg': '#F9F9F9', 'text_dark': '#1A1A1A', 'text_body': '#333333',
    'text_muted': '#666666', 'divider': '#CCCCCC',
    'font_heading': '微软雅黑', 'font_body': '微软雅黑',
}
```

### Functions

| Function | Purpose | Key Params |
|----------|---------|------------|
| `add_slide(prs)` | Add blank slide | Auto-finds blank layout |
| `page_header(slide, title, subtitle, C)` | Title + subtitle + divider line | Standard content page header |
| `kpi_card(slide, left, top, width, height, number, label, trend, C)` | KPI metric card | number (big), label, trend (+8.3%) |
| `bar_chart(slide, left, top, data, C)` | Horizontal bar chart | data: [(label, pct, val), ...] |
| `comparison_bars(slide, left, top, metrics, C)` | Before/after comparison | metrics: [(label, v_old, v_new, pct_old, pct_new), ...] |
| `donut_chart(slide, cx, cy, radius, inner_radius, sectors, C)` | Donut chart (simplified) | sectors: [(name, pct_str, color), ...] |
| `highlight_cards(slide, left, top, cards, C)` | Highlight card row | cards: [(title, desc, accent_color), ...] |
| `text(slide, left, top, width, height, txt, font_size, color, bold, align, font_name, C)` | Single-line text | color: role name or hex |
| `multiline(slide, left, top, width, height, lines, font_size, color, C)` | Multi-line text | lines: list of strings |
| `rect(slide, left, top, width, height, fill, line, C)` | Rectangle | fill/line: role name or hex |
| `rrect(slide, left, top, width, height, fill, line, C)` | Rounded rectangle | Same as rect |
| `oval(slide, left, top, width, height, fill, line, C)` | Ellipse | Same as rect |
| `top_bar(slide, color, C)` | Top accent bar | Brand color strip |
| `copy_decorations(slide, template_slide)` | Copy decorations from template | Skips long text (>50 chars) and images |
| `copy_logo(slide, template_slide, color_hints)` | Copy LOGO from template | Only finds GROUP shapes (shape_type==6) |

### Color Resolution
- Hex value: `'#2E6504'` → used directly
- Role name: `'primary'` → looks up `C['primary']`
- Missing role: returns `'#000000'` (never crashes)
```

**约 30 行**。这是 Build 模式能工作的最低信息量。

### 修正 2：Dial Action Map 不引入新 content.json 字段

v2 方案建议 LLM 写 `layout_variant` 和 `animation` 到 content.json，但这两个字段当前渲染器不识别。

**修正**：Dial Action Map 只约束 LLM 的**已有字段选择行为**，不引入新字段。

```markdown
### Dial Action Map (V/M/D → LLM decisions using EXISTING content.json fields)

| VARIANCE | Action (use existing goal/layout_variant/style params) |
|----------|------------------------------------------------------|
| 1-3 | Use `goal:"content"` with centered layouts; equal-width cards; `--layout-variant centered` or `standard` in generate_ppt() call |
| 4-7 | Mix `goal:"content"` with `goal:"features"`; feature first card; `--layout-variant sidebar-left` or `asymmetric` in generate_ppt() call; vary which pages have images |
| 8-10 | Use diverse goal types per page; avoid any repeated layout family; `--layout-variant asymmetric`; insert section dividers between every topic shift |

| MOTION | Action |
|--------|--------|
| 1-3 | No special action — default transitions only |
| 4-7 | Ensure cover slide has `goal:"hook"` (gets fade-in); section dividers get entrance animation automatically |
| 8-10 | Same as 4-7 plus: request `--motion 8` in generate_ppt() call; more section dividers for animation variety |

| DENSITY | Action (bullet count per page in content.json) |
|---------|------------------------------------------------|
| 1-3 | 2-3 bullets per content page; insert breathing pages (goal:"section" or goal:"content" with ≤2 bullets) after every 2 content pages |
| 4-7 | 3-5 bullets per content page; mix: some pages 3 bullets, some 6+ (triggers two-column) |
| 8-10 | 6+ bullets on data/overview pages; use `component_type:"group"` + `component_category:"infographic"` for dense pages; no breathing pages |
```

**关键变化**：
- `layout_variant` 和 `animation` 不再作为 content.json 字段，而是作为 `generate_ppt()` 的 CLI 参数
- LLM 通过 `--layout-variant` 和 `--motion` 参数影响渲染，而不是在 content.json 中写渲染器不认识的字段
- 这完全兼容现有系统 — 这些参数已经存在于 `generate_ppt()` 签名中

### 修正 3：三大不可压缩核心明确标注

在 SKILL.md 结构中，明确标注哪些内容是"不可压缩的"：

```markdown
## ⚠️ Non-Negotiable Sections (DO NOT compress or remove)

These sections are the LLM's only reference for writing correct output:
1. **content.json Format** — LLM must know the exact schema to write valid content
2. **brand.json Format** — LLM must know brand spec structure for enterprise mode
3. **build_helpers API** — LLM must know function signatures to write build.py
4. **Content Design Rules** — LLM must know which content patterns trigger which rendering
5. **Key Constraints** — LLM must know API signatures to write correct python-pptx code
6. **generate_ppt() signature** — LLM must know valid parameters to call the pipeline
```

## 修正后的 SKILL.md 完整结构

| # | Section | Lines | Status | Non-Negotiable |
|---|---------|-------|--------|----------------|
| 1 | Frontmatter | 11 | keep | yes (skill format) |
| 2 | Title + python-pptx ref | 2 | compress (1 link + 1 line) | yes (API reference) |
| 3 | Execution Workflow | 45 | keep + Design Read/Dial in Step 1 | yes (workflow) |
| 4 | Content Design Rules | 20 | upgrade (12→20 rows, merge 4 rules) | **yes** |
| 5 | When to Activate | 8 | keep | no |
| 6 | **Design Constraints** | **~265** | **NEW** | no (but high value) |
| 7 | Python API | 15 | compress | **yes** (generate_ppt signature) |
| 8 | Design Atoms | 8 | compress (drop Status col) | no |
| 9 | 10 Diagram Types | 14 | keep | no |
| 10 | Image Engines | 30 | compress | no |
| 11 | Animation System | 3 | compress | no |
| 12 | Enterprise Project Structure | 15 | keep | no |
| 13 | **content.json Format** | **35** | **keep** | **yes** |
| 14 | **brand.json Format** | **23** | **keep** | **yes** |
| 15 | Page Revision Syntax | 8 | compress | no |
| 16 | Component Library | 80 | compress | no |
| 17 | **Build Helpers API** | **~30** | **NEW** | **yes** |
| 18 | **Key Constraints** | **17** | **keep** | **yes** |
| 19 | CLI Quick Reference | 10 | compress from 67 | no |
| | **Total** | **~654** | | |

**对比 v2 方案的 ~690 行，v3 方案 ~654 行**（build_helpers +30 行，但 Dial Action Map 少 3 行因为不引入新字段，CLI 压缩多省 5 行）。

## 修订原则总结

| 原则 | 含义 | 违反后果 |
|------|------|----------|
| **不删标准** | content.json/brand.json 格式、build_helpers API、generate_ppt() 签名不可压缩 | LLM 不知道怎么写正确输出 |
| **不改流程** | 5 步工作流结构不变，只在 Step 1/3 内增加约束 | 改流程 = 改系统行为 = 风险 |
| **不引入新字段** | Dial Action Map 只用现有 content.json 字段和 generate_ppt() 参数 | 新字段渲染器不认识 = 空声明或报错 |
| **不替代 CDR** | Design Constraints 补充 Content Design Rules，不替代 | CDR 是"做什么触发什么渲染"，Constraints 是"不做什么" |
| **不省 build_helpers** | Build 模式是交付级质量的核心路径 | 省了 = Build 模式无法工作 |

## 与 v2 方案的差异

| 项 | v2 方案 | v3 修正 | 原因 |
|----|---------|---------|------|
| build_helpers API | 缺失 | +30 行 | Build 模式核心，不可省 |
| Dial Action Map | 引入 layout_variant/animation 新字段 | 只用现有 generate_ppt() 参数 | 渲染器不认识新字段 |
| content.json 新字段 | layout_variant, animation | 不新增 | 同上 |
| Non-Negotiable 标注 | 无 | 明确 6 大不可压缩核心 | 防止未来压缩时误删 |
| 总行数 | ~690 | ~654 | build_helpers +30 但 Dial 精简 -3, CLI 多省 -5 |

## 执行检查清单

在写入 SKILL.md 之前，确认：

- [ ] content.json 格式完整保留（35 行，一字不改）
- [ ] brand.json 格式完整保留（23 行，一字不改）
- [ ] build_helpers API 参考已包含（~30 行）
- [ ] generate_ppt() 签名已包含（Python API 节）
- [ ] Content Design Rules 升级但不丢失原有 8 条规则
- [ ] Design Constraints 不与 Content Design Rules 矛盾
- [ ] Dial Action Map 不引入 content.json 新字段
- [ ] AI Tells 无重复项
- [ ] 全英文，中文术语括号标注
- [ ] Consistency Locks 不与 Color 节重复
- [ ] Key Constraints 完整保留（API 签名、OOXML 值约定）
- [ ] 运行 `python -m pytest tests/ -q` 全部通过（SKILL.md 改动不影响代码，但确认无副作用）
