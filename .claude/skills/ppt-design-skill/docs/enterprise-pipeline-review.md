# Enterprise Pipeline 设计方案 — 深度审查报告

> 审查版本: v1.3 | 审查日期: 2026-07-11 | 审查方法: 逐行对照代码库 + python-pptx 1.0.2 API 实测

---

## 一、必须修复的错误（Blocker）

### B1. §3.2 提取内容表声称 `slide_master.theme.color_scheme` 可访问 — **API 不存在**

**位置**: 第 158-159 行

```
| 主题色 | `slide_master.theme.color_scheme` | 品牌颜色 |
| 字体方案 | `slide_master.theme.font_scheme` | 标题/正文字体 |
```

**实测结果**: python-pptx 1.0.2 的 `SlideMaster` 对象**没有 `theme` 属性**。`dir(slide_master)` 中不存在 `theme`。主题色/字体方案存储在 OOXML 包的 `theme/theme1.xml` 中，python-pptx 不提供 Python 级别的主题色提取 API。

**修复方案**:
- 明确标注需通过 XML 解析提取：遍历 `slide_master.element` 的 `p:clrMap` 子元素获取颜色映射
- 或者通过 `slide_master.part.package.part_related_by(Relationship.THEME)` 获取 theme part 的 XML
- 提取内容表应改为：

```
| 主题色 | slide_master XML p:clrMap / theme part XML a:clrScheme | 需 XML 解析，无 Python API |
| 字体方案 | theme part XML a:fontScheme | 需 XML 解析，无 Python API |
```

### B2. §3.7 表格称 `slide_layouts[6]` 为空白布局 — **仅对默认模板成立**

**位置**: 第 372 行

```
| 布局选择 | `slide_layouts[6]` 空白布局 | `slide_layouts[mapped_index]` 模板布局 |
```

**实测结果**: 默认模板 `slide_layouts[6]` 是 "Blank"，但企业模板的布局数量和索引完全不同。FreeStyleRenderer 当前写死 `[6]` 在企业模板上可能越界或选错布局。

**修复方案**:
- FreeStyleRenderer 应改为按名称查找空白布局（遍历 `slide_layouts` 找 `name == "Blank"` 的），而非硬编码索引
- 或保持现有逻辑（FreeStyle 不用企业模板），但文档需明确：`slide_layouts[6]` 仅对默认模板有效，EnterpriseRenderer 不可硬编码索引

### B3. §7.1 `--version` 与 Python/argparse 冲突

**位置**: 第 595 行

```
| `--version` | int | 指定输出版本号（覆盖已有版本，需配合 --project） |
```

**问题**: `--version` 是 CLI 工具的惯例参数（用于显示程序版本号，如 `app --version` → `0.2.0`）。argparse 的 `parser.add_argument('--version')` 会与 `action='version'` 惯例冲突。虽然此文档用 `type=int` 不是 `action='version'`，但用户心理模型会产生混淆。且当前代码中 `__version__ = "0.2.0"` 已存在。

**修复方案**: 重命名为 `--output-version` 或 `--ver`，避免与程序版本号混淆：

```
| `--output-version` | int | 指定输出版本号（覆盖已有版本，需配合 --project） |
```

### B4. §7.1 `--density` 参数与现有 `--density` 冲突定义

**位置**: 第 594 行

**问题**: 当前 `cli.py` 第 37 行已定义 `--density` 参数：`parser.add_argument("--density", type=int, choices=range(1, 11))`。文档中新增的 `--density` 描述"覆盖 business_mode 默认值"是扩展现有参数语义，不是新增参数。但 §4.2 和 §7.1 表述为"新增"，可能误导实现者重复添加。

**修复方案**: 在 §7.1 表格中标注 `--density` 为"扩展现有参数"，非"新增"：

```
| `--density` | int 1-10 | **扩展现有参数**：Enterprise Pipeline 中覆盖 business_mode 默认值 |
```

### B5. §5.1 / §3.2 brand.json 的 colors 键名与现有 ThemeMapper 颜色角色不兼容

**位置**: 第 174-176 行、第 473-481 行

brand.json 定义：
```json
"colors": {
    "primary": "#1A3C6E",
    "secondary": "#E8491D",
    "accent": "#2ECC71",
    "background": "#FFFFFF",
    "text": "#333333",
    "text_secondary": "#666666"
}
```

现有 ThemeMapper/ThemeComposer 的颜色角色：
```python
"primary", "on-primary", "secondary", "accent",
"background", "foreground", "muted", "muted-foreground", "border", "destructive"
```

**不兼容点**:
1. brand.json 用 `text` / `text_secondary`，系统用 `foreground` / `muted-foreground`
2. brand.json 缺少 `on-primary`（primary 上的文字色）、`muted`、`border`、`destructive`
3. 如果直接把 brand.json 的 colors 喂给 ThemeMapper，`foreground` 和 `muted-foreground` 会找不到

**修复方案**: 统一键名，brand.json 使用与 ThemeMapper 一致的角色名：

```json
"colors": {
    "primary": "#1A3C6E",
    "on-primary": "#FFFFFF",
    "secondary": "#64748B",
    "accent": "#2ECC71",
    "background": "#FFFFFF",
    "foreground": "#333333",
    "muted": "#F1F5F9",
    "muted-foreground": "#666666",
    "border": "#E2E8F0",
    "destructive": "#EF4444"
}
```

最小可用示例也需同步更新。

### B6. §12.9 Connector 代码中 `line_xml` 查找路径错误

**位置**: 第 1559 行

```python
line_xml = connector._element.find(qn("a:ln"))
```

**实测结果**: Connector 的 XML 结构是 `<p:cxnSp><p:spPr>...<a:ln>...</a:ln></p:spPr></p:cxnSp>`。`a:ln` 是 `p:spPr` 的子元素，不是 `p:cxnSp` 的直接子元素。`connector._element.find(qn("a:ln"))` 在 LXML 中**不会递归搜索**（`find` 只搜索直接子元素），所以会返回 `None`。

**修复方案**:

```python
spPr = connector._element.find(qn("p:spPr"))
line_xml = spPr.find(qn("a:ln")) if spPr is not None else None
if line_xml is not None:
    tail_arrow = etree.SubElement(line_xml, qn("a:tailEnd"))
    # ...
```

或者使用 `connector._element.find(".//" + qn("a:ln"))` 带XPath递归搜索。

### B7. §12.5.2 TextMeasurer 的 CHAR_WIDTH_TABLE 经验值有系统性偏差

**位置**: 第 1094-1106 行

```python
CHAR_WIDTH_TABLE = {
    10: 0.08,   # 10pt → 每字符 0.08 英寸
    ...
}
```

**验证**: 10pt 字号下每字符 0.08 英寸 = 10pt × 0.008。但实际排版中：
- Calibri 10pt 的平均字符宽度约为 0.055-0.065 英寸（小写），0.075-0.085 英寸（大写）
- 中文字符 10pt 约为 0.14 英寸（而非 0.08 × 2 = 0.16）
- 经验值偏高约 20-30%，会导致节点偏大，不是偏小

影响：节点尺寸偏大不是严重问题（不会溢出），但会浪费空间，降低密集布局的信息密度。

**修复方案**: 
- 降低系数为 `0.006` 而非 `0.008`（更保守）
- 或者添加 10% 安全边距注释说明这是有意偏高
- 远期用 Pillow render 精确测量

### B8. §3.2 PPT 模板识别规则"目录下唯一 .pptx 文件"有逻辑漏洞

**位置**: 第 131 行

```
| PPT 模板 | 目录下唯一 `.pptx` 文件 | 否 |
```

**问题**: `output/v1/presentation.pptx`、`output/v2/presentation.pptx` 等历史版本也在项目目录下。虽然它们在 `output/` 子目录中，但"目录下"表述含糊。如果递归扫描，会把历史版本也识别为模板。

**修复方案**: 明确只扫描项目文件夹**根目录**（不递归），或排除 `output/` 子目录：

```
| PPT 模板 | 项目根目录下（不递归）唯一的 `.pptx` 文件 | 否 |
```

---

## 二、设计缺陷（需要补充/修改）

### D1. §3.2 TemplateAnalyzer 无法从模板提取颜色/字体 — 缺少实现路径

文档声称"从模板提取品牌颜色/字体"，但 python-pptx 没有提供 `slide_master.theme` 属性。TemplateAnalyzer 必须通过 XML 解析实现，文档应提供具体实现方案：

```python
# 方案1: 通过 slide_master 的 part 获取 theme part
theme_part = slide_master.part.package.part_related_by(
    pptx.opc.constants.Relationship.THEME
)
theme_element = theme_part.element
# 从 theme_element 解析 a:clrScheme 和 a:fontScheme

# 方案2: 从 p:clrMap 映射获取（仅获取颜色角色映射，不含具体RGB值）
clr_map = slide_master.element.find(qn('p:clrMap'))
# clr_map 有 val 属性映射到 theme color indices
```

建议在 §3.2 增加"模板解析实现约束"小节，明确 XML 解析方案。

### D2. §3.7 EnterpriseRenderer 的 `Presentation(template_path)` 缺少模板不存在时的错误处理

**位置**: 第 371 行

```python
Presentation(template_path)
```

**问题**: 如果 `template_path` 是无效路径或损坏的 .pptx 文件，`Presentation()` 会直接抛异常。文档未定义降级行为。

**修复方案**: 增加降级逻辑：

```python
def _create_presentation(self, template_path):
    if template_path and os.path.exists(template_path):
        try:
            return Presentation(template_path)
        except Exception:
            # 模板损坏，降级为空白创建
            return Presentation()
    return Presentation()
```

### D3. §3.5 ReviewGate 的 `--no-review` 参数语义不清

**位置**: 第 592 行

```
| `--no-review` | flag | 显式跳过确认环节（与 --review 互斥） |
```

**问题**: 默认行为已经是"跳过确认"（不传 `--review` 时直接生成）。`--no-review` 的存在意义是什么？只有在某种默认启用 review 的场景下才有用，但文档说默认不启用。

**修复方案**: 删除 `--no-review`，或在 `brand.json` 中增加 `"review_by_default": true` 选项，此时 `--no-review` 才有意义。

### D4. §4.1 新增布局 ID 15 (Code Block) 与 §12.11 Diagram Focus (ID 20) 的 goal 映射冲突

**位置**: 第 425 行 vs 第 1590 行

- Layout ID 15 "Code Block" 映射 goal: `code-demo`
- Layout ID 20 "Diagram Focus" 映射 goal 包含: `process`, `pipeline`, `funnel`, `timeline`, `hierarchy`, `cycle`, `swot`, `comparison`, `pyramid`, `architecture`

**问题**: 如果未来某个 goal 同时匹配多个布局，`LayoutRegistry.get_layout_by_goal()` 当前只返回**第一个匹配**（`content-generator.py:319-323`），且直接 fallback 到 `content-with-title`。新增布局后，goal 到布局的映射优先级不明确。

**修复方案**: 定义明确的 goal→layout 优先级规则，或改为 goal 可以映射到多个布局（由 DesignDecider 根据场景选择）。

### D5. §10 目录结构缺少 `renderer/diagram/layout_engine.py` 与 §12.5 的对应关系

**位置**: 第 829 行

目录结构中有 `layout_engine.py`（坐标计算、自适应排版），但 §12.8 的 `LayoutEngine` 类放在哪里没有明确对应。建议在 §12.3 架构设计中注明 `LayoutEngine` 位于 `renderer/diagram/layout_engine.py`。

### D6. §9.4 latest 链接的 Windows junction 实现缺少错误处理

**位置**: 第 772 行

```
- **Windows**：目录 junction `os.system('mklink /J ...')`，无需管理员权限
```

**问题**: `os.system()` 返回退出码但不抛异常。如果创建 junction 失败（如路径含空格未加引号），代码会静默失败。

**修复方案**: 使用 `subprocess.run()` 检查返回码，或用 Python 的 `os.symlink()` + `os.makedirs(junction, ...)` 替代（Python 3.8+ 支持在 Windows 上创建 junction）。

### D7. §3.4 EnterpriseDesignDecider "密度: 固定低密度"描述不准确

**位置**: 第 256 行

```
| 密度 | 固定低密度 | 根据 business_mode 和 density 参数调整 |
```

**问题**: FreeStyle DesignDecider 当前已支持 `density` 参数（`cli.py:37`, `design_decider.py:70`），不是"固定低密度"。虽然当前实现中 `density` 仅传递给 DesignDecider 但并未在 PageDesign 中实际影响字号/间距，但参数已存在。

**修复方案**: 改为"密度参数存在但效果有限"而非"固定低密度"。

### D8. §6.1 content.json 的 `cards` 字段与 PageContent dataclass 不兼容

**位置**: 第 567-571 行

content.json:
```json
"cards": [
    {"title": "智能匹配", "body": "AI算法精准匹配", "image": "dashboard.png"},
    ...
]
```

当前 `PageContent` dataclass (content_generator.py:14-24):
```python
@dataclass
class PageContent:
    position: int
    goal: str
    title: str = ""
    subtitle: str | None = None
    bullets: list[str] = field(default_factory=list)
    metrics: list[dict[str, str]] | None = None
    quote: dict[str, str] | None = None
    chart_data: dict[str, Any] | None = None
    image_keywords: str = ""
```

**问题**: `PageContent` 没有 `cards` 字段，也没有 `image` 字段。content.json 的 `cards` 和 `image` 无法直接映射到现有数据结构。

**修复方案**: 在 §12.11 "ContentGenerator 扩展"部分已有 `diagram_type` 和 `diagram_data` 的扩展，但还缺少 `cards` 和 `image` 的扩展。应补充：

```python
@dataclass
class PageContent:
    # ... 现有字段 ...
    image: str | None = None           # 显式指定图片文件名
    cards: list[dict] | None = None   # 卡片数据
    diagram_type: str | None = None
    diagram_data: dict | None = None
```

### D9. §2.1 分流逻辑的 `_generate_ppt_freestyle` 函数不存在

**位置**: 第 73 行

```python
return _generate_ppt_freestyle(query, **kwargs)
```

当前 `__init__.py` 中 generate_ppt 的 FreeStyle 逻辑是内联的，没有抽取为独立函数。文档伪代码暗示需要重构为 `_generate_ppt_freestyle()`，但未在实现计划中体现。

**修复方案**: 在 Phase A 实现计划中增加"重构 generate_ppt() 为分流结构，现有逻辑提取为 _generate_ppt_freestyle()"。

---

## 三、遗漏和待补充项

### M1. 模板占位符类型识别方案缺失

§3.7 提到"识别模板占位符类型（title/body/image/chart）"，但 python-pptx 的占位符类型体系与文档假设不同：

- `PP_PLACEHOLDER.TITLE` (1) — 标题
- `PP_PLACEHOLDER.CENTER_TITLE` (3) — 居中标题
- `PP_PLACEHOLDER.SUBTITLE` (4) — 副标题  
- `PP_PLACEHOLDER.BODY` (2) / `PP_PLACEHOLDER.OBJECT` (7) — 正文/通用
- `PP_PLACEHOLDER.CHART` (8) — 图表
- `PP_PLACEHOLDER.PICTURE` (18) — 图片
- `PP_PLACEHOLDER.TABLE` (12) — 表格

文档应列出占位符类型到 PageContent 字段的完整映射表。

### M2. 教学动画的实现方案过于模糊

§4.3 提到 4 种教学动画类型，但 python-pptx 没有动画 API。当前 `_apply_transition` 仅实现了幻灯片切换，**不是元素级动画**。Click-by-click reveal 需要在 `<p:timing>` XML 中为每个段落添加独立的 click effect，这需要大量 XML 构造。

建议增加动画 XML 模板示例，或在 Phase G 中标注为"研究性任务"。

### M3. 缺少 content.json 与现有 `--content` 参数的关系说明

当前 CLI 有 `--content` 参数（cli.py:34），指向一个 JSON 文件。新增的 `content.json` 是放在项目文件夹中的，两者功能重叠。

需要明确：
1. `--content` 与项目文件夹中的 `content.json` 同时存在时谁优先？
2. `content.json` 是否复用 `_load_user_content()` 的解析逻辑？当前解析将 JSON 扁平化为 context dict，无法识别 `slides[]` 结构。

### M4. §3.1 ProjectScanner 的 "文件名含 logo 的图片文件" 规则不够精确

需要定义"含 logo"的精确匹配规则：
- `logo.png` — 匹配
- `company-logo.png` — 匹配？
- `logo_dark.png` — 匹配？
- `technology.png` — 不匹配

建议：文件名（去掉扩展名后）完全等于 `logo`，或以 `logo` 开头 + 分隔符（`_`/`-`），或以 `logo` 结尾。

### M5. 缺少 Enterprise Pipeline 的 `generate_ppt()` API 签名扩展

§2.1 的分流伪代码中 `generate_ppt(query, **kwargs)` 需要新增 `project`、`business_mode`、`review`、`review_file`、`output_version`、`from_version`、`history` 等参数。当前函数签名（`__init__.py:17-43`）没有这些参数。

建议在 §7 或 §2 中增加完整的扩展后 `generate_ppt()` 签名。

### M6. DiagramEngine 的 `add_group_shape` 验证结果

**实测确认**: `slide.shapes.add_group_shape(shapes_iterable)` 在 python-pptx 1.0.2 中可用，接受 `Iterable[BaseShape]` 参数，返回 `GroupShape`。但有一个关键限制：**只有已经添加到 slide 上的 shape 才能被加入 group**。不能先创建 group 再往里面加 shape。

这意味着实现顺序必须是：
1. 先把所有节点/连线 add 到 slide 上
2. 收集这些 shape 对象
3. 调用 `add_group_shape(shapes_list)` 打组

文档应在待决事项中注明此约束。

### M7. `add_connector` 的坐标单位是 EMU 而非英寸

§12.9 的代码示例中直接传 `Inches()` 值给 `add_connector`，实测确认这是正确的——python-pptx 的 `add_connector` 接受 `Length` 类型（`Inches()` 返回 EMU 值）。文档正确，但建议添加注释说明坐标单位。

### M8. `conn.line.width` 只接受整数 EMU，不接受 float

**实测发现**: `conn.line.width = 1.5` 会抛出 `TypeError: value must be an integral type`。必须使用 `conn.line.width = Pt(1.5)` 传入 EMU 值。

文档 §12.9 的代码 `connector.line.width = Pt(width_pt)` 是正确的，但 §12.6 DiagramStyle 中 `connector_width_pt: float = 1.5` 的使用者需注意必须经过 `Pt()` 转换，不能直接赋 float。

### M9. `build_freeform` 的坐标单位也是 EMU

**实测确认**: `slide.shapes.build_freeform(start_x, start_y)` 接受 EMU 值。传入 `Inches()` 值是正确的。`add_line_segments` 的坐标也是 EMU。循环图的弧形箭头用 freeform 实现是可行的。

### M10. ROUNDED_RECTANGLE 的圆角半径可通过 `adjustments` 控制

**实测确认**: `shape.adjustments[0]` 返回当前圆角比例（默认约 0.167），可通过 `shape.adjustments[0] = 0.5` 调整。

文档 §12.6 DiagramStyle 中 `node_corner_radius: str = "rounded"` 可实现为：
- `"sharp"` → `adjustments[0] = 0.02`
- `"rounded"` → `adjustments[0] = 0.15`（默认）
- `"pill"` → `adjustments[0] = 0.5`

建议在 §12.10 或 §12.6 中补充此实现细节。

---

## 四、文档内部不一致

### I1. §4.1 新增布局 ID 20-22 与 §10 目录注释不一致

**§4.1**: 布局 ID 20 "Diagram Focus"、ID 21 "Diagram + Text"、ID 22 "Table Focus"

**§10 第 820 行**: `layout_registry.py # 扩展：新增密集布局（ID 12-20）+ 图形布局（ID 21-23）`

注释说"ID 21-23"，但 §4.1 只定义了 ID 20-22。ID 23 不存在。应改为"ID 12-22"。

### I2. §4.1 布局 ID 20 的 goal_mapping 与 §12.11 不完全一致

**§4.1 第 430 行**: `"process, pipeline, funnel, timeline, hierarchy, cycle, swot, pyramid, architecture"`

**§12.11 第 1590-1592 行**: `"process, pipeline, funnel, timeline, hierarchy, cycle, swot, comparison, pyramid, architecture"`

§12.11 多了 `comparison`。应以 §12.11 为准（更完整），§4.1 需补充 `comparison`。

### I3. §7.1 与 §9.6 重复定义 CLI 参数

§7.1 定义了 `--version`、`--from-version`、`--history`，§9.6 又单独列出了这三个参数。两处描述一致，但重复容易导致修改时遗漏。建议 §9.6 改为"参见 §7.1"。

### I4. §3.2 BrandSpec 的 colors 键名与 §5.1 brand.json 的 colors 键名不一致

**§3.2 第 175-176 行**:
```python
# {"primary": "#1A3C6E", "secondary": "#E8491D", "accent": "#2ECC71",
#  "background": "#FFFFFF", "text": "#333333", "text_secondary": "#666666"}
```

**§5.1 第 474-480 行**:
```json
"colors": {
    "primary": "#1A3C6E",
    "secondary": "#E8491D",
    "accent": "#2ECC71",
    "background": "#FFFFFF",
    "text": "#333333",
    "text_secondary": "#666666"
}
```

两处一致，但都与 ThemeMapper 的角色名不兼容（见 B5）。修改 B5 时需同步修改这两处。

### I5. §12.11 注释"合并原 ID 15 Diagram+Notes 和 ID 21 Diagram Focus"有误

**位置**: 第 1586 行

```
# 布局 ID 20: Diagram Focus（合并原 ID 15 Diagram+Notes 和 ID 21 Diagram Focus）
```

ID 15 在 §4.1 中定义为 "Code Block"，不是 "Diagram+Notes"。此注释是旧版残留，应删除或修正。

---

## 五、建议优化（非阻塞）

### S1. TextMeasurer 应使用 Pillow render 精确测量

当前经验值估算（CHAR_WIDTH_TABLE）存在 20-30% 偏差。Pillow 的 `ImageDraw.textbbox()` 可精确测量文字尺寸：

```python
from PIL import Image, ImageDraw, ImageFont
img = Image.new('RGB', (1, 1))
draw = ImageDraw.Draw(img)
font = ImageFont.truetype("arial.ttf", 14)
bbox = draw.textbbox((0, 0), "Hello World", font=font)
width_inches = (bbox[2] - bbox[0]) / 96
height_inches = (bbox[3] - bbox[1]) / 96
```

风险：字体文件在用户机器上可能不存在。可用 font fallback。

### S2. 增加图形节点的垂直居中对齐实现方案

文档未说明如何在 Shape 内实现文字垂直居中。实测确认两种方式：
- `shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER` 仅水平居中
- 垂直居中需通过 XML：`text_frame._element.find(qn('a:bodyPr')).set('anchor', 'ctr')`

建议在 §12.5 或 DiagramStyle 中补充此实现细节。

### S3. 模板 LOGO 继承机制需详细说明

§3.7 提到"模板 master 上已有 LOGO → 自动继承"，但实测默认模板的 master 只有 placeholder shapes（title/text/date/footer/slide_number），没有图片 shape。企业模板的 LOGO 可能在：
1. Slide Master 的 shapes 中（非 placeholder）
2. 某个 Slide Layout 中
3. 作为背景水印

需要定义 LOGO 检测规则：遍历 `slide_master.shapes` 找 `shape.shape_type == MSO_SHAPE_TYPE.PICTURE` 或 `AUTO_SHAPE` 且名称含 "logo" 的 shape。

### S4. 建议在 §9.2 版本号确定逻辑中排除 `latest` junction

```python
for entry in os.listdir(output_dir):
    if entry.startswith("v") and entry[1:].isdigit():
```

`latest` 不以数字开头，所以不会被误匹配。但如果是 junction 指向的目录，`os.listdir` 不会列出 junction 目标目录本身（只列出 junction 名）。此逻辑正确，无需修改，但建议添加注释说明 `latest` 不受影响。

### S5. Connector 的 `begin_connect` / `end_connect` 的 connection site 索引缺少文档

实测确认 `conn.begin_connect(shape, 3)` 中的 `3` 是 connection site 索引（0-7 对应8个方向）。但 python-pptx 没有文档说明哪个索引对应哪个方向。建议在 §12.9 补充 connection site 索引说明，或建议使用坐标计算方式而非 `begin_connect`。

### S6. 建议增加 `--project` 参数对 `query` 参数可选化的支持

§7.2 示例 `python -m ppt_pro_max --project my_project/ --history` 中 `query` 参数是必需的（cli.py:23），但 `--history` 模式不需要 query。需要调整 argparse 使 query 在 `--history` 模式下可选。

---

## 六、python-pptx API 实测验证总结

| API | 文档声称 | 实测结果 | 状态 |
|-----|---------|---------|------|
| `Presentation()` 空白创建 | 可用 | ✅ 可用 | 正确 |
| `Presentation(path)` 模板创建 | 可用 | ✅ 可用 | 正确 |
| `slide_layouts[6]` = "Blank" | 默认模板 | ✅ 默认模板可用，企业模板不确定 | 需注意 |
| `slide_master.theme.color_scheme` | 可提取颜色 | ❌ 不存在此属性 | **B1** |
| `add_connector(STRAIGHT/ELBOW/CURVE)` | 可用 | ✅ 三种均可用 | 正确 |
| `add_table(rows, cols, ...)` | 可用 | ✅ 可用 | 正确 |
| `add_group_shape(shapes)` | 可用 | ✅ 可用，需先 add shapes 到 slide | M6 |
| `build_freeform()` | 可用 | ✅ 可用，支持闭合路径 | 正确 |
| `MSO_SHAPE.TRAPEZOID` | 可用 | ✅ 可用 | 正确 |
| `MSO_SHAPE.DIAMOND` | 可用 | ✅ 可用 | 正确 |
| `MSO_SHAPE.OVAL` | 可用 | ✅ 可用 | 正确 |
| `MSO_SHAPE.CHEVRON` | 可用 | ✅ 可用 | 正确 |
| `MSO_SHAPE.RIGHT_ARROW` | 可用 | ✅ 可用 | 正确 |
| `conn.line.color.rgb` | 可设置 | ✅ 可用 | 正确 |
| `conn.line.width = float` | 可设置 | ❌ 必须用 Pt() | **M8** |
| `conn.line.dash_style = MSO_LINE.DASH` | 可设置 | ✅ 可用 | 正确 |
| `connector._element.find(qn("a:ln"))` | 可找到 line 元素 | ❌ 需通过 spPr 中转 | **B6** |
| `shape.fill.fore_color.brightness` | 可直接设置 | ❌ 必须先设 rgb | 需注意 |
| `text_frame._element bodyPr anchor` | 垂直居中 | ✅ `set('anchor', 'ctr')` 可用 | S2 |
| `ROUNDED_RECTANGLE.adjustments[0]` | 圆角控制 | ✅ 可用 | M10 |
| `Windows mklink /J` | 无需管理员 | ✅ 实测可用 | 正确 |
| `PP_PLACEHOLDER.*` 类型体系 | title/body/image/chart | ✅ 有 TITLE/BODY/CHART/PICTURE/TABLE | M1 |
| `placeholder.placeholder_format.type` | 可识别类型 | ✅ 返回 PP_PLACEHOLDER 枚举 | 正确 |

---

## 七、修复优先级总结

| 优先级 | 编号 | 描述 |
|--------|------|------|
| **P0** | B1 | `slide_master.theme` API 不存在，需改为 XML 解析 |
| **P0** | B5 | brand.json colors 键名与 ThemeMapper 不兼容 |
| **P0** | B6 | Connector `a:ln` 查找路径错误 |
| **P0** | B8 | 模板识别需排除 output/ 子目录 |
| **P1** | B3 | `--version` 与惯例冲突 |
| **P1** | B4 | `--density` 标注为"新增"实际是"扩展" |
| **P1** | D1 | TemplateAnalyzer 缺少 XML 解析实现方案 |
| **P1** | D2 | EnterpriseRenderer 缺少模板加载降级 |
| **P1** | D8 | content.json 的 cards/image 字段与 PageContent 不兼容 |
| **P1** | I2 | §4.1 与 §12.11 的 goal_mapping 不一致 |
| **P1** | I5 | §12.11 注释有旧版残留 |
| **P2** | B2 | slide_layouts[6] 硬编码问题 |
| **P2** | B7 | TextMeasurer 经验值偏高 |
| **P2** | D3 | `--no-review` 语义不清 |
| **P2** | D4 | goal→layout 映射优先级不明确 |
| **P2** | D5 | LayoutEngine 文件位置未明确 |
| **P2** | D9 | _generate_ppt_freestyle 不存在 |
| **P2** | I1 | §10 注释 ID 21-23 应为 12-22 |
| **P2** | I3 | §7.1 与 §9.6 重复定义 |
| **P2** | M1-M10 | 各种遗漏补充 |
