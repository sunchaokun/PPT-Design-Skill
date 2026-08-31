# PPT Design Skill 使用手册

## 1. 核心流程

正式交付必须遵循：

```text
需求确认
→ PPT 结构设计
→ 视觉方案设计
→ 用户确认方向
→ 生成 PPTX
→ PPTX → PDF → PNG
→ LLM 逐页检查 PNG
→ 修订并重新导出
→ 用户确认最终效果
→ 交付
```

Python 运行成功不是完成条件。PNG 视觉检查通过后才可以进入最终交付。

当前示例适用于 `pptx-designer 1.0.0b10`。仓库维护 6 个完整案例，覆盖技术架构、基础设施研究、
科学叙事、建筑文化、城市策略和高定美妆编辑。案例库见
[examples/README.md](../examples/README.md)，在线预览与下载见
[案例画廊](https://sunchaokun.github.io/PPT-Design-Skill/)。

在生成前，先把用户需求写成验收合同，并在 PNG 检查时逐项对照：

| 编号 | 用户要求 | PNG 中应看到的证据 | 优先级 |
|---|---|---|---|
| R1 | 视觉上克制、专业 | 统一色彩、明确层级、足够留白 | MUST |
| R2 | 解释三阶段流程 | 有三个可读且有顺序关系的阶段 | MUST |
| R3 | 保留品牌 Logo | 位置、比例和清晰度符合要求 | SHOULD |

以上只是格式示例，实际内容必须来自用户需求。每个 `MUST` 条件都必须在
导出的 PNG 中找到可观察证据；发现不满足时必须回到源码或内容修订。

## 2. 安装

请先克隆仓库，并从仓库根目录运行安装器。安装器会自动安装已发布的
`pptx-designer` Python 库，同时安装 Skill 文件包：

```powershell
git clone https://github.com/sunchaokun/PPT-Design-Skill.git
cd PPT-Design-Skill

python installer/install.py --platform claude --force
python installer/install.py --platform codex --force
python installer/install.py --platform opencode --force
python installer/install.py --platform deepseek-harness --force
python skill/scripts/check_runtime.py
```

也可以使用 `--platform all` 安装到所有已支持平台。DeepSeek Harness 的
技能根目录为 `~/.dsh/skills`（全局）和项目下的 `.dsh/skills`。

## 3. 三种模式

### Build Mode

交付级、逐元素精确控制，使用公开的 `pptx_designer` Build API。

### FreeStyle Mode

使用 `generate_ppt()`。自然语言 `query` 和结构化 `content` 是同一 FreeStyle
模式的两种输入形式；需要复现主题时，应先用 `ThemeComposer` 解析完整主题，
再通过 `theme=` 传入。

### VI Build Mode

使用企业模板或品牌母版时进入 VI Build：使用 `extract_design_context()` 分析
模板、确认框架页和可写槽位，使用 `compile_atomic()` 添加内容页，再通过
`VIBuildDelivery` 完成交付并对完整结果导出 PNG 检查。不能默认承诺复杂
PowerPoint master、SmartArt、动画和私有 OOXML 的完全复刻。详细流程见
[template-brand.md](../skill/references/template-brand.md)。

## 4. FreeStyle 生成

### FreeStyle：`generate_ppt()`

FreeStyle 是库的自动生成路径，有两种输入形式：

```python
from pptx_designer import generate_ppt
from pptx_designer.renderer.theme import ThemeComposer

# 主题驱动的快速草稿
theme = ThemeComposer().compose(style="dark-tech", seed=17)
result = generate_ppt(
    "AI startup investor pitch",
    theme=theme,
    output="output/draft.pptx",
)

# 内容驱动的结构化草稿
generate_ppt(
    content={
        "title": "Q4 Revenue Review",
        "pages": [
            {"goal": "hook", "title": "Q4 2026", "subtitle": "Record quarter"},
            {"goal": "data", "title": "Key metrics", "bullets": ["Revenue: $12.8M"]},
        ],
    },
    theme=theme,
    output="output/structured.pptx",
)
```

两者都是 FreeStyle；`content` 只提供更强的内容和页面目标控制，不提供逐
元素坐标控制。传入完整 `theme` 后，不要同时传 `style`、`palette`、`fonts`、
`decoration`、`layout`、`mood` 或 `style_seed`；这些发现参数会被忽略并产生
警告。跨进程或持久化主题前，可用 `validate_resolved_theme(theme)` 校验。

### Build Mode：可复现的精确设计

交付级、品牌化、复杂图示和像素级布局使用 Build Mode：

```python
from pptx_designer import Presentation
from pptx_designer.renderer.theme import ThemeComposer
from pptx_designer.tools.cards import kpi_card
from pptx_designer.tools.layout import page_header
from pptx_designer.tools.shapes import rect

theme = ThemeComposer().compose(style="professional", seed=17)
prs = Presentation(theme=theme, strict_theme=True)
slide = prs.slides.add_slide(prs.slide_layouts[6])
page_header(slide, "Q4 Revenue Report", "Financial Summary")
kpi_card(slide, 1.0, 2.0, 3.5, 1.5, "$12.8M", "Revenue", "+23%")
rect(slide, 0.5, 6.8, 12.3, 0.08, fill="primary")
prs.save("output/report.pptx")
```

`Presentation()` 默认创建 16:9 画布（13.333 × 7.5 英寸）。完整主题会让公共
helper 自动继承颜色和字体；显式传入的 `C`、颜色或字体仍可用于刻意的局部覆盖。

## 5. 交付与质量门

- 先定义客户级视觉目标、视觉锚点和页面完成度，再设计页面结构和视觉方向。
- 先确定受众和页面目标，再选择组件。
- 交付级 PPT 先输出页面结构和视觉方案，得到用户确认后再完整生成。
- 一页一个主要结论；过多内容必须拆页或改成图表、表格、流程图。
- 大面积留白必须承担明确的焦点、节奏或层级功能；如果只是内容规划不足，
  必须返工。
- 6 个以上要点通常需要双栏、卡片、表格或图示。
- 科研、学术、医疗内容不能默认使用商业融资页面范式。
- 全 deck 锁定一套颜色、字体、间距和组件系统。
- 图片必须保持比例，重要信息必须保持 PowerPoint 原生可编辑。
- 禁止使用虚假精确数据、无来源结论和无意义装饰。

完整设计规则见 [SKILL.md](../skill/SKILL.md)。

## 6. VI Build 和模板

VI Build 的模板 token、框架页保护和验收规则见
[template-brand.md](../skill/references/template-brand.md)。

## 7. PNG 视觉检查

```powershell
powershell -ExecutionPolicy Bypass -File skill/scripts/render_pptx.ps1 `
  -InFile output/report.pptx `
  -OutDir output/report-rendered
```

LLM 必须查看导出的 PNG，检查：

- 标题和信息层级；
- 文本可读性、溢出、遮挡和裁切；
- 页面密度和留白；
- 大面积留白是否承担明确的叙事功能；如果只是因为内容规划不足而空置，
  必须标记为 `NEEDS_REVISION`；
- 每页是否都有清晰的视觉锚点和完整结论；“极简”不能成为页面未完成的理由；
- 图表、图示和图片的清晰度；
- 页间一致性；
- 是否符合用户指定的场景和风格；
- 是否存在低级模板感或 AI 生成感。

发现问题后必须修改源码，重新生成 PPTX、PDF、PNG，再次检查。

视觉检查采用两道门：第一道先判断整体视觉效果和客户级完成度；第二道再
检查溢出、遮挡、不可读、需求遗漏、数据/引用和可编辑性等严重缺陷。早期
的基础缺陷检查继续保留，但视觉效果不再被它淡化。

## 8. 运行时依赖

```powershell
python installer/install.py --platform all --force
python skill/scripts/check_runtime.py
```

核心依赖由 `pptx-designer` 安装。PPTX 到 PDF 的渲染优先使用 PowerPoint
COM；无 PowerPoint 时需要 LibreOffice 和 Poppler。
