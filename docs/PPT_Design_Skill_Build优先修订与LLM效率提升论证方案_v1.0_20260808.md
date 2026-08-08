# PPT Design Skill Build 优先修订与 LLM 效率提升论证方案

版本：v1.0  
日期：2026-08-08  
定位：面向通用 PPT 生产能力的 skill 修订方案，不限于视频、读书或单一行业。

## 一、核心结论

当前 `ppt-design-skill` 的真正核心不是 FreeStyle、组件库或 Enterprise Pipeline，而是 Build 模式。

Build 模式具备：

- 每页独立构图；
- 精确的 x / y / w / h 控制；
- 运行级字体和 CJK 字体控制；
- 图片裁切与效果处理；
- 自由形状、连接线、表格、图表和 OOXML；
- 阴影、渐变、3D、动画和过渡；
- 不受固定内容 JSON 布局限制的动态设计能力。

因此，skill 修订不应继续围绕“统一调用所有模块”展开，而应转向：

> 以 Build 模式为高质量 PPT 的主路径，保留其他能力的有限用途，减少 LLM 的错误选择和无效探索。

修订的目标不是增加更多说明，而是降低 LLM 的决策负担，缩短从需求到正确 `build.py` 的路径。

## 二、当前 skill 的真实结构

### 2.1 `generate_ppt()` 的实际分流

当前入口并不是单一流水线，而是多个分支：

```text
proposal=True
    → ProposalGenerator

project=...
    → EnterprisePipeline
    → 源码已标记 deprecated

普通调用
    → StoryPlanner
    → DesignDecider
    → ContentGenerator
    → PrecisionRenderer
    → FreeStyle

Build 模式
    → LLM 编写 build.py
    → Build Helpers / python-pptx / OOXML
    → 手动执行
```

Build 模式并不是 `generate_ppt()` 的一个普通参数，而是一条独立的、脚本驱动的生产路径。

### 2.2 各模块的合理定位

| 模块 | 合理定位 | 不应承担的职责 |
|---|---|---|
| Build | 高质量定制设计和最终交付 | 不应被固定模板限制 |
| FreeStyle | 快速探索、低成本草稿 | 不应代表最终视觉质量 |
| ProposalGenerator | 早期方向预览 | 不应直接等同于最终完整 PPT |
| Component Library | 结构化图表和标准组件 | 不应默认支配创意页面 |
| EnterprisePipeline | 历史企业流程 | 不应继续作为主推荐路径 |
| VI Build | 有模板和品牌规范的企业项目 | 不应替代无模板 Build |
| ui-ux-pro-max | 提供设计系统和视觉参考 | 不应直接决定 PPT 页面结构 |

## 三、现有问题为什么降低 LLM 效率

### 3.1 入口选择成本过高

LLM 需要先判断应该使用：

- FreeStyle；
- Proposal；
- `project=` EnterprisePipeline；
- Component Library；
- VI Build；
- Build Script。

其中部分路径已被标记为 deprecated，部分路径适用范围有限，导致上下文中存在大量无效选择。

### 3.2 文档、源码和能力边界不完全一致

审计中发现：

- skill 文档、渲染器和实际布局注册表的数量描述并不完全一致；
- EnterprisePipeline 和 Beautify 已标记 deprecated，但文档仍保留较完整的使用说明；
- `ui-ux-pro-max` 的文档主要面向 Web、移动端和桌面 UI，不能把触控、导航、React Native 规则全部套用于 PPT；
- QA 能力并非所有 Build 脚本都会自动经过；
- 版本管理主要接在 Enterprise 路径上，手写 Build 脚本需要自行约定输出和版本；
- 组件库的匹配逻辑更适合结构化内容，不适合所有创意构图。

这些不一致会让 LLM 反复查找、猜测和试错。

### 3.3 Build Helpers 不是完整布局引擎

Build Helpers 提供了大量实用函数，例如：

- `rect`、`rrect`、`oval`、`text`、`multiline`；
- `kpi_card`、`bar_chart`、`comparison_bars`、`donut_chart`；
- `hero_slide`、`section_divider`、`cta_slide`；
- `gradient_text`、`seal_stamp`、`circle_image`、`duotone_image`；
- `shape_3d`、`pattern_fill`、`frosted_panel`、`brush_divider`、`neon_border`、`ink_splash`；
- 动画和过渡函数。

但它们本质上是构件和辅助函数，不是完整的设计决策系统。LLM 仍然必须决定：

- 页面信息关系；
- 构图重心；
- 元素坐标；
- 图片与文字关系；
- 字体、行距和留白；
- 哪些元素应该完全自定义。

因此，skill 修订应当帮助 LLM 正确使用 Build Helpers，而不是把所有页面强行转换成 helper 组件。

## 四、修订后 LLM 效率如何提升

### 4.1 模式判断效率

在 skill 开头增加明确的决策树：

```text
是否需要像素级定制、独特构图或逐页控制？
    是 → Build
    否 ↓

是否有现成模板、LOGO、页眉页脚或品牌 Token？
    是 → VI Build
    否 ↓

是否只是快速探索或低成本草稿？
    是 → FreeStyle
```

组件库只在以下情况下调用：

- 流程图；
- 组织结构；
- KPI / 信息图；
- 时间线；
- SWOT；
- 结构化数据图表。

这样可以避免 LLM 把不适合创意设计的路径当成默认路径。

### 4.2 Build 脚本生成效率

新版 skill 应提供一个最小但完整的 Build 契约：

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from ppt_pro_max.build_helpers import *

W, H = 13.333, 7.5
C = {...}

def build_slide_01(prs):
    slide = add_slide(prs)
    # 当前页面独立设计
    return slide

def build_slide_02(prs):
    slide = add_slide(prs)
    # 当前页面独立设计
    return slide

prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)

for builder in [build_slide_01, build_slide_02]:
    builder(prs)

prs.save("output.pptx")
```

契约需要明确：

- 画布比例；
- 当前任务的设计 Token；
- 图片必须按比例裁切；
- CJK 字体设置方式；
- 每页独立函数；
- 输出命名；
- 结构检查入口；
- 是否允许直接使用 python-pptx / OOXML。

这样 LLM 不需要每次重新猜测 Build 脚本的基本骨架。

### 4.3 API 调用效率

新版 skill 应把“必须知道”和“按需查阅”分开。

#### 每次 Build 必须加载

- Build 模式说明；
- `build_helpers` 函数签名；
- 图片裁切规则；
- CJK 字体规则；
- 画布、坐标和输出规则；
- 基础结构 QA。

#### 按任务加载

- 图表 API；
- OOXML 效果；
- 动画；
- 3D；
- 组件库；
- 模板复制；
- 图片效果。

这比每次让 LLM 面对整套 1200 多行参考文档更高效。

### 4.4 调试效率

Build 模式的错误应被分类，而不是只返回“脚本失败”：

```text
API_ERROR       helper 参数错误或导入错误
ASSET_ERROR     图片、字体或素材缺失
GEOMETRY_ERROR  元素越界或尺寸异常
TEXT_ERROR      文本框不足、换行或字号问题
RENDER_ERROR    PPT 无法渲染或动画 XML 异常
DESIGN_WARNING  结构通过，但视觉上需要人工复核
```

LLM 才能针对具体问题修改 `build.py`，而不是无目的地重写整份脚本。

## 五、Build-first 的修订内容

### 5.1 修订 `SKILL.md` 的结构

建议重排为：

```text
1. 任务类型判断
2. 模式选择决策树
3. Build 模式主流程
4. Build 脚本契约
5. Build Helpers 快速索引
6. 图片和字体规则
7. 页面级设计原则
8. 结构 QA 与视觉 QA 边界
9. FreeStyle / VI Build / Component Library 的适用范围
10. 已废弃能力说明
11. 完整参考文档索引
```

Build 相关内容应排在前面，deprecated 能力不应与主路径并列展示。

### 5.2 修订 Build Helpers 的使用说明

不应只按函数名称列出 API，而应按设计任务组织：

| 设计任务 | 推荐工具 |
|---|---|
| 基础版面 | `rect`、`rrect`、`oval`、`text`、`multiline` |
| 标题和章节 | `page_header`、`hero_slide`、`section_divider` |
| 数据重点 | `kpi_card`、`bar_chart`、`comparison_bars` |
| 结构图 | Component Library 或原生图表/形状 |
| 情绪视觉 | `circle_image`、`duotone_image`、`artistic_image` |
| 特殊风格 | `brush_divider`、`ink_splash`、`neon_border`、`seal_stamp` |
| 高级效果 | OOXML、3D、渐变、动画 |

同时明确：helper 是可选构件，不是页面模板。

### 5.3 修订 UI/UX 集成边界

`ui-ux-pro-max` 在 PPT 任务中只承担：

- 视觉风格建议；
- 字体组合；
- 配色和对比度；
- 间距系统；
- 信息层级；
- 数据图表建议；
- 反模式提醒。

以下规则不应默认套用到普通 PPT：

- 44×44pt 触控目标；
- Tab / Bottom Navigation；
- React Native 栈规则；
- 表单错误反馈；
- Web 响应式断点；
- App 屏幕的交互状态。

## 六、QA 设计：以 Build 结果为对象

### 6.1 结构 QA

Build 输出后检查：

- PPT 是否能打开；
- 页面比例是否正确；
- 页面数量是否符合计划；
- 元素是否越界；
- 图片是否缺失或拉伸；
- 文本框是否明显不足；
- 字体是否存在；
- 输出路径和版本是否正确。

### 6.2 视觉 QA

必须通过渲染图或人工预览检查：

- 页面是否仍然像同一套 PPT；
- 页面构图是否重复；
- 标题是否在缩略图下可读；
- 图片裁切是否破坏主体；
- 文字和图片是否互相争夺注意力；
- 信息层级是否清楚；
- 空白是否有意图；
- 视觉效果是否服务于内容；
- 是否出现模板化、伪 UI 和装饰堆叠。

结构 QA 通过不代表视觉 QA 通过。

### 6.3 QA 的现实边界

不能把所有审美判断硬编码成自动规则。应区分：

- `fatal`：文件损坏、图片缺失、严重越界；
- `warning`：字号偏小、对比度偏低、文本可能溢出；
- `review`：页面节奏、创意、视觉疲劳和叙事效果。

`review` 必须保留人工确认，而不是由程序自动判定。

## 七、对其他能力的处理方式

### 7.1 FreeStyle

保留，但明确为：

- 快速草稿；
- 主题探索；
- 非关键页面试验；
- 内容框架预览。

不能将 FreeStyle 产出默认视为高质量交付稿。

### 7.2 Component Library

保留为结构化内容工具，但取消“复杂内容总是优先组件库”的绝对规则。

当用户需要独特视觉叙事、情绪化页面或非标准构图时，应优先使用 Build 直接设计。

### 7.3 Enterprise Pipeline 和 Beautify

既然源码已标记 deprecated，skill 中应：

- 标注 deprecated；
- 说明不建议新任务使用；
- 保留迁移说明；
- 不再把它们放在主流程示例中。

### 7.4 VI Build

继续保留，适用于：

- 企业模板；
- 品牌规范；
- 既有封面、目录、尾页；
- LOGO、页眉、页脚和水印合规。

VI Build 与无模板 Build 是两个不同的精确生产路径。

## 八、预期效率提升

这里不先承诺未经测试的百分比，而用可验证指标衡量。

### 8.1 预期改善

- LLM 更少走错模式；
- 更少调用已经废弃的路径；
- 更少把组件库当作创意版式；
- 更少出现 helper 参数错误；
- 更快形成可运行的 `build.py`；
- 修订时只改具体页面，而不是重写全稿；
- 结构错误更早暴露；
- Build 输出更容易复现和版本管理。

### 8.2 不应过度承诺的部分

- 不保证每页都有创意；
- 不保证图片选择完全正确；
- 不保证自动 QA 能判断审美；
- 不保证所有电脑的字体渲染一致；
- 不保证 UI/UX 搜索结果直接适合所有 PPT 类型。

## 九、回归测试方案

修订 skill 之前，应准备至少 6 类测试任务：

1. 商务汇报；
2. 教育课程；
3. 数据报告；
4. 产品或方案提案；
5. 品牌发布；
6. 编辑、社交媒体或视频视觉素材。

每类至少测试：

- 1 次 FreeStyle 草稿；
- 1 次 Build 高质量定制；
- 有模板的类型再测试 1 次 VI Build；
- 含图表的类型测试 Component Library 的可选路径。

比较指标：

- 模式选择是否正确；
- 脚本首次运行是否成功；
- API 错误数量；
- 结构 QA 通过情况；
- 渲染质量；
- 页面重复程度；
- 人工修改轮次；
- 最终用户满意度。

## 十、最终建议

当前不建议继续开发一个覆盖所有模块的“大工作台”，也不建议只增加一份项目说明文件。

建议直接修订 `ppt-design-skill`，但采用增量方式：

1. 先清理模式定位和 deprecated 说明；
2. 将 Build 模式提升为高质量 PPT 主路径；
3. 增加 Build 脚本契约和任务型 helper 索引；
4. 补充 Build 输出的结构 QA；
5. 明确视觉 QA 仍需渲染和人工判断；
6. 限定 ui-ux-pro-max 在 PPT 中的使用边界；
7. 让组件库成为可选结构化工具；
8. 用多类型 PPT 做回归测试；
9. 测试通过后再决定是否改变默认路由。

最终要提升的不是“skill 的信息量”，而是：

> 让 LLM 更快进入正确的 Build 路径，更少面对无效选项，更稳定地写出可运行、可修改、可检查的 PPT 设计脚本。
