# Theme-Guided LLM Design Workflow

状态：REVIEW-READY，已基于 `E:\pptx-designer` 源码审查更新（2026-08-31）

## 1. 文档目的

本文档提出 `ppt-design-skill` 的重新设计方向。

目标不是复制竞争对手的大量预制组件和固定页面模板，而是在保留现有
`pptx-designer` 能力与主题资源的前提下，让 LLM 更充分地进行视觉判断，
并让主题真正参与整个设计过程。

核心命题：

> 让系统提供更好的视觉知识、设计语言和反馈，让 LLM 做更多设计判断；
> 不用更多模板和组件替代 LLM 的判断。

## 2. 当前问题的重新定义

当前 skill 并不是缺少设计原则、主题数据或底层 API。问题在于三者之间的
连接方式：

```text
主题库：提供 palette、typography、style 和 ThemeComposer
skill：要求 LLM 写 visual direction 和 C 字典
Build Mode：让 LLM 在 blank canvas 上逐个放置通用元素
```

主题因此被压缩成若干参数，后续页面生成主要由 LLM 临场完成。常见结果是：

- 主题色被使用，但主题的视觉性格没有被使用；
- 主题只影响颜色和字体，没有影响构图、页面节奏和信息层级；
- LLM 退回标题、矩形、卡片和普通图表；
- 当页面缺少视觉锚点时，LLM 通过增加阴影、描边、边框和材质补偿；
- 局部效果增加，但整套 deck 变得重复、沉重或模板化。

因此，问题不是“高级 API 太少”，也不是“需要更多组件”，而是缺少一个
让主题持续参与 LLM 设计决策的中间层。

## 3. 方案目标

### 3.1 必须实现的目标

1. 让主题从参数集合升级为可被 LLM 理解的设计语言。
2. 让主题影响页面构图、视觉焦点、层级、节奏和组件选择。
3. 保留现有 `pptx-designer` 公共 API 的灵活组合能力。
4. 让 LLM 负责受众分析、视觉方向、页面结构和表达取舍。
5. 让最终 PNG 反馈进入设计修正，而不仅是技术错误检查。
6. 让同一主题可以跨不同领域适配，而不是绑定成固定页面模板。
7. 保持 PPTX 原生可编辑，避免用整页图片替代设计。

### 3.2 明确不做的事

- 不复制竞争对手的大量预制页面和主题专属组件。
- 不建立第二层 API 封装或新的固定调用语法。
- 不要求每页调用高级 API。
- 不设置高级效果数量、页面数量或能力预算等 KPI。
- 不把某一种效果绑定为 Hero、数据页或 closing 页的必选方案。
- 不用主题选择替代领域判断和页面目标分析。
- 不把 palette-only 变化当成结构性设计方案。
- 不把“生成成功、形状数量正确、API 被调用”当成视觉质量证明。

## 4. 设计原则

### 4.1 主题是设计语言，不是颜色包

主题至少应帮助 LLM 理解：

- 视觉主张和情绪；
- 适合的受众、场合和内容领域；
- 标题、正文、指标、图表、图片和装饰的表达方式；
- 页面密度和叙事节奏；
- 可反复使用的视觉母题；
- 应避免的表达；
- 在现有 API 中可采用的实现路径。

### 4.2 LLM 先做设计判断，再做 API 选择

正确顺序：

```text
理解 brief
→ 判断受众、领域和沟通目标
→ 选择或组合视觉方向
→ 建立整套 deck 的视觉语法
→ 为每页确定叙事角色和视觉焦点
→ 选择现有 API 和原生元素
→ 渲染并根据 PNG 反馈修正
```

错误顺序：

```text
查看可用 API
→ 选择几个高级效果
→ 想办法把效果放进页面
→ 用主题色统一
```

### 4.3 主题约束设计方向，但不取代 LLM

主题应该提供边界、语言和可用素材，但不能变成不可变的页面模板。

同一主题允许 LLM 根据内容选择不同的：

- 页面构图；
- 主视觉类型；
- 数据表达方式；
- 图像与图表比例；
- 信息密度；
- 空白和节奏。

### 4.4 视觉质量优先于技术能力展示

每个高级能力都必须回答：

```text
它解决了页面中的什么问题？
它是否强化了信息层级？
它是否符合主题和领域？
它是否让最终 PNG 更专业？
```

无法回答时，使用更简单的原生方案。

## 5. 目标架构

新的 skill 不增加大型组件层，而是增加一个轻量的主题引导层：

```text
现有主题资源
    ↓
Theme Knowledge / Theme Lock
    ↓
LLM 视觉方向与整稿艺术指导
    ↓
页面语义与构图决策
    ↓
现有 pptx-designer 公共 API
    ↓
PPTX → PDF → PNG
    ↓
主题兑现度与视觉质量反馈
```

### 5.1 现有主题资源

优先复用当前已经存在的：

- `PALETTES`；
- `STYLES`；
- `TYPOGRAPHY`；
- `recommend_styles()`；
- `ThemeComposer()`；
- 现有图表、图形、图片、文本和布局 helper；
- 已有示例 PPTX、视觉方向文档和 PNG 评审记录。

这些资源不需要重新包装成一套新的 API。

主题资源已经封装在 `pptx-designer` 中；源码审查确认，当前版本已经建立了
主题贯穿生成和渲染的运行时边界。现有主题能力分成三层：

1. 主题数据层：palette、style、typography、decoration、layout 和 effect
   preset；
2. 主题上下文层：`validate_resolved_theme()`、`set_presentation_theme()`、
   `set_slide_theme()`、`resolve_color_context()` 和语义角色别名；
3. 生成器层：`generate_ppt()`、`professional_renderer`、Build Mode helper、
   `BuildSpec`、`merge_vi_design_context()` 和 VI Build Context。

因此，skill 不应再假设主题只是一份需要手工复制的 `C` 字典，也不应重复实现
一层主题传递机制。下一步的核心是让 LLM 正确地产生并锁定主题上下文，然后
把它交给库的现有入口。

### 5.1.1 源码审计结论（2026-08-31）

对 `E:\pptx-designer\src` 和主题、Build、VI 相关回归测试的审查结论如下：

| 能力 | 当前状态 | 对 skill 的含义 |
| --- | --- | --- |
| 完整主题传递 | 已实现：`generate_ppt(theme=...)`，并返回 `theme_context` | skill 先将 Theme Lock resolve 成完整主题并显式传入，不再只传 `C` |
| presentation/slide 继承 | 已实现：presentation 默认值、slide 覆盖、显式局部覆盖 | Build Mode 代码可少传重复参数，但局部明确覆盖仍有效 |
| 语义色和字体 | 已实现：`resolve_color_context()` 及 typography/semantic role 别名 | 页面代码选择角色，避免绑定具体色值 |
| helper 主题继承 | 已覆盖 text、shapes、SVG、layout、images、charts、cards 等路径 | 不需要为每个主题建立专属组件 |
| 可复现与 fallback | 已实现来源、seed、fallback 记录；显式 atoms 可覆盖 preset | skill 必须保留 seed、source 和 fallback 证据 |
| Build Mode | 已通过主题继承和视觉应用回归 | LLM 负责构图，库负责承载主题默认值 |
| VI Build Mode | 已有 `merge_vi_design_context()`、`VIBuildSession`、模板约束和 BuildSpec | skill 必须通过受保护的 VI 合并入口保留模板锁，再做自由页面设计 |
| 主题效果兑现 | 颜色、字体、装饰、布局变体已有应用证据；部分文字/图片效果会报告 `not_applied` | QA 要区分“已应用”“未应用”和“视觉上不需要应用” |

审查所运行的主题、图表继承、Build Spec 和 VI Context 测试共 57 项，全部通过。
这意味着旧文档中“需要修订库代码才能建立主题传递边界”的结论已经过期；
对应的代码级要求请以 `docs/pptx-designer-theme-integration-requirements.md`
中的“已完成项 / 剩余项”记录为准。

需要保留的一项环境风险：直接运行系统 Python 时可能导入旧的
`site-packages\pptx_designer`，而不是 `E:\pptx-designer\src`。skill 的安装、
运行和 CI 必须确保使用升级后的包版本，否则会出现“源码已具备、实际运行未接入”
的假阴性。

### 5.1.2 对 skill 的新要求

1. 主题发现阶段使用 `recommend_styles()` 或显式风格选择，生成可解释的
   Theme Lock；不要让每页重新随机选主题。
2. 将 Theme Lock 用 `ThemeComposer.compose(...)` resolve 成完整主题，记录
   source、seed、fingerprint、包版本和模块路径；不要在下游重新手工拼接不同的
   颜色字典。
3. FreeStyle 使用 `generate_ppt(theme=resolved_theme)`；完整主题存在时不要再
   传 `style_seed` 或其他发现参数。
4. Build Mode 使用 `Presentation(theme=resolved_theme, strict_theme=True)`，让
   helper 自动继承主题；只有表达页面语义时才做局部覆盖。
5. VI Build Mode 使用
   `merge_vi_design_context(template_context, resolved_theme, page_context)`，
   先处理冲突，再交给 `VIBuildSession` 与 BuildSpec。
6. 生成后读取 `theme_context`、来源、seed、fallback、`not_applied` 和 ignored
   discovery arguments 诊断，
   但最终是否成功必须通过 PDF/PNG 视觉审查判断。
6. 如果主题没有改变页面的构图、焦点、密度或节奏，问题优先回到 LLM 的
   visual thesis 和 page plan，而不是继续增加库组件。

### 5.1.3 不应通过修改库代码解决的问题

以下问题不应通过在库内增加固定页面模板来解决：

- 不应为每个主题预设一整套封面、数据页和流程页；
- 不应把主题绑定到固定的 API 组合；
- 不应让主题运行时替 LLM 决定页面叙事；
- 不应把主题字段扩展成新的二次 DSL；
- 不应通过默认开启更多阴影、描边、渐变来制造“炫酷”。

库代码的职责是保证主题表达能力可被稳定消费，LLM 的职责仍然是选择和
组合这些能力。

### 5.2 Theme Knowledge

Theme Knowledge 是面向 LLM 的主题知识表达，不是新的组件层或固定页面
DSL。它可以由库内已有主题数据和少量设计元数据生成；如果现有库无法提供
某些语义信息，应优先补充主题元数据和传递机制，而不是新增大量组件。

它负责把现有主题资源解释成可用于设计判断的内容，包括：

- 主题名称和视觉主张；
- 颜色的角色关系，而不是单独的色值；
- 字体的层级和使用场景；
- 形状、线条、角度、尺度和空间倾向；
- 图表、图片、指标和正文的视觉语言；
- 适合的页面角色和内容密度；
- 主题的标志性元素；
- 主题的禁用模式和风险；
- 现有公共 API 的候选实现方式。

### 5.3 Theme Lock

在生成前形成一份本次任务的主题锁定结果，作为 LLM 和 QA 的共同依据。

建议最小结构如下：

```yaml
theme_lock:
  version: 1
  status: active
  source_theme: "主题候选或 ThemeComposer 输入"
  visual_thesis: "一句话视觉主张"
  audience_promise: "观众应感受到什么"

  intended_tokens:
    palette:
    typography:
    decoration:
    layout:
    safe_margin:
    spacing_rhythm:
    corner_treatment:

  visual_grammar:
    scale_contrast:
    density_rhythm:
    alignment_logic:
    image_language:
    chart_language:
    decoration_language:
    whitespace_purpose:

  deck_rhythm:
    opening:
    development:
    evidence:
    transition:
    closing:

  page_guidance:
    hero:
    overview:
    comparison:
    chart:
    process:
    timeline:
    editorial:
    closing:

  preferred_existing_apis:
  optional_existing_apis:
  forbidden_patterns:
  fallback_rules:
```

这里的 `page_guidance` 是设计建议，不是固定模板；`preferred_existing_apis`
是候选能力，不是必须调用清单。

Theme Lock 不等于 resolved theme。后者必须是
`ThemeComposer.compose(...)` 的完整返回值，单独存为可复现输入；它包含库所需的
`name`、`atoms`、`colors`、`semantic_roles`、`typography`、`decoration`、
`layout_variant` 和 `source`。将上述 Theme Lock YAML 直接传入
`generate_ppt(theme=...)` 是无效的。

### 5.4 主题版本与持续修订

Theme Lock 必须是项目级的当前状态，而不是只存在于一次对话或某个脚本中的
临时变量。用户确认风格修订后，递增版本并完整替换当前 Theme Lock；后续生成
只读取当前版本。旧版本可以保留用于追溯，但不能自动作为输入。

每次生成前检查 task-init、Theme Lock、resolved-theme 文件、页面计划和生成脚本
的版本与 fingerprint 是否一致；检查实际导入的 `pptx-designer` 包版本和模块路径。
局部 `C` 或 typography 覆盖必须由当前 resolved theme 派生并说明页面语义理由。
这样即使项目经过多轮内容或风格修改，也不会让早期页面脚本把主题带回旧状态。

## 6. LLM 的职责边界

### 6.1 LLM 必须负责的判断

- brief、受众、场合和领域分析；
- 观众需要理解、记住、感受或决定的内容；
- 主题选择和视觉方向解释；
- 整套 deck 的视觉主张；
- 页面顺序和叙事节奏；
- 每页的页面角色、核心 takeaway 和证据；
- 每页的视觉焦点和构图方式；
- 信息密度与留白的取舍；
- 现有 API 与原生元素的组合；
- 渲染后的视觉判断与返工方向。

### 6.2 系统应该帮助 LLM 的部分

- 展示主题候选及其视觉知识；
- 固定已确认的主题 token；
- 提供现有公共 API 的准确签名和能力范围；
- 提供现有示例和可观察的视觉证据；
- 提供页面容量、字体和渲染限制；
- 记录实际生成的主题和资源使用；
- 输出清晰的 PNG 预览，帮助 LLM 进行视觉复盘。

### 6.3 系统不应该替 LLM 决定的部分

- 每页必须使用哪种高级效果；
- 每套 deck 必须调用多少能力；
- 所有页面必须使用同一种布局；
- Hero 必须使用描边、阴影或边框；
- 数据页必须使用某一种图表；
- 主题必须绑定一套不可变化的页面模板。

## 7. 重新设计后的工作流

### 阶段 A：理解 brief

先确定：

- 主题、受众、场景和语言；
- 内容领域和证据要求；
- 观众要做出的行动或判断；
- 页数和演示时长；
- 图片、品牌、模板和编辑性约束。

### 阶段 B：主题发现

使用现有主题资源产生候选方向，但不直接把候选结果当作最终设计。

LLM 应解释：

- 为什么该主题适合当前受众和内容；
- 主题会如何影响标题、图表、图片和页面节奏；
- 主题可能产生的风险；
- 哪些地方需要克制或使用 fallback。

### 阶段 C：建立 Theme Lock

在内容规划和代码生成前，锁定：

- visual thesis；
- 主题 token；
- 视觉语法；
- 页面角色指导；
- deck rhythm；
- 禁用模式。

Theme Lock 是本次任务的设计依据。它不要求新建渲染组件，但需要依赖库内
的主题传递机制，才能从 LLM 决策真正落到最终 PPTX。

### 阶段 D：页面级设计

每页先定义：

- 页面目标；
- 一句话 takeaway；
- 证据和内容；
- 页面角色；
- 视觉焦点；
- 构图关系；
- 信息密度；
- 与前后页的节奏关系。

然后再选择现有的文本、形状、图表、图片和 diagram API。

### 阶段 E：使用现有 API 生成

Build Mode 仍然允许 LLM 精确控制坐标和组合，但代码必须继承 Theme Lock：

- 颜色使用锁定角色，不随页面临时漂移；
- 字体使用锁定层级；
- 间距和安全边界保持一致；
- 图表、图片和装饰遵循主题语言；
- 高级效果只在有明确页面问题时使用；
- 重要信息保持原生可编辑。

### 阶段 F：PNG 视觉复盘

除了技术缺陷，还要回答：

- 主题是否一眼可感知；
- 页面是否体现主题的视觉主张；
- 主题是否贯穿而没有僵化重复；
- 页面焦点是否清晰；
- 页面之间是否有节奏变化；
- 是否退化为通用卡片或默认布局；
- 是否出现为了炫酷而加入的无意义效果；
- 是否符合领域和受众；
- 是否保持可读性与编辑性。

## 8. 对现有生成模式的影响

### Build Mode

Build Mode 是重点验证对象。

保留精确坐标和公共 API 的自由组合，但新增 Theme Lock 作为生成前的设计
约束和生成后的视觉验收依据。

### FreeStyle Mode

向 `generate_ppt()` 传入锁定的主题信息和视觉方向，但不要求源码级 API
调用，也不要求固定页面结构。重点验证主题是否能影响最终结果，而不是
只出现在 prompt 中。

### VI Build Mode

主题必须服从模板 DNA 和品牌约束。Theme Lock 需要区分：

- 可继承的模板规则；
- 可调整的主题表达；
- 禁止修改的品牌元素。

当 `merge_vi_design_context()` 报告 conflicts 时，不能仅以“用户已批准”为由继续
使用被拒绝的合并结果。必须先形成新的已批准 template context，明确修改或移除
相应 lock，再重新合并并保留审计记录。

## 9. QA 设计

### 9.1 主题兑现度

至少检查：

- Theme Lock 中的视觉主张是否在 PNG 中可见；
- 颜色、字体、网格和图表语言是否一致；
- 页面构图是否体现主题，而不是仅仅换色；
- 主题是否支持页面目标，而不是妨碍信息表达；
- 装饰是否有视觉或语义作用。

### 9.2 LLM 设计质量

- 是否根据页面目标选择视觉形式；
- 是否有明确 focal point；
- 是否形成有节奏的整套 deck；
- 是否有合理的尺度、密度和留白变化；
- 是否避免重复卡片、随机渐变和无意义材质；
- 是否适合领域、受众和使用场景。

### 9.3 技术交付质量

- 无溢出、重叠、裁切和乱码；
- 图表语义和数据正确；
- 图片比例和裁切正确；
- PPTX 可打开；
- 重要文字、形状、图表和图示保持可编辑；
- PDF 和 PNG 与 PPTX 视觉一致。

### 9.4 能力追踪

能力追踪只记录实际使用的 API：

```yaml
capabilities_used:
  - capability: comparison_bars
    slide: 4
    problem_solved: "比较两个阶段的差距"
    theme_role: "证据页的主视觉"
    fallback: "native editable bars"
    png_evidence: "slide-04-chart-readable"
```

没有使用高级 API 不构成失败；使用了却无法说明视觉价值，才需要复查。

## 10. 验证计划

验证重点不是“主题字段是否被写入”，而是主题引导是否提升最终设计质量。

### 10.1 第一阶段：同一 API 的流程对比

固定：

- brief；
- 内容；
- 页面数量；
- `pptx-designer` 版本；
- 渲染链路；
- 随机种子。

比较：

- A：当前普通生成流程；
- B：加入 Theme Knowledge 和 Theme Lock 的新流程。

两组都使用同一批现有公共 API，不增加组件，不增加特效配额。

### 10.2 第二阶段：跨领域验证

至少覆盖：

- 科技/产品发布；
- 企业战略或商业汇报；
- 科研或学术表达；
- editorial / 品牌叙事。

每类都要检查主题是否真正改变表达，同时没有破坏领域语义。

### 10.3 第三阶段：跨生成模式验证

- Build Mode；
- FreeStyle Mode；
- VI Build Mode。

如果某一模式无法传递 Theme Lock，需要明确记录为模式边界，而不是假设
所有模式都已经支持。

### 10.4 通过标准

新流程只有在以下条件同时满足时才算有效：

- 整体视觉专业度提升；
- 主题可感知且贯穿整套 deck；
- 页面结构和节奏更有意图；
- 没有新增 P0/P1 级可读性或导出问题；
- 没有以主题效果换取信息质量下降；
- 仍保持原生可编辑；
- LLM 能解释关键设计取舍。

## 11. 主要风险

### 风险 1：Theme Lock 变成新的僵化模板

控制方式：只锁定设计语言和边界，不锁定固定页面或固定 API 组合。

### 风险 2：主题知识过于抽象，LLM 仍然只记住颜色

控制方式：主题知识必须包含页面构图、视觉焦点、节奏和反例，并通过 PNG
验收验证主题是否真正体现。

### 风险 3：为了保持灵活性，结果重新退化成空白画布

控制方式：保留主题驱动的页面级设计建议、视觉母题和已有示例证据，不能
只提供 token。

### 风险 4：LLM 仍然过度使用高级效果

控制方式：以页面问题和 PNG 结果为判断依据，不设能力数量目标；明确要求
优先使用能解决问题的最简单方案。

### 风险 5：主题表达与 PPTX 导出不一致

控制方式：把最终 PPTX-PDF-PNG 渲染作为必经路径，不以 HTML 或 Python
运行成功代替视觉验收。

## 12. 下一轮细节审查问题

源码能力边界已经确认，下一轮应审查 skill 的落地质量：

1. Theme Knowledge 哪些字段应直接从库的主题结果生成，哪些需要人工补充？
2. Theme Lock 是否落盘为 YAML/JSON，并作为 PNG QA 的对照输入？
3. 如何用跨主题 PNG 基线区分“只是换色”和“真正改变视觉语言”？
4. VI Build 中主题与模板 DNA 的冲突优先级是否符合品牌交付要求？
5. skill 安装和 CI 如何保证不会误用旧版 `site-packages`？
6. 第一组 A/B 测试应选择哪个已有案例作为最小验证样本？

## 13. 当前结论

本方案暂不要求增加大量组件，也不要求重建主题库。

第一阶段只需要验证一个核心假设：

> 在使用相同 `pptx-designer` API 的前提下，如果让 LLM 获得结构化的主题
> 视觉知识、Theme Lock 和 PNG 反馈，它能否比当前流程做出更有主题感、更
> 有层次、更专业的 PPT。

如果这个假设不成立，再判断问题属于 skill 决策、主题上下文使用、渲染器
应用还是导出链路。主题传递基础设施已经具备，后续不应再以扩张组件或 API
作为默认答案，而应先用 PNG 证据定位问题。
