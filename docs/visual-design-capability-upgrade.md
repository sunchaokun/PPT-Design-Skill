# PPT Design Skill 视觉设计能力升级方案

状态：开发设计文档

版本：v1.0

更新时间：2026-09-04

## 1. 文档目的

本文件定义一次性升级 `ppt-design-skill` 视觉设计能力的目标架构、设计资产、提示词编排、PPT-first / SVG-assisted 实现方式、执行流程、验收标准、可行性和预期效果。

本方案基于对本仓库和 `ppt-master` 工作流的源码级对照分析形成。目标不是复制 `ppt-master` 的实现，而是把我们现有的：

- brief-first 需求确认；
- Theme Lock；
- Build / FreeStyle / VI Build 三种模式；
- `pptx_designer` 公共 API；
- PPTX → PDF → PNG 视觉验收；
- 原生可编辑 PowerPoint 输出；

升级为一套更有艺术指导能力、更容易生成高完成度页面、同时保持可复现和可审查的设计系统。

本方案不把能力拆成“先做基础版、以后再补视觉能力”。目标是一次性定义并交付完整闭环。内部仍然存在文件依赖和执行顺序，但它们是实现顺序，不是产品能力的分期取舍。

## 2. 核心判断

### 2.1 当前问题不是缺少模板

当前仓库已经具备主题、设计方向、页面 archetype、Theme Lock 和三种生成模式。问题在于，这些资产更多描述“应该遵守什么”，没有足够明确地描述“这一页具体应该怎么构图”。

因此，LLM 在面对陌生主题时容易退回到安全结构：

```text
标题 + 几个卡片 + 一张图 + 底部装饰线
```

这会导致页面技术上正确、信息也完整，但视觉表现偏保守、同质化、缺少可识别的艺术指导。

### 2.2 SVG 不是自动产生高级感的原因

SVG 只提供更大的表现空间。若没有构图语法、素材策略和页面级示例，LLM 仍然会在 SVG 或 PowerPoint 中绘制矩形、平均分栏和普通卡片。

高级感来自以下链路，而不是来自 SVG 本身：

```text
Communication Job
  → Art Direction
  → Visual Grammar
  → Composition Recipe
  → High-fidelity Page Prototype
  → PPT-native Execution + optional SVG assistance
  → Rendered Review
  → Directional Revision
```

### 2.3 `ppt-master` 的真正优势

对本机 `ppt-master` 源码、实际项目产物和远程仓库 HEAD 的核对表明，它的视觉优势主要来自以下组合：

1. 角色化的设计决策链：Strategist、Template_Designer、Executor、Image_Generator 等角色各自承担不同判断；
2. 设计规格和执行锁：`design_spec.md`、`spec_lock.md` 将视觉命题转成可执行约束；
3. 风格级视觉语法：风格文件不仅定义色彩，还定义构图几何、材质、深度、图像处理和装饰纪律；
4. 页面级构图目录：页面先选择信息关系和主构图，再添加裁切、光影、材质和跨页关系；
5. PPT-first 的页面实现：PowerPoint 原生对象承担主要内容和构图，SVG 只辅助局部复杂视觉；
6. 第一页方法验证和最终质量门：P01 先验证方法级偏差，再连续生成其余页面，最后集中修订。

结论：我们需要补的是“艺术指导中间层”，不是简单增加模板数量。

## 3. 目标与非目标

### 3.1 目标

升级后，普通用户只需提供主题、受众和内容，Skill 就能在不依赖用户掌握深度设计 Prompt 的情况下：

- 自动判断适合的视觉方向；
- 生成两到三套真正不同的完整方向提案；
- 为每页选择有沟通理由的构图模式；
- 生成具有明显视觉个性的原生 PPTX 页面，并在必要时使用 SVG 辅助复杂视觉元素；
- 在不同页面类型间保持统一的视觉世界；
- 对“普通模板感”进行方向性识别和返工；
- 保留重要文本、图表、形状和图示的可编辑性；
- 让生成过程可追溯、可复现、可审查。

### 3.2 非目标

本次方案不追求：

- 复制 `ppt-master` 的文件、代码或专有实现；
- 为每个主题制作一个固定 PPT 模板；
- 用随机装饰强行制造“炫酷”；
- 用整页截图替代可编辑 PPT 内容；
- 让所有领域都使用高对比、强动效或复杂装饰；
- 用自动评分模型替代最终的 PNG 视觉审查；
- 为了视觉效果牺牲事实准确性、阅读性和品牌约束。

## 4. 目标架构

升级后的完整链路如下：

```text
用户 Brief / Source Material
          ↓
Brief & Domain Interpreter
          ↓
Strategic Direction Composer
  ├─ audience outcome
  ├─ communication mode
  ├─ visual thesis
  ├─ style rendering pack
  └─ direction candidates × 3
          ↓
Visual Grammar Resolver
  ├─ composition family
  ├─ page rhythm
  ├─ motif / material
  ├─ image treatment
  └─ forbidden patterns
          ↓
Page Plan + Composition Recipes
          ↓
PPT-native Build Executor
  ├─ native PowerPoint visual layer
  ├─ optional SVG assistance layer
  ├─ native editable text/data
  ├─ prepared imagery/assets
  └─ Theme Lock inheritance
          ↓
P01 Method Gate
          ↓
Full Deck Generation
          ↓
PPTX → PDF → PNG
          ↓
Visual Effect Gate + Defect Gate
          ↓
Directional Revision or Local Revision
          ↓
Final Delivery
```

### 4.1 三层设计资产模型

现有 Theme Lock 应继续保留，但增加两个上层资产层：

| 层级 | 解决的问题 | 典型内容 |
|---|---|---|
| Theme Tokens | 用什么颜色、字体和基础组件 | palette、typography、spacing、radius、chart language |
| Visual Grammar | 页面如何组织视觉关系 | asymmetry、scale contrast、image registration、depth、motif |
| Case-derived Prototypes | 具体页面如何长出来 | 直接复用真实案例的 PPT / PNG 样片、页面骨架、Recipe 和对象边界；必要时附局部 SVG |

Theme Tokens 负责一致性，Visual Grammar 负责艺术指导，Case-derived Prototypes 负责降低 LLM 的构图不确定性。

### 4.2 案例与原型的单一事实源

案例和原型不应维护两套独立资产。一个经过验证的案例同时承担两种角色：

- **Case**：面向用户展示和下载的完整 PPTX、PDF、PNG、源码和内容上下文；
- **Prototype**：指向该案例中一个或多个页面的可检索元数据视图，记录 page job、Composition Family、Rendering Pack、Recipe 和 object map。

Prototype 不复制案例的 PPTX、PNG 或图片资源，也不另行制作一份视觉相同但内容不同的“模板”。它只保存索引和解释：

```text
Case source of truth
  ├── source PPTX / build source
  ├── rendered PNG / PDF
  ├── assets
  └── case documentation
          ↓
Case-derived prototype view
  ├── case_id
  ├── slide_ids
  ├── visual grammar tags
  ├── composition IDs
  ├── recipe summary
  └── editable object map
```

未来新增案例时，应同时完成案例登记和原型标注；未来补齐构图覆盖时，优先从已有案例中发现可复用页面，只有确实缺少视觉先例时才新设计页面。案例展示、模型参考和回归验证必须引用同一份源文件，避免三者发生视觉漂移。

## 5. 资产目录设计

建议在 `skill/` 下新增以下目录：

```text
skill/
├── references/
│   ├── visual-composition-catalog.md
│   ├── visual-rendering-packs/
│   │   ├── editorial-luxe.md
│   │   ├── dark-cinematic-tech.md
│   │   ├── architectural-material.md
│   │   ├── scientific-atlas.md
│   │   ├── brutalist-data.md
│   │   └── ...
│   ├── composition-recipes/
│   │   ├── cover.md
│   │   ├── evidence.md
│   │   ├── process.md
│   │   ├── comparison.md
│   │   ├── data.md
│   │   └── transition.md
│   └── visual-execution.md
├── templates/
│   ├── case-prototypes/
│   │   └── README.md                # mapping contract; no duplicated case assets
│   └── case-prototype-index.json    # searchable prototype view
└── scripts/
    ├── inspect_visual_assets.py
    ├── validate_visual_pack.py
    └── refresh_case_prototypes.py
```

这些目录属于 Skill 的设计知识和执行资产，不与用户项目的 `ppt_tasks/` 混在一起。项目运行时只复制或引用被选中的资产，避免把整个目录加载进模型上下文。案例本体是唯一资产源：仓库中的案例目录（首轮可直接复用 `examples/new_examplex/`）保存 PPTX、渲染结果、Recipe、object map 和局部 SVG；`case-prototype-index.json` 以及其中的 prototype record 只保存索引和引用。安装到其他环境时，可以整体复制一次被选中的案例目录，或配置一个外部案例库根目录，但不得按 prototype 再复制一份。

### 5.1 Visual Rendering Pack

每个 Rendering Pack 应该回答“这个视觉世界如何被画出来”，而不是只给一个风格名称。

建议采用以下结构：

```markdown
# Editorial Luxe

## Identity
- pack_id: editorial-luxe
- best_for: luxury, architecture, culture, premium proposals
- compatible_domains: architecture, brand, strategy, editorial

## Visual Thesis
用克制的画册秩序承载一个具有材质和空间感的主视觉。

## Composition Behavior
- asymmetrical split rather than equal columns
- one dominant visual mass per page
- title block occupies the quiet field
- recurring contour appears at different scales

## Material / Depth
- limestone, paper, glass, soft shadow, architectural plane
- depth from layering and opacity, not generic drop shadows

## Typography Behavior
- quiet sans for body
- high-contrast display scale for key statements
- short labels with deliberate tracking

## Image Treatment
- full-fidelity crop
- registered image fragments
- soft atmospheric wash
- no unrelated stock imagery

## Composition Recipes
- asymmetrical title field
- image as architectural plane
- detached contour windows
- editorial figure crossing a column edge

## Anti-patterns
- three equal feature cards
- centered title plus decorative gradient
- evenly distributed ornaments
- generic blue technology background

## Prompt Fragment
...

## Review Focus
...
```

每个 Pack 必须同时包含：视觉命题、构图行为、材质/深度、字体行为、图片处理、页面配方、反模式、Prompt 片段和验收重点。

### 5.2 Composition Catalog

构图目录应使用“沟通任务 + 结构 + 可选修饰”的形式，而不是单纯的模板名称。

目录采用三层分类，避免把不同维度混成互斥 Family：

```text
Page Job（要完成什么沟通任务）
  → Composition Structure（空间如何组织）
      → Modifier / Cross-page Behavior（如何变体、延续或加强）
```

Recipe 先选择 `page_job`，再选择 `primary_structure`，最后组合 `modifiers` 和 `cross_page_behavior`。`cover`、`data` 属于 Page Job；`asymmetric split`、`image as canvas` 属于 Structure；`cross-page continuity` 属于 Cross-page Behavior，不再把它们当作同层级的互斥模板。

示例：

| ID | 主结构 | 适合沟通任务 | 视觉特征 |
|---|---|---|---|
| `c1-01` | Full-bleed title field | 封面、章节入口 | 满幅主视觉 + 原生标题叠加 |
| `c1-02` | Quiet field / dominant mass | 高级封面、观点页 | 大面积安静区 + 单一大尺度对象 |
| `c2-01` | Asymmetric split | 解释、提案、证据 | 非对称分栏，避免均匀两列 |
| `c2-02` | Image as canvas | 场景解释、空间叙事 | 图片成为空间背景，文本绑定到安全区域 |
| `c3-01` | Detached shape sequence | 过程、演进、视觉叙事 | 图像或对象被拆成多个有节奏的片段 |
| `c3-02` | Irregular mosaic | 案例、作品、素材集合 | 不等尺寸图块组成一个整体 |
| `c3-03` | Curve array | 进程、层级、连续性 | 元素沿弧线或波形排列 |
| `c4-01` | Oversized numeral / glyph | 数据、章节、强调 | 超大低透明度字符作为空间锚点 |
| `c4-02` | Editorial pull quote | 观点、研究、叙事 | 引文跨越列边界，成为视觉事件 |
| `c5-01` | Cross-page camera move | 连续场景、地图、产品 | 同一视觉对象跨页推近、平移或换景 |

目录中的每个条目应包含：

- page jobs；
- recommended content texture；
- geometry recipe；
- compatible modifiers；
- editable-object boundary；
- common failure modes；
- compatible domains；
- one or more page-composition prototypes。

### 5.3 Composition Recipe

每个页面在执行前必须解析成一张 Composition Recipe：

```yaml
page_id: P04
page_job: explain-mechanism
primary_structure: c2-02
modifiers:
  - M2-01-directional-scrim
  - M3-05-contour-echo
visual_anchor: right-side registered system diagram
occupied_zones:
  - left: title and one-sentence claim
  - right: dominant diagram field
  - bottom: source / caption
content_texture: one-claim-plus-evidence
motif_behavior: the contour recurs at smaller scale from P02
avoid:
  - equal-width cards
  - centered diagram with four callout boxes
editable_boundary:
  - text: native
  - data: native chart/table when applicable
  - decorative geometry: SVG or native shape by fidelity decision
```

Recipe 的作用是让 Executor 不再从空白画布自由猜测，而是把创意集中在少数高影响的构图决策上。

### 5.4 Case-derived Page Prototype

案例派生页面原型不是最终内容模板，也不是要求未来页面 1:1 复制。它应直接指向一个已存在并经过审查的案例页面：使用案例中的 PPTX、PNG/PDF 和源码作为事实源，再通过轻量 metadata 解释其视觉方法。必要时可以引用局部 SVG 素材，但不得复制一份独立原型资产。它必须展示：

- 页面比例和 safe area；
- 主视觉对象的尺度关系；
- 标题与视觉之间的距离；
- 主要装饰或材质如何形成层次；
- 空白区域的目的；
- 图片裁切和跨对象注册关系；
- 可编辑文本和数据的边界；
- 哪些对象应使用 PowerPoint 原生形状、文本、图表和图片；
- 哪些局部效果才值得使用 SVG 辅助。

每个 Visual Rendering Pack 至少应通过案例派生原型覆盖以下页面角色：

- cover / opening；
- section transition；
- evidence / explanation；
- comparison / sequence；
- data / chart；
- ending / CTA。

原型引用的案例可以使用占位文本和示例图，但必须保留真实的构图逻辑。原型的作用是传递空间和视觉语法，不是提供可直接替换的模板句子。案例的主要实现应能映射到可编辑 PPT 对象；SVG 只记录那些用原生 PowerPoint 难以稳定表达的局部复杂几何、材质或遮罩效果。每个案例还必须记录图片、字体和 SVG 的来源、许可证、分发权限、客户专属/敏感内容状态，以及是否允许进入模型上下文；不具备相应权限的素材不得随 Skill 安装或被原型检索加载。

## 6. Prompt 与角色编排

### 6.1 Prompt 不应是一段巨型总提示词

应该采用“路由 + 定向加载”的方式：

```text
Global workflow rules
  → selected domain paradigm
  → selected rendering pack
  → selected composition catalog entries
  → selected page-composition prototype summaries and previews when needed
  → current page recipe
  → executor instructions
```

这样做有三个好处：

- 避免所有风格和构图知识同时进入上下文；
- 让当前页面只继承真正相关的规则；
- 让不同领域拥有不同的视觉纪律，而不是全局套同一个审美。

### 6.2 Strategist Prompt 增强

Strategist 需要新增以下输出：

```yaml
visual_direction:
  visual_thesis: ...
  audience_promise: ...
  rendering_pack: editorial-luxe
  visual_grammar:
    scale_contrast: high
    asymmetry: intentional
    material_depth: layered-plane
    image_behavior: registered-crops
    decoration_behavior: sparse-structural
  composition_vocabulary:
    preferred: [c1-02, c2-01, c4-02]
    forbidden: [uniform-card-grid, centered-feature-stack]
  cross_page_motif: ...
  page_rhythm:
    - P01: anchor
    - P02: breathing
    - P03: dense-evidence
    - P04: cinematic
```

Strategist 不只选择“dark-tech”或“professional”，而是要写清楚：

- 为什么这个方向适合受众；
- 什么视觉元素承担主记忆点；
- 页面如何形成节奏；
- 哪些常见页面结构必须禁止；
- 哪些页面允许极简，哪些页面必须有足够的证据密度。

### 6.3 方向候选必须是完整解决方案

每次需要用户选择方向时，至少提供 2 个真正不同的候选；每个方向都应包含完整的：

- visual thesis；
- page grammar；
- rendering behavior；
- composition vocabulary；
- image strategy；
- typography behavior；
- page archetype map；
- tradeoff。

仅仅改变蓝色、紫色和绿色不能算不同方向。

建议方向示例：

```text
A. Architectural Editorial
非对称建筑平面、纸张/石材层次、安静标题区、细线和大尺度空间。

B. Dark Cinematic System
深色画布、局部发光、轨道/节点/测量线、连续视觉状态。

C. Material Atlas
样本编号、切片、图谱、材质纹理、档案式注释和研究节奏。
```

### 6.4 Executor Prompt 增强

Executor 在绘制每页前必须完成以下内部判断：

1. 页面沟通任务是什么；
2. 当前 Composition Recipe 的主结构是什么；
3. 页面视觉锚点是什么；
4. 哪个元素拥有最大尺度、最高对比或最强材质；
5. 哪些元素必须保持安静；
6. 当前页如何延续或变化跨页 motif；
7. 哪些文本、数据和图表必须原生可编辑；
8. 是否存在默认卡片网格冲动；
9. 如果存在，什么更有意义的空间关系可以替代它。

建议加入硬性执行句：

```text
Do not begin with a component inventory.
Begin with the page-scale composition and the dominant visual relationship.
The first layout candidate must not be a uniform card grid unless equality,
comparison, or capacity is the page's actual meaning.
```

### 6.5 反保守机制

反保守不是要求每页都复杂，而是要求模型在使用安全结构前说明理由。

对于以下情况，Executor 必须返回上游重新判断：

- 连续两页使用相同卡片结构，但页面任务不同；
- 主视觉只是装饰，无法承担页面记忆点；
- 视觉方向没有改变页面的构图，只改变了颜色；
- 页面虽然留白很多，但没有明确的聚焦或节奏目的；
- 页面结构可以被轻易替换成普通标题页而不损失信息关系；
- 所有元素都拥有相近的视觉权重，没有主次。

## 7. 生成流程调整

### 7.1 一次性目标流程

完整能力上线后，交付级任务执行以下流程：

1. brief 和 acceptance contract；
2. domain paradigm；
3. 三个完整视觉方向；
4. 用户确认或采用推荐方向；
5. Theme Lock + Visual Grammar Lock；
6. page plan；
7. 每页 Composition Recipe；
8. 选定 Rendering Pack 和页面构图原型；
9. 生成 P01；
10. P01 方法门；
11. 连续生成剩余页面；
12. PPTX / structural checks，并检查所用 SVG 辅助元素；
13. PPTX → PDF → PNG；
14. Visual Effect Gate；
15. Defect Gate；
16. 方向性返工或局部返工；
17. 再次渲染和验收；
18. 交付 PPTX、源码、PDF、PNG 和记录。

### 7.2 P01 方法门

P01 不只是检查标题有没有溢出，还要验证：

- Rendering Pack 是否真的改变了页面气质；
- Composition Recipe 是否能落地；
- 主视觉是否足够强；
- 页面是否脱离了默认卡片逻辑；
- 字体层级、图像处理和装饰纪律是否成立；
- 当前方法能否扩展到正文页而不变成重复模板。

如果 P01 失败，应返回 Visual Grammar 或 Composition Recipe，而不是优先微调字号。

### 7.3 方向性返工和局部返工

| 问题 | 返回层级 |
|---|---|
| 概念不成立、视觉太保守、像通用模板 | Visual Direction / Rendering Pack |
| 页面关系不成立、主视觉弱、结构重复 | Composition Recipe / Page Plan |
| 图像裁切、字号、对齐和溢出 | Executor local edit |
| 主题没有实际应用、颜色回退 | Theme Lock / resolved theme |
| 内容过多导致布局失效 | Content Strategy / Page Plan |
| 模板品牌冲突 | VI Build context / protected merge |

## 8. 与现有 `pptx_designer` 的兼容方案

本方案不要求放弃 `pptx_designer`，而是重新定义它在系统中的位置：

### 8.1 Build Mode

继续使用 `Presentation(theme=resolved_theme, strict_theme=True)` 和公开 `pptx_designer.tools.*` API。

新增要求：

- Python 构图前必须有 Composition Recipe；
- 页面主体优先使用 PowerPoint 原生对象；复杂几何、局部材质、遮罩或特殊视觉层才使用 SVG 或 documented helper；
- 重要文本、数据、图表和需要编辑的结构保持原生；
- 视觉装饰和编辑性之间的选择必须记录；
- 页面不是由组件目录决定，而是由页面任务和视觉语法决定。

### 8.2 FreeStyle Mode

FreeStyle 可以继续使用 `generate_ppt()`，但必须传入完整 resolved theme，并把方向中的 visual grammar、page rhythm 和 forbidden patterns 作为结构化 content 或主题上下文的一部分。

FreeStyle 不应承担最复杂的页面构图。遇到 Composition Recipe 需要精确原生布局、复杂关系或局部 SVG 辅助时，路由器应显式返回 `NEEDS_BUILD_MODE`，由 workflow 重新生成 Build source；不能假设 `generate_ppt()` 自动拥有 Build Mode 的像素级控制。

### 8.3 VI Build Mode

VI Build 以 `skill/references/public-api.md` 当前确认的公开 API 为准，先统一 `extract_design_context()`、`extract_design_dna()` 和 `VIBuildDelivery` 的实际可用性，再保留 `merge_vi_design_context()` 的保护规则。未通过 API matrix 前，VI Build 升级项为 `BLOCKED`。新能力只影响模板允许自由设计的区域：

- 不覆盖模板锁定的品牌字体、Logo 和框架；
- 在模板内容页允许使用 Rendering Pack 和 Composition Recipe；
- 保留模板的 Master/Layout 和品牌 DNA；
- 视觉创新必须服从模板边界；
- 冲突仍然作为 preflight failure 处理。

### 8.4 SVG 与可编辑性边界

建议采用混合策略：

| 对象 | 默认实现 |
|---|---|
| 标题、正文、标签、来源、数字 | 原生文本 |
| 图表、数据表、关键流程节点 | 原生图表/形状或 documented helper |
| 几何装饰、材质平面、复杂背景 | SVG / supported vector |
| 复杂关系轮廓 | SVG 或原生 Boolean / freeform，按 fidelity 决定 |
| 重要图片 | 原生图片对象，保留比例和独立编辑能力 |
| 整页截图 | 禁止作为默认实现 |

“更炫”不能成为把整页内容烘焙成图片的理由。

## 9. 资产包目标规格与接入策略

为达到目标效果，最终完整资产包至少需要：

### Rendering Packs

- `editorial-luxe`；
- `dark-cinematic-tech`；
- `architectural-material`；
- `scientific-atlas`；
- `brutalist-data`；
- `cinematic-product`。

### Composition Families

- cover / opening；
- title + dominant visual；
- asymmetric split；
- image as canvas；
- evidence figure；
- process / sequence；
- comparison；
- data / chart；
- mosaic / gallery；
- section transition；
- closing / CTA；
- cross-page continuity。

### Case-derived Page Prototypes

每种 Rendering Pack 的长期目标是至少 6 条案例派生页面原型记录，即 6 个 Pack 合计至少 36 条高完成度 PPT-first 先例。首轮不要求重新制作全部 36 个页面，而是优先接入仓库已有高质量案例：案例目录继续保存唯一的 PPTX、PNG/PDF、页面 Recipe、object map 和局部 SVG；`case-prototype-index.json` 只登记 `case_id`、`slide_ids`、页面角色、Pack、summary 和这些文件的引用路径。一个案例可以派生多个页面原型记录，但同一页面不得另存一份“原型版”PPTX 或 PNG。缺失的 Pack、页面角色和构图关系再逐步补齐。

首轮 MVP 固定为 2 个 Rendering Pack、4 个 Composition Families、6～8 条案例派生原型记录，覆盖至少 4 种 Page Job，且每个首轮 Pack 至少覆盖 3 种页面角色。6 个 Pack、12 个 Family、36 条原型记录是长期覆盖目标，不是首轮阻塞条件。当前案例盘点必须先标出哪些案例缺少 PPTX、Recipe、object map 或许可信息；缺失证据的案例只能进入 `BLOCKED` 清单，不能计入合格原型数量。

### Prompt Assets

每个 Pack 需要提供：

- direction prompt fragment；
- strategist decision rules；
- executor composition rules；
- anti-pattern rules；
- review checklist；
- one worked example。

## 10. 工程实现建议

### 10.1 新增配置格式

建议使用 Markdown 作为人类可读规范，YAML / JSON 作为机器索引。

```yaml
pack_id: dark-cinematic-tech
version: 1
summary: precise nocturnal system for technical narratives
compatible_domains: [technical, product, architecture]
composition_families: [cover, evidence, process, data, continuity]
rendering:
  line: fine luminous strokes
  texture: subtle grid and controlled grain
  depth: glow and layered dark planes
  material: glass / graphite / light field
  mood: precise, nocturnal, engineered
composition_behavior:
  - localized accent only
  - one dominant diagram or visual field per page
anti_patterns:
  - purple gradient background
  - equal card grid without semantic equality
  - decorative circuit lines with no relationship to content
prototypes:
  - id: case-ai-agent-cover-01
    case_id: ai-agent-operating-system
    slide_ids: [P01]
    case_root: examples/new_examplex/<case-slug>/
    preview: <case_root>/renders/slide-01.png
    source: <case_root>/source/<case-slug>.pptx
    recipe: <case_root>/recipe/slide-01.md
    object_map: <case_root>/object-map/slide-01.md
    page_roles: [cover]
```

### 10.2 新增索引

运行时只读取索引，再按选择读取具体文件：

```json
{
  "dark-cinematic-tech": {
    "summary": "...",
    "best_for": ["technical", "product", "architecture"],
    "prototype_count": 6,
    "composition_families": ["cover", "evidence", "process", "data"],
    "prototype_ids": ["case-ai-agent-cover-01"]
  }
}
```

### 10.3 校验脚本

`validate_visual_pack.py` 应检查：

- pack metadata 是否完整；
- 所有 prototype 的 `case_id` 和 `slide_ids` 是否有效；
- preview 是否存在且可读取；
- 候选 source PPTX 是否可重开，并检查其与 build source、预览之间的 provenance 及可编辑对象；
- 存在 SVG fragments 时，才检查 SVG 是否可解析并符合安全约束；
- prototype 是否声明 page role 和 composition family；
- 是否存在 rendering、composition、anti-pattern 和 review sections；
- index 是否与实际文件一致。

`inspect_visual_assets.py` 应输出可供开发者快速复核的清单，不负责审美评分。

### 10.4 使用时自动刷新

案例和原型的一致性不应依赖开发者记忆。每次 Skill 启动、读取 `case-prototype-index.json` 之前，运行 `refresh_case_prototypes.py` 的轻量 preflight：

1. 读取案例与原型索引；
2. 扫描配置的案例根目录，比较已发现的案例与索引，报告新增、删除或未登记的 `case_id`；未完成契约的新增案例不得静默进入候选集；
3. 对每个被登记案例的 PPTX、预览、Recipe、object map 和 SVG 计算文件指纹（路径、大小、修改时间，必要时 SHA-256）；
4. 若指纹未变化，只确认 `case_id`、`slide_ids` 和路径仍可解析，不重复渲染；
5. 若发现变化，只对受影响案例执行 PPTX 重开检查、页面结构提取和必要的 PPTX→PDF→PNG 刷新；
6. 更新机器生成的 freshness / validation 状态（写入未纳入源码索引的 `.cache/case-prototype-state.json`，或在只读安装中保存在进程内），向运行时返回最新的 prototype records；
7. 若刷新失败，阻止该原型进入检索结果，并输出明确的案例路径、页面 ID 和失败原因。

该刷新是“按需增量”的，不是每次使用都完整渲染全部案例。源案例文件仍由开发者或构建流程修改；自动流程只更新派生预览、指纹和验证状态，不覆盖源 PPTX，也不偷偷改变 Recipe 或 object map。CI 仍应定期执行全量校验，防止长期未被使用的案例变陈旧。

## 11. 验收合同

### 11.1 功能验收

| ID | 条件 | 通过标准 |
|---|---|---|
| F1 | Visual Pack 可发现 | 索引可定位，路径和 schema 校验通过 |
| F2 | 方向候选完整 | 首轮至少提供 2 个真正不同的方向；每个方向均有视觉命题、语法、构图、图片、字体和反模式 |
| F3 | 页面 Recipe 可执行 | 每页拥有主结构、视觉锚点、占用区和编辑边界 |
| F4 | 案例派生原型可复用 | 索引可通过 `case_id` / `slide_ids` 定位真实案例页面，并明确哪些对象应原生实现、哪些局部可使用 SVG |
| F4a | 案例与原型单一事实源 | prototype record 不复制 PPTX、PNG、PDF 或 SVG；案例更新后可定位受影响的原型记录 |
| F5 | Build Mode 兼容 | 仍使用公开 `pptx_designer` API |
| F6 | 可编辑性保留 | 关键文本、数据和图表不是整页图片 |
| F7 | 可复现 | 相同 brief、seed、pack 和 resolved theme 产出可追踪结果 |

### 11.2 视觉验收

至少用 3 个不同领域的 deck 进行回归：技术系统、科学证据、品牌/建筑叙事。

每套 deck 必须检查：

- 首屏 3 秒内有明确主视觉和 first read；
- 至少两种不同的页面构图关系；
- 不出现连续无理由的默认卡片网格；
- 视觉方向不仅体现在配色，也体现在构图、图片处理和页面节奏；
- 页面密度随信息任务变化；
- 每页仍可在演示尺寸下阅读；
- 复杂 SVG 不遮蔽关键可编辑内容；
- 方向性失败会触发回到 Visual Grammar / Composition Recipe 的返工。

为使结果可比较，每页增加 6 项 0～2 分评分：

| 维度 | 0 分 | 1 分 | 2 分 |
|---|---|---|---|
| focal point | 无主视觉 | 有主视觉但竞争明显 | 首读明确且层级稳定 |
| composition distinction | 通用模板 | 有局部变化 | 空间关系明显区别于默认模板 |
| domain fit | 与领域冲突 | 基本适配 | 证据、语气和视觉结构高度匹配 |
| readability | 关键内容不可读 | 需放大或解释 | 演示尺寸下直接可读 |
| editability | 关键内容被烘焙 | 部分可编辑 | 文本、数据、图表和主要视觉可编辑 |
| cross-page rhythm | 页面互不相关 | 有弱连续性 | 节奏、motif 和结构变化清晰 |

总分满分 12 分，单页至少 9 分；`readability` 和 `editability` 不得低于 1 分；出现溢出、遮挡、错误事实、违反模板锁或整页截图替代关键内容时直接 `BLOCKED`。baseline 与 upgraded 必须使用相同 brief、内容、素材、模板、seed、包版本和页面数量；由一名主 Reviewer 评分，另一名 Reviewer 复核首屏、P01 和所有低于 2 分的项，保留逐页 evidence 和 revision record。

### 11.3 预期质量变化

在没有人工重新设计页面的情况下，预期达到：

| 指标 | 当前状态 | 目标状态 |
|---|---|---|
| 默认模板感 | 经常出现 | 明显减少 |
| 视觉方向与页面结构绑定 | 偏弱 | 强绑定 |
| 页面构图多样性 | 主要依赖 LLM 临场发挥 | 由可检索构图语法稳定提供 |
| 首屏视觉记忆点 | 不稳定 | 每页有明确 anchor，封面和章节页有强 hook |
| 复杂视觉复现门槛 | 依赖用户 Prompt 能力 | Skill 内部承担主要指导责任 |
| 复现和审查能力 | 有 Theme Lock | Theme + Grammar + Recipe + Prototype 全链路可追溯 |

这些是质量目标，不是未经测试的量化承诺。最终效果必须用回归 deck 的 PNG 和 acceptance contract 证明。

## 12. 可行性评估

### 12.1 技术可行性：中高（完成前置契约后）

方案可以建立在现有能力之上：

- `pptx_designer` 已支持主题、公开组件和 SVG 相关能力；
- 当前 skill 已有 Theme Lock、page plan、visual direction 和 PNG review；
- `pptx-master` 的资产组织方式可以作为设计方法参考，不需要复制其代码；
- Rendering Pack、Composition Catalog 和页面构图原型都是静态文档/资源，不需要先开发复杂服务；
- 新增索引和校验脚本可以用当前 Python 运行环境完成。

主要技术风险不在“能否生成”，而在“页面构图原型、公开 API、原生对象映射和 SVG 辅助边界是否被正确维护”。

### 12.2 设计可行性：中高

这是最需要真实验证的部分。写出目录和 Prompt 并不自动产生优秀作品，必须为每个 Rendering Pack 制作真正有质量的 PPT-first 页面构图原型和 worked example。若原型本身普通，模型只会更稳定地生成普通页面。

因此，设计资产质量是本方案的主要瓶颈，而不是代码量。

### 12.3 运行可行性：中高（MVP）；长期维护取决于自动刷新落地

运行时开销主要是：

- 多读取少量选定的 Markdown 规则；
- 读取当前方向对应的页面构图原型 summary，并在需要时查看 PNG/PDF 预览；
- 生成 Composition Recipe；
- P01 和最终质量检查。

通过索引和 selected-only loading，可以避免加载所有风格和构图目录。对普通 deck，额外上下文成本可控制在可接受范围内。

### 12.4 可维护性：中高

资产采用独立 Pack 和 Catalog 条目后，新增一个风格不会修改整个核心流程。每个包可以独立验证、独立回归和版本化。

需要严格维护：

- pack schema；
- case / prototype 命名及 `case_id`、`slide_ids` 关系；
- 可选 SVG fragments 的安全约束；
- public API compatibility；
- Theme Lock 与 Visual Grammar Lock 的字段边界；
- 资产索引和实际文件的一致性；
- 案例文件与原型记录不得出现重复副本。

### 12.5 一次性能力框架与渐进资产覆盖

“一次性到位”应指能力框架一次性完整落地，而不是要求第一轮就制作所有原型。首轮以已有案例作为真实资产种子，先让索引、检索、Recipe、Prompt 和回归验证闭环运行；原型覆盖率在后续按相同契约扩充。

本方案建议一次性完成：

- 目标架构；
- Prompt / role orchestration；
- schema 和索引；
- 首轮 2 个 Rendering Packs；
- 首轮 4 个 Composition Families；
- 已有案例接入形成 6～8 条 PPT-first 案例派生原型记录；
- 6 个 Pack、12 个 Family、36 条案例派生原型是长期完整覆盖目标，不是首轮阻塞条件；
- 校验脚本；
- 3 套回归 deck；
- 文档和验收记录。

这已经足够验证系统是否能明显改变视觉产出。未来增加原型属于资产覆盖扩展，不需要重新设计基础架构。

## 13. 主要风险与控制措施

| 风险 | 表现 | 控制措施 |
|---|---|---|
| 视觉资产本身普通 | Prompt 更复杂但结果仍普通 | 每个 Pack 必须有高完成度 PPT-first 页面样片和 worked example |
| 过度追求炫酷 | 阅读性下降、内容被装饰压过 | 每个构图模式必须绑定 communication job |
| 模板化升级为另一种模板化 | 所有页面重复某种拼贴 | Recipe 只提供结构词汇，Executor 必须根据 page job 变化 |
| SVG 破坏可编辑性 | 关键内容变成图片 | 明确 editable boundary，关键文字和数据保留原生 |
| 上下文过长 | LLM 遗忘当前页面规则 | index + selected-only loading |
| 领域错配 | 科研页使用品牌营销视觉 | domain paradigm 优先，Pack 只在兼容域中候选 |
| 视觉复核只修局部 | 源头方向错误仍保留 | 方向性问题必须返回 Grammar / Recipe 层 |
| 资产与实现漂移 | 文档说能做，API 实际不支持 | 原型和执行规则加入 public API 兼容校验 |
| 质量目标无法证明 | 只凭主观称“更漂亮” | 使用固定回归 deck、PNG evidence 和 acceptance matrix |

## 14. 最终交付物

一次性完成后，仓库应新增或更新：

```text
docs/
└── visual-design-capability-upgrade.md

skill/references/
├── visual-composition-catalog.md
├── visual-execution.md
├── visual-rendering-packs/
├── composition-recipes/
└── ...

skill/schemas/
├── rendering-pack.schema.json
├── composition-entry.schema.json
├── composition-recipe.schema.json
├── prototype-record.schema.json
├── runtime-trace.schema.json
└── acceptance-record.schema.json

skill/templates/
├── case-prototypes/
│   └── README.md
└── case-prototype-index.json

skill/scripts/
├── inspect_visual_assets.py
├── validate_visual_pack.py
└── refresh_case_prototypes.py

tests/
├── test_visual_assets.py
├── test_composition_recipe.py
└── fixtures/visual-packs/

examples/
└── visual-capability-regression/
```

每次正式生成还应在 task 目录记录：

- active rendering pack；
- selected composition entries；
- prototype ids；
- Visual Grammar Lock 版本；
- Theme Lock 版本；
- resolved theme fingerprint；
- seed；
- 生成器和包版本；
- P01 gate 结果；
- final PNG review 结果；
- 方向性修订记录。

## 15. 结论

本次升级的核心不是让 Skill“更会写 SVG”，也不是增加更多 PPT 组件，而是让 Skill 能够替普通用户完成原本需要设计师 Prompt Engineering 才能完成的判断：

```text
这页要让观众看到什么？
视觉重心是什么？
应该采用什么空间关系？
哪种材质和图像处理能承载这个主题？
哪些安全结构必须拒绝？
这个视觉系统如何跨页变化而不失去身份？
```

`ppt-design-skill` 已经有需求确认、主题锁定、可编辑生成和视觉验收的骨架。补上 Visual Grammar、Composition Catalog、Rendering Packs、PPT-first 页面构图原型和角色化执行 Prompt 后，它才能从“可靠的 PPT 生成流程”进一步成为“普通用户也能复现高完成度视觉作品的设计系统”。
