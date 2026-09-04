# PPT Design Skill 视觉设计能力升级实施计划

状态：实施计划

对应设计文档：[visual-design-capability-upgrade.md](visual-design-capability-upgrade.md)

更新时间：2026-09-04

## 1. 计划目的

本文件把《PPT Design Skill 视觉设计能力升级方案》转成可执行的工程计划，明确：

- 每个工作包要完成什么；
- 要新增或修改哪些文件；
- 工作包之间的依赖关系；
- 每个工作包的输入、输出和验收证据；
- 哪些问题必须返回上游重新设计；
- 如何证明最终能力真的改善了视觉产出。

本计划遵循“一次性到位”的能力框架目标：Schema、索引、Prompt 编排、运行时路由和验收机制在同一个升级范围内完成；原型资产首轮优先接入已有案例，后续按相同契约扩充。文档中的阶段只是执行顺序和质量门，不代表拆分成多个产品版本。

## 2. 目标交付范围

本次升级完成后，应同时具备以下能力：

### 2.1 首轮 MVP 与长期覆盖目标

“一次性到位”只要求能力契约一次性成立，不要求首轮制作全部长期资产。首轮完成的阻塞条件固定为：

- 2 个 Rendering Pack；
- 4 个 Composition Families；
- 6～8 条案例派生 Prototype Record，来自当前已有案例；
- 至少覆盖 4 种 Page Job，且每个首轮 Pack 至少覆盖 3 种页面角色；
- 角色编排、P01 gate、三种模式路由、正式 Schema、增量 preflight 和评分验收链路可运行；
- 现有案例中缺失的 PPTX、Recipe、object map、素材许可信息必须明确标记为 `BLOCKED`，不能伪装成合格原型。

长期非阻塞目标为 6 个 Pack、12 个 Family、36 条 Prototype Record；长期目标不参与首轮完成判定，但必须写入覆盖率报告。

1. Theme Tokens、Visual Grammar、Composition Prototypes 三层设计资产模型；
2. Rendering Pack 的规范、索引、加载和验证机制；
3. Composition Catalog 与 Composition Recipe；
4. 案例派生的 PPT-first 页面构图原型记录；
5. 选定方向后按需加载 Prompt / role instructions；
6. Strategist 生成完整视觉方向；
7. Executor 根据 Recipe 进行 PPT-first 构图；
8. SVG 仅作为局部复杂视觉的辅助路径；
9. P01 方法门和方向性返工机制；
10. Build / FreeStyle / VI Build 的兼容策略；
11. 视觉资产校验与索引一致性检查；
12. 三套跨领域回归 deck；
13. 可追溯的 Theme、Grammar、Recipe、Prototype 和 Review 记录。

## 3. 不变的架构边界

### 3.1 PPT-first

PowerPoint 是主要设计和交付工具。原生对象承担：

- 标题、正文、标签、脚注和来源；
- 图表、数据表、关键流程节点和可编辑数据；
- 主要页面结构、分栏、面板、图像框架和必要的几何形状；
- 用户后续需要修改的内容。

### 3.2 SVG-assisted

SVG 只用于局部效果，且必须有明确的视觉工作：

- 原生形状难以稳定表达的复杂轮廓；
- 特殊裁切、遮罩或注册关系；
- 局部材质纹理或复杂装饰；
- 需要保持视觉质量的特殊矢量片段。

SVG 不得：

- 取代整页 PPT 作为主要页面源；
- 包含本应可编辑的标题、正文、数据或图表；
- 作为绕过 PowerPoint 原生对象能力的默认方案；
- 让原型和生成结果只能以图片形式编辑。

### 3.3 公开 API

实现继续使用公开的 `pptx_designer` API，不恢复或复制旧的 `ppt_pro_max` 实现。所有新增调用必须在 `skill/references/public-api.md` 或对应公开文档中有依据。

### 3.4 现有质量门继续有效

新增视觉能力不能削弱现有规则：

- brief 和 acceptance contract 仍然是入口；
- Theme Lock 仍然是主题事实源；
- VI Build 的模板锁定字段仍然优先；
- PPTX → PDF → PNG 仍然是正式交付路径；
- PNG 视觉检查仍然是最终质量证据；
- 真实数据、引用、可编辑性和字体可用性仍然必须检查。

## 4. 总体依赖图

```text
WP0 基线与资产盘点
        ↓
WP1 规范、Schema 与索引
        ↓
WP2 Visual Rendering Packs
        ↓
WP3 Composition Catalog
        ↓
WP4 案例接入与原型登记
        ↓
WP5 Prompt / Role Orchestration
        ↓
WP6 Build / FreeStyle / VI Build 接入
        ↓
WP7 校验脚本和运行时证据
        ↓
WP8 回归 deck 与对比评估
        ↓
WP9 文档、安装和示例同步
        ↓
WP10 最终审查与交付
```

WP2、WP3 和 WP4 是设计资产链；WP5 和 WP6 是运行链；WP7、WP8、WP9 和 WP10 是验证和交付链。不能跳过资产链直接修改 Prompt，也不能只通过代码检查而不做回归 deck 的 PNG 对比。WP4 不再创建独立的“原型资产包”，而是把已有案例登记为可检索原型；案例文件和原型记录必须保持一对多关系。

## 5. 总体实施原则

### 5.1 一次性完成，不做虚假的中间版本

所有目标字段、资产格式、运行入口和验收证据在本次变更中统一定义。禁止出现以下状态：

- 文档声称有 Rendering Pack，但没有索引和实际文件；
- Prompt 要求 Composition Recipe，但执行器仍然完全自由构图；
- 原型只提供图片，但没有 PPT-first 的对象边界；
- Build Mode 接入了新字段，FreeStyle 和 VI Build 却出现隐式冲突；
- 只有“看起来更炫”的单页，缺少跨领域回归证据。

### 5.2 先锁契约，再制作资产

先确定 schema、文件命名、索引和验证规则，再制作 Pack 和原型。不能先批量制作未定义格式的素材，最后再倒推运行时如何读取。

### 5.3 先页面工作，再装饰

任何构图模式都必须绑定 page job。任何材质、光影、纹理和装饰都必须绑定视觉工作。没有沟通工作的装饰不进入正式资产包。

### 5.4 先 PPT 方案，再决定 SVG 辅助

制作原型时先完成 PowerPoint 原生对象设计，再判断局部是否需要 SVG。不能因为 SVG 更容易画，就把整页设计转成 SVG。

### 5.5 质量评价必须基于证据

“更高级”“更炫”“更有设计感”不能单独作为验收结论。必须记录：

- 页面结构发生了什么改变；
- 主视觉是什么；
- 哪些规则阻止了模板化结果；
- PNG 中看到了什么证据；
- 哪些对象仍保持可编辑。

### 5.6 MVP 资源、负责人和降级规则

以下是用于排期的粗略估算，不是工时承诺：

| 工作 | 主要 owner | MVP 估算 | 可并行性 |
|---|---|---:|---|
| Schema、索引、路径和状态机 | 工程 | 8～16 小时 | 可与案例盘点并行 |
| 6 个案例证据盘点与补齐 | 设计/内容 | 12～24 小时 | 可按案例并行 |
| 2 Pack、4 Family 与 Prompt 接入 | 设计 + Prompt | 12～24 小时 | Pack 可并行，需共享 Schema |
| 三模式路由、P01、校验和缓存测试 | 工程 | 12～20 小时 | 依赖 Schema 和路由契约 |
| 三套回归 deck 与双人评分 | 设计 QA | 12～20 小时 | 依赖 MVP 资产闭环 |

首轮 MVP 预计为 56～104 个有效工时；36 条长期原型记录、6 Pack 和 12 Family 另行排期。每个工作包必须指定一名 owner、一名 reviewer 和可交付证据。若案例缺少合法素材或公开 API 未统一，降级为 `BLOCKED` 并减少覆盖数量，不得用占位文件冒充通过；若渲染刷新超时，交付级 Build 等待并失败，快速 FreeStyle 只允许无 prototype 降级并记录 trace。

## 6. WP0：基线与现有资产盘点

### 6.1 目标

在任何实现前建立现有仓库的基线，避免重复建设、破坏已有接口或覆盖用户已有改动。

### 6.2 检查范围

检查以下文件和目录：

- `skill/SKILL.md`；
- `skill/references/workflow.md`；
- `skill/references/design-direction-coach.md`；
- `skill/references/design-principles.md`；
- `skill/references/public-api.md`；
- `skill/references/template-brand.md`；
- `skill/references/qa-and-delivery.md`；
- `skill/templates/task-init/`；
- `skill/scripts/`；
- `examples/new_examplex/`；
- `tests/`；
- `skill.json`；
- `README.md` 和 `docs/README*.md`。

### 6.3 输出

新增或更新一份内部盘点记录，至少记录：

```text
现有 Theme Lock 字段：...
现有 visual direction 字段：...
现有 Build Mode 入口：...
现有 FreeStyle 入口：...
现有 VI Build 入口：...
现有 SVG 能力：...
现有 PNG review 能力：...
不能破坏的 public API：...
需要新增的字段：...
```

VI Build 兼容性必须单独形成 API matrix：当前公共文档同时出现 `extract_design_dna()`、`extract_design_context()`、`merge_vi_design_context()` 和 `VIBuildDelivery`。在升级进入实现前，以 `skill/references/public-api.md` 和已安装 `pptx_designer` 的实际导出为准，明确：

- `extract_design_context()` 是否是 VI Build 的唯一公开模板上下文入口；
- `extract_design_dna()` 是公开兼容别名、底层分析函数，还是应从主流程移除；
- `merge_vi_design_context()` 的输入输出结构和冲突记录方式；
- `VIBuildDelivery` 是否真实存在于当前公开包，若存在则补入 public API，否则 README 和计划不得继续引用。

在 matrix 通过前，VI Build 属于 `BLOCKED`，不能用新增 Prompt 掩盖 API 不一致。

盘点记录可以并入实施日志，不要求成为最终用户文档。

当前基线（2026-09-04）：`examples/new_examplex/` 下已发现 6 个案例目录：`vertical_city_retrofit`、`car_t_single_cell_paper`、`louvre_abudhabi`、`ai_infrastructure_economics`、`couture_lipstick_atelier`、`ai_agent_operating_system`。当前扫描发现每个案例都有位于 `output/` 或 `rendered/` 的候选 PPTX/PDF，但尚未确认其与 `build.py` 的 provenance；同时未发现可统一命名的 `object-map` / `recipe` 文件。因此这些案例不能直接计入合格 Prototype，必须先完成 provenance、证据补齐或明确标记为 `BLOCKED`。该基线只描述当前状态，实施时应重新扫描确认。

### 6.4 验收

- 没有重复定义 `Theme Lock` 已有字段；
- 已明确新字段属于 Theme、Grammar、Recipe 还是 Prototype；
- 已明确哪些内容只写文档，哪些内容需要脚本；
- 已检查工作区现有删除和未提交状态，没有覆盖用户改动；
- `git diff --check` 通过。

## 7. WP1：规范、Schema 与索引

### 7.1 目标

先建立可读、可索引、可验证的统一数据契约。

### 7.2 新增规范文件

新增：

```text
skill/references/visual-composition-catalog.md
skill/references/visual-execution.md
skill/references/visual-rendering-packs/README.md
skill/references/composition-recipes/README.md
skill/templates/case-prototypes/README.md
```

其中：

- `visual-composition-catalog.md` 定义构图条目、ID、页面任务、兼容关系和反模式；
- `visual-execution.md` 定义 Executor 的 PPT-first 执行顺序和 SVG 边界；
- Rendering Pack README 定义风格包字段和加载规则；
- Composition Recipe README 定义每页 Recipe 字段；
- Case Prototype README 定义案例目录、原型映射记录和证据要求，不定义第二套 PPTX/PNG 资产目录。

### 7.3 统一字段

#### Rendering Pack

```yaml
pack_id: string
version: integer
summary: string
best_for: [string]
compatible_domains: [string]
visual_thesis: string
composition_behavior: [string]
rendering:
  line: string
  texture: string
  depth: string
  material: string
  mood: string
typography_behavior: [string]
image_treatment: [string]
composition_families: [string]
anti_patterns: [string]
prompt_fragment: string
review_focus: [string]
```

#### Composition Entry

```yaml
composition_id: string
name: string
page_jobs: [string]
content_textures: [string]
geometry_recipe: [string]
compatible_modifiers: [string]
compatible_domains: [string]
editable_boundary:
  native: [string]
  svg_assisted: [string]
  forbidden_as_svg: [string]
failure_modes: [string]
prototype_ids: [string]
```

#### Composition Recipe

```yaml
page_id: string
page_job: string
primary_structure: string
modifiers: [string]
visual_anchor: string
occupied_zones: [object]
content_texture: string
motif_behavior: string
native_object_plan: [object]
svg_assistance: [object]
avoid: [string]
review_questions: [string]
```

#### Case-derived Prototype Record

```text
case-prototype-record/
├── prototype.yaml        # metadata view, not a copied asset bundle
└── case_ref:
    case_id: string
    case_root: string
    slide_ids: [string]
    preview_paths: [string]
    source_paths: [string]
    recipe_paths: [string]
    object_map_paths: [string]
```

案例目录是唯一事实源：`source_paths` 指向案例中的 PPTX，`preview_paths` 指向案例中的 PNG/PDF，`recipe_paths` 和 `object_map_paths` 指向案例随附的解释文件。Prototype record 不复制这些文件；`svg-fragments` 如果存在，也必须留在案例目录中并通过路径引用。

关系约束：一个 `case_id` 可以派生多个 `prototype_id`（例如同一 deck 的不同页面或页面组），但每个 prototype record 必须用 `slide_ids` 明确范围；原型记录删除或重建不应影响案例本体。案例更新后，受影响的 prototype record 必须重新审查其页面截图、Recipe、object map 和索引字段。

### 7.4 索引

新增：

```text
skill/references/visual-rendering-packs/rendering_packs_index.json
skill/references/composition-recipes/composition_index.json
skill/templates/case-prototype-index.json
```

索引只保存发现和路由所需的轻量字段，不复制完整 Prompt，不嵌入整套 SVG 或二进制文件。

### 7.5 验收

- 所有字段有明确的 owner；
- Theme Lock 不与 Visual Grammar Lock 重复保存同一事实；
- 索引可以定位每个 Pack、Composition 和 Prototype；
- Prototype record 指向案例中的 PPTX/PNG/PDF，不复制独立的原型文件；
- 所有路径规则、命名规则和版本字段已经写入 README；
- schema 示例能被后续校验脚本解析。

### 7.6 正式 Schema 契约

YAML 只用于人类编写；交付前必须生成并提交 Draft 2020-12 JSON Schema，不能把下面的示例当作 Schema 本身。目标文件为：

```text
skill/schemas/rendering-pack.schema.json
skill/schemas/composition-entry.schema.json
skill/schemas/composition-recipe.schema.json
skill/schemas/prototype-record.schema.json
skill/schemas/runtime-trace.schema.json
skill/schemas/acceptance-record.schema.json
```

最低约束如下：

| 对象 | 必填字段 | 关键约束 |
|---|---|---|
| Rendering Pack | `pack_id`, `version`, `summary`, `compatible_domains`, `composition_families`, `rendering`, `anti_patterns` | ID 使用 kebab-case；`version` 为正整数；Family 引用必须存在 |
| Composition Entry | `composition_id`, `page_jobs`, `structure`, `compatible_modifiers`, `editable_boundary`, `prototype_ids` | Page Job、modifier、prototype 引用必须可解析；原生对象与 SVG 禁用对象不得交集 |
| Composition Recipe | `recipe_id`, `page_id`, `page_job`, `primary_structure`, `visual_anchor`, `occupied_zones`, `native_object_plan`, `svg_assistance`, `review_questions` | 坐标归一化到 0～1；zone 不得越界；`svg_assistance` 必须声明 purpose 和 allowed boundary |
| Prototype Record | `prototype_id`, `case_id`, `slide_ids`, `pack_id`, `composition_ids`, `page_roles`, `summary`, `source_paths`, `preview_paths`, `recipe_paths`, `object_map_paths`, `license_status`, `context_allowed`, `redaction_status` | ID 唯一；路径必须为相对路径且禁止 `..`；所有引用存在；许可、上下文权限和脱敏状态必须明确；不得包含二进制副本 |
| Runtime Trace | `run_id`, `mode`, `prototype_ids`, `recipe_ids`, `seed`, `package_version`, `p01_gate`, `final_visual_gate` | gate 只能为 `PASS` / `NEEDS_REVISION` / `BLOCKED`；正式生成必须可追溯 |
| Acceptance Record | `record_id`, `target_id`, `criteria`, `score`, `reviewer`, `evidence`, `status` | MVP 视觉记录固定 6 项 criteria；score 为逐项 0～2 分且必须等于各项之和；每项必须有 evidence；状态与阈值一致 |

所有 Schema 还必须声明 `additionalProperties: false`（允许显式 `extensions`），并使用 `$defs` 共享 ID、relative-path、gate-status、score 等定义。校验器负责结构和引用一致性；语义质量由 acceptance record 和 PNG/PPTX 证据负责，不能声称 JSON Schema 能自动判断审美。

### 7.7 字段 owner 与继承优先级

字段只能有一个事实 owner；其他层级只能引用或在允许范围内覆盖：

```text
VI template locks
  > Theme Lock
  > Domain Paradigm
  > Visual Grammar Lock / Rendering Pack
  > Page Recipe
  > Executor local decision
```

- VI template locks 只能在 VI Build 中出现，控制母版、Logo、字体、边距和品牌禁区；
- Theme Lock 控制项目级视觉命题、色彩意图、字体意图、密度和整体节奏；
- Domain Paradigm 控制证据表达、术语、图表习惯和受众适配；
- Rendering Pack / Visual Grammar 控制材质、构图倾向、图片行为、反模式和跨页 motif；
- Page Recipe 只能决定当前页面的 Page Job、Structure、Anchor、zone 和对象计划；
- Executor 只能在上述边界内调整坐标、字号、裁切和局部实现。

冲突不得静默覆盖：记录 `conflict_type`、双方字段、采用的上层规则和人工/Reviewer 决策。Template lock 与可编辑性安全规则是硬阻塞项；Recipe 与 Pack 冲突时，只有在 Recipe 明确声明 `override_reason` 且不违反 Theme/Template lock 时才可通过。

## 8. WP2：Visual Rendering Pack

### 8.1 目标

把“风格名称”升级成可执行的视觉指导包，控制线条、材质、深度、图像处理、字体行为、构图倾向和反模式。

### 8.2 一次性制作的 Pack

首轮只制作 2 个 Pack；长期目标扩展到以下 6 个 Pack：

1. `editorial-luxe`；
2. `dark-cinematic-tech`；
3. `architectural-material`；
4. `scientific-atlas`；
5. `brutalist-data`；
6. `cinematic-product`。

### 8.3 每个 Pack 的制作步骤

对每个 Pack 逐项完成：

1. 明确适用领域和不适用领域；
2. 写一条不可替换为普通风格标签的 visual thesis；
3. 写 composition behavior；
4. 写材质、线条、纹理、深度和 mood；
5. 写字体比例、字重和标签行为；
6. 写图片构图、裁切和色调规则；
7. 绑定可用 Composition Families；
8. 写至少 5 条 anti-pattern；
9. 写 Prompt Fragment；
10. 写 review focus；
11. 添加至少一个 worked example；
12. 写入 index；
13. 运行 Pack 校验；
14. 用一个实际页面验证它是否改变了构图，而不只是改变颜色。

### 8.4 Pack 的质量要求

每个 Pack 必须能回答：

- 它要让观众感受到什么；
- 它通过什么页面结构实现这种感受；
- 它如何处理主视觉和内容的关系；
- 它怎样使用留白、尺度、密度和跨页节奏；
- 哪些结构它坚决不使用；
- 什么时候应该回到普通原生 PPT 结构；
- 什么时候可以使用局部 SVG。

### 8.5 验收

- 首轮 2 个 Pack 文件均存在并通过完整字段校验；长期验收再要求 6 个 Pack 文件均存在；
- index 与文件一致；
- 每个 Pack 包含完整字段；
- 每个 Pack 至少有一个 worked example；
- anti-pattern 不是泛泛的“不要太普通”，而是可观察结构；
- review focus 能在 PNG 中找到对应证据；
- Pack 不强迫科研、医疗等领域使用不合适的营销视觉。

## 9. WP3：Composition Catalog

### 9.1 目标

建立从 page job 到页面空间关系的中间层，让 LLM 在构图时有可执行的选择，而不是从空白画布随机发挥。

### 9.2 首轮与长期 Composition Families

首轮从以下目录中选择 4 个 Family；长期目标覆盖全部 12 类：

1. cover / opening（Page Job）；
2. title + dominant visual（Page Job）；
3. asymmetric split（Composition Structure）；
4. image as canvas（Composition Structure）；
5. evidence figure（Page Job）；
6. process / sequence（Page Job）；
7. comparison（Page Job）；
8. data / chart（Page Job）；
9. mosaic / gallery（Composition Structure）；
10. section transition（Page Job）；
11. closing / CTA（Page Job）；
12. cross-page continuity（Cross-page Behavior）。

实现时必须把混合清单拆成三层字段：`page_job`（沟通任务）、`structure`（空间结构）和 `modifiers` / `cross_page_behavior`（修饰与跨页行为）。检索先按 Page Job，再组合 Structure 和 Modifier；不能把 12 个条目当作同一层级的互斥模板。

### 9.3 每个构图条目必须包含

- 唯一 ID；
- 页面沟通任务；
- 适合的内容纹理；
- 主空间关系；
- 推荐尺度关系；
- 推荐 occupied zones；
- 可组合的 modifier；
- 可使用的 Rendering Pack；
- 原生对象实现建议；
- 可选 SVG 辅助边界；
- 失败模式；
- 至少一个 prototype id。

### 9.4 构图选择规则

Executor 选择构图时按以下顺序：

```text
页面任务
→ 主信息关系
→ 主构图结构
→ 视觉锚点
→ 图片/材质/深度处理
→ 原生对象实现
→ 必要的局部 SVG 辅助
```

必须禁止以下倒序：

```text
先列组件
→ 先放卡片
→ 再把内容填进去
→ 最后加装饰
```

### 9.5 验收

- 12 个 Family 都能绑定真实 page job；
- 至少有 3 个 Family 明确拒绝均匀卡片网格；
- 至少有 3 个 Family 明确跨页或视觉连续性；
- 每个 Family 都给出 PPT 原生实现建议；
- SVG 只作为可选局部辅助；
- 目录可以被 Recipe 和 Prompt 引用；
- 条目之间没有同义重复或仅换颜色的伪差异。

## 10. WP4：案例接入与原型登记

### 10.1 目标

为 LLM 提供真实的高完成度页面参考，让它学习空间关系、尺度、密度、材质和可编辑边界，而不是只阅读文字规则；同时保证案例和原型只有一套可维护的文件。

### 10.2 数量与首轮来源

长期目标是每个 Rendering Pack 至少 6 条案例派生原型记录，共至少 36 条记录。首轮优先接入 `examples/new_examplex/` 和其他现有高质量案例；只要案例已有或补齐 PPTX、PNG/PDF、Recipe、object map，并完成索引登记，就可以作为第一批有效原型，不要求首轮重新制作 36 个页面，也不为同一页面另存一份 prototype PPTX/PNG。

长期目标下每个 Pack 至少覆盖：

- cover / opening；
- section transition；
- evidence / explanation；
- comparison / sequence；
- data / chart；
- ending / CTA。

### 10.3 案例接入与登记要求

每个接入案例的原型记录必须：

1. 用 `case_id` 和 `slide_ids` 指向唯一案例及其页面范围；
2. 引用案例已有的 PPTX、PNG/PDF、Recipe 和 object map 路径；
3. 标明主要原生对象及局部 SVG 辅助对象，如有；
4. 使用真实但不具误导性的示例内容；
5. 证明页面不是把标题、正文、图表或数据烘焙进整页图片；
6. 记录案例的依赖字体、外部图片和安装/渲染前提；
7. 为每个外部图片、字体、SVG 和案例素材记录来源、许可证、是否可随 Skill 分发、是否含客户专属/敏感内容以及是否允许进入模型上下文；
8. 能够说明它传递的是构图方法，不是要求 1:1 复制的固定版式；
9. 不在 prototype record 目录中复制 PPTX、PNG、PDF 或 SVG 二进制文件。

如果现有案例缺少 Recipe 或 object map，应补充到案例目录本身，再登记原型；不得为了“快速接入”在原型目录另建一套旁路解释文件。

### 10.4 原型 object map

示例：

```markdown
| Object | Implementation | Editable | Reason |
|---|---|---:|---|
| Main title | native text | yes | user content changes |
| Hero image | native picture | yes | crop may change |
| Architectural contour | native freeform / optional SVG fragment | partial | complex local shape |
| Metric value | native text/chart | yes | factual data must remain editable |
| Background texture | optional SVG fragment | no | decorative only |
```

### 10.5 原型审查

每个案例派生原型必须经过：

- PPTX 重开检查；
- 预览 PNG 检查；
- 文字和对象可编辑性检查；
- 图片裁切和比例检查；
- 案例目录中的 `recipe.md` 与实际页面对照；
- 案例目录中的 `object-map.md` 与实际对象对照；
- 是否出现模板化卡片堆叠检查。

### 10.6 验收

- 首轮已有案例形成 6～8 条案例派生原型记录，并覆盖 2 个首轮 Pack、4 个 Family、至少 4 种 Page Job；
- 每个首轮 Pack 至少覆盖 3 个页面角色；长期验收再要求每个 Pack 的 6 个页面角色齐全；
- 原型中至少一部分使用非对称、尺度对比、跨页连续或材质层次；
- 原型可以映射到原生 PPT 对象；
- SVG 只出现在明确记录的局部；
- 每个原型记录都有可供模型读取的 summary，且 `case_id`、`slide_ids` 和所有引用路径可解析；
- 每个案例都有素材许可和脱敏状态；不允许分发或进入上下文的素材被替换、裁切或标记为不可用；
- 原型预览在演示尺寸下仍可辨识主视觉和主信息。

## 11. WP5：Prompt / Role Orchestration

### 11.1 目标

让普通用户不需要自己掌握深度设计 Prompt，Skill 通过角色和按需加载自动承担设计判断。

### 11.2 运行时落点与状态机

本轮不假设存在一个新的多代理服务，但必须提供一个可运行的 workflow runner（建议为 `skill/scripts/run_visual_workflow.py`）。`SKILL.md` 和 workflow reference 负责角色决策规则；runner 负责状态持久化、路由、暂停/恢复、P01 gate 调用和失败回退；其他 Python 工具负责 Schema 校验、PPTX inspection、渲染、缓存和证据生成。若未来引入专用 orchestrator，必须保持相同的输入输出契约。

```text
brief
  → Router：选择 Build / FreeStyle / VI Build
  → Strategist：生成 direction candidates
  → user confirmation（或 one-shot 明确跳过）
  → Theme Composer：写入 Theme Lock / resolved theme
  → Composition Planner：为每页选择 Page Job + Structure + Recipe + Prototype refs
  → P01 Gate：先生成并检查首个代表页
      PASS → Executor 生成完整 deck
      NEEDS_REVISION → 回到 Planner 或 Strategist
      BLOCKED → 暂停并请求用户/素材/模板决策
  → Reviewer：PNG、结构、编辑性和 acceptance record
      NEEDS_REVISION → 返回对应决策层
      PASS → 交付
```

每一步都必须写入任务目录的机器可读中间状态，建议最小结构为：

```text
ppt_tasks/<task-id>/design/
├── brief.json
├── direction-candidates.json
├── theme-lock.json
├── page-plan.json
├── prototype-selections.json
├── runtime-trace.json
└── acceptance-record.json
```

文件写入采用临时文件加原子替换；状态至少包含 `status`、`attempt`、`updated_at`、`parent_artifact` 和 `failure_reason`。用户确认后从 `theme-lock.json` 恢复；中断后优先读取最后一个 `PASS` 状态，不重复执行已完成的上游步骤。

模式路由必须有明确落点：有模板或品牌合规要求进入 VI Build；需要精确坐标和交付级质量进入 Build；只需要快速目标驱动草稿才进入 FreeStyle。FreeStyle 不会隐式变成 Build：如果 P01 或用户需求要求精确构图，路由器必须显式返回 `NEEDS_BUILD_MODE`，由 workflow 重新生成 Build source，而不是假设 `generate_ppt()` 提供像素级控制。

### 11.3 角色边界

#### Strategist

负责：

- audience outcome；
- communication intent；
- visual thesis；
- Rendering Pack 候选；
- Visual Grammar；
- 三套完整方向；
- page rhythm；
- Composition Family 候选；
- 方向风险和反模式。

不负责：

- 逐个绘制 PPT 元素；
- 在没有内容证据时发明数据；
- 直接决定每个装饰对象的坐标。

#### Composition Planner

负责：

- 每页 page job；
- primary structure；
- modifiers；
- visual anchor；
- occupied zones；
- native object plan；
- optional SVG assistance；
- page-to-page motif behavior。

#### Executor

负责：

- 根据 Recipe 进行 PPT-first 构图；
- 原生文本、形状、图片、图表和图示；
- 必要的局部 SVG 辅助；
- 尺度、层级、间距、裁切和对齐；
- 记录实现和可编辑边界。

不负责：

- 重新选择已锁定的视觉方向；
- 重新解释模板锁定字段；
- 用整页截图绕过原生对象；
- 使用未准备或未声明的素材。

#### Reviewer

负责：

- 对照 acceptance contract；
- 判断视觉方向是否真正改变构图；
- 判断页面是否普通模板化；
- 判断是局部问题还是方向性问题；
- 生成可追踪的 revision record。

### 11.4 按需加载顺序

```text
Global workflow
→ case-prototype preflight（读取指纹；仅对变化案例增量刷新）
→ domain paradigm
→ selected Rendering Pack
→ selected Composition entries
→ prototype summaries and previews when needed
→ current page Recipe
→ Executor instructions
```

禁止默认加载全部 Pack、全部构图条目和全部原型内容。索引负责发现，正文负责执行，预览负责必要时的视觉核对。Preflight 每次 Skill 使用自动执行，但未发现指纹变化时只做轻量路径和索引检查。

### 11.5 Strategist 输出合同

必须新增或扩展：

```yaml
visual_direction:
  visual_thesis: string
  audience_promise: string
  rendering_pack: string
  visual_grammar:
    scale_contrast: low|medium|high
    asymmetry: none|optional|intentional
    material_depth: string
    image_behavior: string
    decoration_behavior: string
  composition_vocabulary:
    preferred: [string]
    forbidden: [string]
  cross_page_motif: string
  page_rhythm: [object]
```

### 11.6 方向性返工合同

以下情况必须返回 Strategist 或 Composition Planner：

- P01 仍然像通用模板；
- Rendering Pack 只改变颜色，没有改变构图；
- visual anchor 不明确；
- 连续页面没有叙事上的结构变化；
- 所有元素视觉权重相近；
- 主视觉不能承担页面记忆点；
- 页面留白没有明确功能；
- 原型方法无法扩展到当前内容。

### 11.7 验收

- 普通用户 brief 不需要自行提供长 Prompt；
- Strategist 输出三套真正不同的完整方向；
- 每套方向包含 Rendering、Grammar、Composition、Image、Typography 和 Anti-pattern；
- Executor 能拿到当前页完整 Recipe；
- 未选定的 Pack 和 Prototype 不进入主要上下文；
- 方向性失败不会被错误地当成字号或对齐问题处理。

## 12. WP6：生成模式接入

### 12.1 Build Mode

修改 `skill/SKILL.md`、`skill/references/workflow.md` 和 `skill/references/design-direction-coach.md`：

- 明确 Build Mode 需要 Composition Recipe；
- 明确 PPT 原生对象是默认实现；
- 明确 SVG 只作为局部辅助；
- 规定 P01 方法门；
- 规定方向性失败的返回层级；
- 记录 selected Pack、Composition IDs 和 Prototype IDs。

Build Mode 的生成脚本仍然使用公开 `pptx_designer` API，并继续记录 resolved theme、seed、package version 和 module path。

### 12.2 FreeStyle Mode

FreeStyle 继续使用 `generate_ppt()`，但需要：

- 传入完整 resolved theme；
- 将 visual grammar、page rhythm 和 forbidden patterns 作为结构化上下文；
- 在复杂 Recipe 需要精确布局时自动转入 Build Mode；
- 不因追求视觉效果而把整页内容转成图片；
- 保留 PNG review gate。

### 12.3 VI Build Mode

VI Build 继续使用 `merge_vi_design_context()`，并增加：

- Rendering Pack 只能作用于模板允许变化的区域；
- Composition Recipe 必须尊重 Master/Layout 和品牌锁定字段；
- Prototype 必须记录哪些元素继承模板、哪些元素为新增内容；
- 发生模板冲突时停止，不静默覆盖模板事实。

### 12.4 可编辑性决策表

| 对象 | 默认实现 | SVG 条件 |
|---|---|---|
| Title / body / label | native text | 禁止整页 SVG 文本替代 |
| Chart / data table | native chart/table | 仅在装饰性背景局部使用 |
| Process / architecture node | native shapes / documented helpers | 复杂轮廓可局部辅助 |
| Hero image | native picture | SVG 只用于特殊局部裁切 |
| Decorative geometry | native shape / freeform | 原生无法稳定表达时使用 |
| Material texture | native fill / prepared asset | 局部 SVG 可选 |
| Full page | native composition | 禁止整页 SVG 作为默认源 |

### 12.5 验收

- 三种模式的字段命名一致；
- Build Mode 仍然可精确控制坐标；
- FreeStyle 不承诺超出其控制能力的复杂构图；
- VI Build 不破坏模板锁定字段；
- 关键内容始终保持独立可编辑；
- 旧的 `ppt_pro_max` 不被重新引入。

## 13. WP7：校验脚本与运行时证据

### 13.1 新增脚本

```text
skill/scripts/validate_visual_pack.py
skill/scripts/inspect_visual_assets.py
skill/scripts/refresh_case_prototypes.py
skill/scripts/run_visual_workflow.py
```

### 13.2 `validate_visual_pack.py`

必须检查：

- Pack metadata；
- Composition entry metadata；
- Prototype record 的 `case_id` 和 `slide_ids`；
- `case_id` 能在案例索引中解析；
- `slide_ids` 能解析到案例中的实际页面；
- `preview_paths`、`source_paths`、`recipe_paths` 和 `object_map_paths` 均存在且位于案例目录内；
- 案例中的 PPTX 可重开、预览可读取、Recipe 和 object map 可读取；
- prototype record 没有复制 PPTX、PNG、PDF 或 SVG 二进制文件；
- SVG fragments 存在时才校验 SVG；
- index 与实际文件一致；
- 版本字段和 ID 唯一；
- 兼容领域和 Composition Family 引用有效。

### 13.3 `inspect_visual_assets.py`

输出：

- Pack 数量；
- Composition 数量；
- Prototype 数量；
- 缺失路径；
- 未引用条目；
- 无 preview 的原型；
- 无可重开 PPTX 或未确认 provenance 的原型；
- 可选 SVG fragments 数量；
- 各 Pack 的页面角色覆盖；
- 各 Pack 的 worked example 状态。

该脚本不负责审美评分，只负责资产完整性和发现问题。

### 13.4 `refresh_case_prototypes.py`

这是每次 Skill 使用时自动调用的增量刷新入口，必须具备幂等性：

状态模型：

```text
indexed → checking → valid
                  ↘ stale → refreshing → valid
                                    ↘ invalid
```

`checking`、`refreshing` 和提交状态必须使用任务级并发锁；状态文件写入采用临时文件、fsync、原子替换，禁止两个任务互相覆盖。缓存记录 `schema_version`、`tool_version`、`case_fingerprint`、`derived_fingerprint`、`status`、`checked_at` 和 `error`。工具升级或 Schema 变化会使相关记录进入 `stale`，而不是继续复用旧结果。

- 读取 `case-prototype-index.json`，计算案例源文件和派生文件指纹；
- 扫描配置的案例根目录，发现新增、删除或未登记的 `case_id`，并将未完成契约的案例排除出候选集；
- 未变化时只验证 `case_id`、`slide_ids` 和引用路径，直接复用现有状态；
- 变化时只重开、检查并重新渲染受影响案例的页面；
- 将 freshness、validation status 和失败原因写入机器生成状态文件（默认 `.cache/case-prototype-state.json`；只读安装时使用进程内状态）；
- 不覆盖源 PPTX，不自动修改 Recipe / object map，不在 prototype record 下生成二进制副本；
- 刷新失败的记录不得进入本次运行的 prototype 候选集，并在 runtime trace 中留下原因。

候选资格与刷新动作分离：`valid` 才能进入候选集；`stale` 可以触发刷新但在刷新完成前不能被正式生成使用；`invalid` 必须排除。默认情况下，刷新只影响当前选中的案例；全量刷新和缓存清理由 CI/维护命令显式执行。正式生成是否等待刷新由策略决定：首轮 MVP 和交付级 Build 必须等待；快速 FreeStyle 可继续运行，但不得使用 stale prototype，并必须在 trace 中记录降级。

验收时至少覆盖四种场景：案例 PPTX 变化、Recipe/object map 变化、无任何变化的快速路径、两个任务并发刷新同一案例。CI 另行提供全量刷新模式，确保长期未被调用的案例也会被定期发现问题。

### 13.5 运行时追踪

每次正式生成记录：

```yaml
visual_runtime:
  rendering_pack: string
  composition_ids: [string]
  prototype_ids: [string]
  grammar_lock_version: string
  theme_lock_version: string
  resolved_theme_fingerprint: string
  seed: integer
  package_version: string
  package_module_path: string
  svg_assistance: [object]
  p01_gate: PASS|NEEDS_REVISION|BLOCKED
  final_visual_gate: PASS|NEEDS_REVISION|BLOCKED
```

### 13.6 验收

- 新增脚本在干净环境可运行；
- 错误信息包含文件、字段和修复建议；
- 校验不会要求不存在的 SVG；
- 运行时追踪能还原一次生成采用了什么 Pack、Recipe 和 Prototype；
- 资产缺失时不会静默回退到随机默认风格。

## 14. WP8：回归 deck 与对比评估

### 14.1 目标

验证新系统是否真的改变视觉产出，而不是只增加了文档、字段和目录。

### 14.2 回归主题

至少完成三套 deck：

1. **技术系统**：架构、流程、治理或 AI 系统；
2. **科学证据**：研究设计、结果、图表、限制和证据边界；
3. **品牌/建筑叙事**：空间、材质、摄影、文化或设计提案。

每套 deck 至少 6 页，覆盖不同页面任务。至少一套 deck 应包含图表或结构化数据，至少一套应包含跨页视觉 motif。

### 14.3 对比版本

每个回归主题至少生成：

- 当前 skill 基线版本；
- 新 Visual Grammar / Recipe 版本。

比较时保持：

- 相同内容；
- 相同事实和数据；
- 相同或等价图片输入；
- 相同输出尺寸；
- 可记录的 seed 和主题版本。

### 14.4 视觉评价矩阵

逐页按 0～2 分评分：

| 维度 | 0 分 | 1 分 | 2 分 |
|---|---|---|---|
| Focal point | 无主视觉 | 有主视觉但竞争明显 | 首读明确且层级稳定 |
| Composition distinction | 通用模板 | 有局部变化 | 空间关系明显区别于默认模板 |
| Domain fit | 与领域冲突 | 基本适配 | 证据、语气和结构高度匹配 |
| Readability | 关键内容不可读 | 需放大或解释 | 演示尺寸下直接可读 |
| Editability | 关键内容被烘焙 | 部分可编辑 | 关键对象可独立编辑 |
| Cross-page rhythm | 页面互不相关 | 有弱连续性 | 节奏、motif 和结构变化清晰 |

每页满分 12 分，至少 9 分；Readability 和 Editability 不得低于 1 分。溢出、遮挡、错误事实、违反模板锁或整页截图替代关键内容直接 `BLOCKED`。Visual thesis、Material、Reproducibility 作为 deck 级 evidence 字段记录，不与逐页分数重复计算。baseline / upgraded 必须使用相同 brief、内容、素材、模板、seed、包版本、页面数量和输出尺寸；主 Reviewer 评分，第二 Reviewer 复核 P01、首屏和所有低于 2 分的项目。

### 14.5 必须记录的失败

不能只记录“新版更好”。需要记录：

- 哪些页面仍然保守；
- 哪些 Pack 没有产生明显区别；
- 哪些 Recipe 导致内容和构图冲突；
- 哪些 SVG 辅助元素破坏了编辑性；
- 哪些领域不适合使用当前 Pack；
- 哪些失败需要修改资产，哪些只需要修改 Prompt。

### 14.6 验收

- 3 套回归 deck 均生成 PPTX、PDF 和 PNG；
- 每套 deck 有 acceptance contract；
- 每套 deck 有 baseline / upgraded 对比；
- 所有 `MUST` 需求通过；
- 至少两种页面构图关系在每套 deck 中出现；
- 不出现无理由的连续同构卡片页；
- 新版视觉差异能用 Pack、Grammar、Recipe 或 Prototype 解释；
- 每页评分记录完整，平均分、最低分和阻塞项达到预设阈值；
- 任何质量问题都有 revision record。

## 15. WP9：文档、安装和示例同步

### 15.1 必须同步的文档

更新：

- `skill/SKILL.md`；
- `skill/references/workflow.md`；
- `skill/references/design-direction-coach.md`；
- `skill/references/design-principles.md`；
- `skill/references/qa-and-delivery.md`；
- `skill/references/public-api.md`，如新增公共 API 用法；
- `README.md`；
- `docs/README.md`；
- `docs/README_EN.md`；
- `docs/usage-guide.md`。

同步内容至少包括：

- PPT-first / SVG-assisted 边界；
- Rendering Pack 和 Composition Catalog 入口；
- Composition Recipe 的作用；
- P01 方法门；
- 方向性返工规则；
- 案例目录与 case-prototype index 的关系；
- 运行时追踪字段；
- 资产校验命令；
- 不承诺的能力边界。

### 15.2 安装和运行

必须验证：

- Skill 安装器会复制新增的 `skill/references/` 文件；
- 新增原型和索引不会被安装器遗漏；
- `--check` 不因可选 SVG fragments 缺失而误报；
- 运行时依赖仍符合 Python `>=3.10` 和现有渲染策略；
- 旧安装目录中的遗留文件不会改变新运行路径。

### 15.3 示例

README 中应加入一个足够小但完整的示例，展示：

```text
brief
→ direction
→ rendering_pack
→ composition_recipe
→ native PPT build
→ optional local SVG assistance
→ PNG review
```

示例不能只展示色板或 `generate_ppt()` 调用，必须能体现“设计方向改变构图”的关系。

### 15.4 验收

- 中英文文档术语一致；
- 文档中的路径和命令真实存在；
- 新能力可以从 README 被发现；
- 安装后文件结构正确；
- 文档没有把 PPT-first 错写成 SVG-first；
- 没有把 Prototype 写成强制 1:1 模板。

## 16. WP10：最终审查与交付

### 16.1 结构审查

执行：

```powershell
python skill/scripts/check_runtime.py
python installer/install.py --check
python skill/scripts/validate_visual_pack.py
python skill/scripts/inspect_visual_assets.py
git diff --check
```

如新增脚本或测试，还应运行：

```powershell
python -m pytest
```

### 16.2 运行审查

对三套回归 deck：

1. 运行生成脚本；
2. 重新打开 PPTX；
3. 运行结构检查；
4. 渲染为 PDF 和 PNG；
5. 查看 contact sheet；
6. 逐页查看 PNG；
7. 对照 acceptance contract；
8. 检查 PPTX 中关键对象的编辑性；
9. 检查 runtime trace；
10. 必要时进行方向性或局部修订；
11. 重新生成、渲染和检查。

### 16.3 最终交付物

必须交付：

- 设计文档；
- 实施计划；
- Schema 和索引；
 - 首轮 2 个 Rendering Packs；
 - 首轮 4 个 Composition Families；
 - 首轮接入的案例库及 6～8 条案例派生原型记录；
 - 6 个 Pack、12 个 Family、36 条原型记录作为长期覆盖目标和覆盖率报告；
- Prompt / role instructions；
- 校验脚本；
- 测试和 fixtures；
- 3 套回归 deck；
- 每套 deck 的 PPTX、PDF 和 PNG；
- 结构检查结果；
- PNG 视觉检查结果；
- revision records；
- 安装和 README 更新。

## 17. 完整性追踪矩阵

| 设计文档目标 | 实施位置 | 验收证据 |
|---|---|---|
| 三层设计资产 | WP1、`skill/references/`、`skill/templates/` | schema、索引、样例 |
| Rendering Pack | WP2 | 首轮 2 个 Pack + index + validator；长期 6 个 |
| Composition Catalog | WP3 | 首轮 4 个 Family + Recipe 引用；长期 12 个 |
| Case-derived Page Prototype | WP4 | 首轮案例派生记录；长期目标为 36 条 |
| 角色化 Prompt | WP5 | Strategist / Planner / Executor / Reviewer 规则 |
| P01 方法门 | WP5、WP6 | P01 gate record |
| PPT-first 实现 | WP6 | PPTX object map + inspection |
| SVG 辅助边界 | WP6、WP7 | svg_assistance trace + optional validation |
| Build Mode | WP6 | Build regression deck |
| FreeStyle Mode | WP6 | FreeStyle regression / routing evidence |
| VI Build Mode | WP6 | template conflict and preservation evidence |
| 视觉验收 | WP8、WP10 | PNG review records |
| 可复现 | WP7、WP8 | seed / theme / pack / recipe trace |
| 文档和安装 | WP9 | install check + docs review |

## 18. 开始和完成的判定

### 18.1 可以开始正式制作资产的条件

- WP0 基线完成；
- WP1 schema 和索引格式通过审查；
- PPT-first / SVG-assisted 边界写入执行规范；
- 案例目录、案例派生记录和 object map 格式确定；
- 现有用户改动已识别并保护。

### 18.2 可以认为能力完成的条件

只有以下条件全部满足，才可以声明升级完成：

- 所有目标文件存在；
- 所有索引和校验通过；
- 首轮 2 个 Pack 和 4 个 Family 有实际内容；
- 首轮 6～8 条案例派生原型记录通过 `case_id` / `slide_ids` / 路径解析，并能从唯一案例取得 PPTX、预览、Recipe 和 object map 证据；
- 6 份正式 JSON Schema、引用校验和路径安全规则通过；
- Router、角色状态机、P01 gate、暂停/恢复和失败回退均有可运行产物；
- VI Build API matrix 通过，文档、Prompt 和实际公开 API 只有一套说法；
- 3 套回归 deck 完成 baseline / upgraded 对比；
- 逐页 0～2 分评分达到阈值，且首屏、P01 和低分项完成第二 Reviewer 复核；
- 所有 `MUST` acceptance rows 为 `PASS`；
- 关键对象可编辑性通过；
- PNG 视觉检查没有未处理的方向性问题；
- 中英文文档同步；
- 安装器和运行时检查通过；
- Git diff 中没有意外文件改动。

上述是首轮 MVP 的完成条件。长期能力完成还需另外满足：6 个 Pack、12 个 Family、36 条原型记录，以及每个 Pack 覆盖 6 种页面角色；这些缺口不能阻塞 MVP 交付，但必须在覆盖率报告中持续可见。

### 18.3 不得用来代替完成判定的证据

以下证据单独出现时都不能证明完成：

- Python 命令退出码为 0；
- PPTX 文件成功保存；
- shape 数量增加；
- 目录中存在若干 Markdown 文件；
- 单张封面看起来很漂亮；
- Prompt 变长；
- 只有一套领域的回归结果；
- 只有 contact sheet，没有逐页 PNG 审查；
- 只有 PNG，没有 PPTX 可编辑性检查。

## 19. 风险复核清单

最终提交前逐项确认：

- [ ] 是否把 PPT-first 误实现成 SVG-first？
- [ ] 是否把 Rendering Pack 做成了只有色板的风格文件？
- [ ] 是否把 Composition Catalog 做成了无 page job 的模板名列表？
- [ ] 是否所有原型都只是封面？
- [ ] 是否把 6 × 6 错算成 24？
- [ ] 是否每个 Pack 都有 6 个页面角色？
- [ ] 是否使用了整页截图替代可编辑 PPT？
- [ ] 是否把 SVG 用到了标题、正文或数据上？
- [ ] 是否允许了无理由的卡片网格？
- [ ] 是否把炫酷装饰放在了内容之前？
- [ ] 是否让模板规则与新视觉规则发生静默冲突？
- [ ] 是否把所有 Pack 都加载进上下文？
- [ ] 是否缺少 P01 方法门？
- [ ] 是否把方向性失败错误地当成局部格式问题？
- [ ] 是否只测试了技术领域？
- [ ] 是否只看了 PNG，没有检查 PPTX 可编辑性？
- [ ] 是否让案例和原型维护了两份 PPTX、PNG 或 SVG？
- [ ] 是否每次 Skill 使用都执行了 case-prototype preflight？
- [ ] 是否对变化案例增量刷新、对未变化案例走快速路径？
- [ ] 是否把 freshness 状态写进了不应被自动改写的源码索引？
- [ ] 刷新失败时，是否阻止失效原型进入候选集并记录原因？
- [ ] 首轮 MVP 与长期 6 × 6 覆盖目标是否被明确区分？
- [ ] 是否提交了 6 份正式 JSON Schema，而不只是 YAML 示例？
- [ ] 是否记录了字段 owner、继承优先级和冲突处理？
- [ ] 是否有真实的 Router、P01 状态和中间产物，而不是只有角色名称？
- [ ] VI Build 的 `extract_design_context()`、`extract_design_dna()` 和 `VIBuildDelivery` 是否已通过 API matrix？
- [ ] 是否记录素材来源、许可证、分发权限和脱敏状态？
- [ ] 是否使用 0～2 分逐页评分，并保持 baseline / upgraded 的控制变量一致？
- [ ] 是否明确 owner、reviewer、工时估算、并行关系和降级策略？
- [ ] 是否修改了用户已有的删除或未提交文件？
- [ ] 是否文档、索引、命令和实际路径不一致？

## 20. 结论

这份计划的执行顺序是：先锁定契约，再制作视觉资产，再接入 Prompt 和生成模式，最后用跨领域回归 deck 证明效果。它不是把能力拆成多个版本，而是保证“一次性到位”时每一层都有输入、输出和验收证据。

最终判断标准不是系统是否拥有更多模板，而是：

> 一个没有深度设计 Prompt 能力的普通用户，只提供主题、受众和内容，也能得到由 Skill 自动完成艺术指导、构图选择、PPT 原生实现和视觉复核的高完成度演示文稿。
