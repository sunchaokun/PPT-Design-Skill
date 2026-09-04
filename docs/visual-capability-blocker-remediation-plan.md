# 视觉能力升级前置阻塞项解除计划

状态：可执行计划

本计划把视觉能力升级中影响工程落地的三个前置问题单独拆出，先建立可信的工程底座，再继续制作 Rendering Pack、Composition Family 和更多案例原型。

## 1. 为什么先采用 MVP

MVP 在这里不是降低最终目标，也不是把架构拆成一个“简陋版本”。它只定义第一轮可验收闭环：

- 能用正式 Schema 验证资产和运行记录；
- 能知道当前 VI Build 到底调用哪些公开 API；
- 能从真实案例得到可追溯的 Prototype Record，或明确证明该案例仍然 `BLOCKED`；
- 能让一次 Skill 使用完成路由、选择、P01、生成和验收追踪。

当前仓库有 6 个案例目录，六个案例均已具备候选 PPTX/PDF 和完整 PNG 渲染；8 条代表页已登记 Recipe、object map 和结构化 Prototype Record。项目负责人已确认案例素材均为自制并批准使用，六个案例也均已通过源码权威的双重重建；因此素材许可、context permission、脱敏、provenance 和首轮 Pack 兼容性不再是当前阻塞项。MVP 所需的编排已由 `run_visual_workflow.py` 的持久化 workflow runner 覆盖，不把“尚未引入专用多代理服务”误列为首轮阻塞；多角色 orchestrator 仍属于后续扩展。此时直接宣称完成 36 条原型，会把“文件数量”误当成“有效设计资产”。因此 MVP 是证据门，不是最终能力上限。

长期目标仍保持不变：6 个 Rendering Pack、12 个 Composition Family、36 条案例派生 Prototype Record，以及三种生成模式的完整回归覆盖。

## 2. 三条实施轨道

### Track A：正式 JSON Schema

目标：把 YAML 示例变成机器可验证的 Draft 2020-12 Schema。

交付文件：

```text
skill/schemas/rendering-pack.schema.json
skill/schemas/composition-entry.schema.json
skill/schemas/composition-recipe.schema.json
skill/schemas/prototype-record.schema.json
skill/schemas/runtime-trace.schema.json
skill/schemas/acceptance-record.schema.json
```

执行顺序：

1. 固定共享定义：ID、版本、相对路径、状态、评分和时间戳；
2. 固定各对象的 required、类型、枚举、数组最小长度和 `additionalProperties: false`；
3. 增加跨文件引用校验规则：Pack、Family、Recipe、Prototype、Case 和 slide ID；
4. 为 `occupied_zones`、`native_object_plan`、`svg_assistance`、acceptance criteria 定义真实对象结构；
5. 编写最小 valid / invalid fixtures；
6. 接入 `validate_visual_pack.py` 和测试。

当前环境检查：6 份 Schema 均可被 Python JSON 解析，但当前环境未发现 `jsonschema` 包，因此尚未完成语义校验。Track A 必须在实现阶段二选一：把 `jsonschema` 纳入安装依赖并锁定版本，或实现明确受限、覆盖上述契约的内置校验器；不能把“JSON 能解析”当作 Schema 校验通过。

通过标准：Schema 可独立加载；合法 fixture 通过；缺字段、非法 ID、绝对路径、`..` 路径、越界 zone、无效引用和错误状态均能失败。

### Track B：VI Build API 统一

目标：以实际安装的 `pptx_designer` 公共导出和 `skill/references/public-api.md` 为唯一依据，消除 `extract_design_dna()`、`extract_design_context()`、`merge_vi_design_context()`、`VIBuildDelivery` 的三套说法。

执行顺序：

1. 记录当前 Python 包版本、模块路径和实际导出；
2. 对四个名称分别确认“存在 / 不存在、公开 / 内部、输入输出、推荐用途”；
3. 运行最小模板 fixture，验证提取、合并、原子 Build 和交付链路；
4. 选择唯一的 VI Build 主调用链；
5. 同步 `skill/SKILL.md`、`skill/references/public-api.md`、`skill/references/template-brand.md`、README 和中英文文档；
6. 为 API matrix 增加回归测试和版本记录。

通过标准：VI Build 文档只保留一条推荐调用链；实际公开 API、示例代码和测试一致；不支持的名称只能作为明确标注的兼容别名或内部实现，不能继续作为并列入口。

初步审计结果和行为记录见 [`docs/visual-capability-api-matrix.md`](visual-capability-api-matrix.md)。当前推荐主链统一为 `extract_design_context() → merge_vi_design_context() → VIBuildDelivery`；`extract_design_dna()` 仅保留为兼容/底层分析入口。`VIBuildDelivery.add → finalize` 已通过最小 fixture、adapter 和模板页所有权回归；真实品牌模板的 slot/atomic plan、模板锁冲突和完整交付回归仍未完成，因此 VI Build 完整能力继续保持 `BLOCKED`。

### Track C：案例证据与 Prototype 登记

目标：把现有案例分成可直接接入、补证据后接入和明确阻塞三类，不制作重复资产。

执行顺序：

1. 扫描 `examples/new_examplex/` 和其他案例根目录；
2. 为每个案例建立 `case_id`、来源、PPTX、预览、Recipe、object map、素材许可和脱敏状态清单；
3. 对已有渲染 PNG 的案例，确认可重开的 PPTX 与 `build.py`、PDF、PNG 的 provenance；只有 PNG 不能成为合格 Prototype；
4. 先补齐最有价值的 6～8 个页面证据，优先覆盖 2 个 Pack、4 个 Family 和至少 4 种 Page Job；
5. 生成 `case-prototype-index.json` 的原型登记；
6. 运行结构检查、PPTX→PDF→PNG 渲染和逐页视觉验收；
7. 缺失源文件、许可或可编辑性证据的案例保留在 `BLOCKED` 清单，不计入 Prototype 数量。

当前盘点结果见 [`docs/case-prototype-readiness.md`](E:/PPT-Design-Skill-V2/docs/case-prototype-readiness.md)。

通过标准：每个有效 Prototype 都能由 `case_id + slide_ids` 定位到唯一案例页面，并取得同一案例中的 PPTX、预览、Recipe、object map 和许可状态；不存在原型副本资产。

## 3. 依赖和并行关系

```text
Track A Schema ─────┐
                    ├─→ Track C Prototype 登记 ─→ MVP 回归验收
Track B API 审计 ───┘
```

Track A 与 Track B 可以并行。Track C 的盘点可以立即开始，但正式登记必须等待 Prototype Schema 和 API 兼容结论；案例素材补齐与 Schema 编写可以并行。

## 4. 当前阻碍与解除条件

| 阻碍 | 当前影响 | 解除条件 |
|---|---|---|
| 缺少正式 Schema | 不能可靠验证资产语义和引用关系 | 6 份 Schema、fixtures、校验测试通过 |
| VI Build API 不一致 | 不能安全承诺模板模式可运行 | API matrix、最小 fixture 和文档同步通过 |
| 案例缺少源 PPTX / Recipe / object map | 不能把渲染图计为可编辑 Prototype | 补齐证据或明确 `BLOCKED` |
| 没有真正的角色状态机 | Prompt 文案不能自动产生编排 | 先用 SKILL workflow + 中间状态文件落地，不要求新增多代理服务 |
| 缺少量化视觉基线 | 不能证明升级有效 | 逐页 0～2 分评分、固定控制变量、双 Reviewer |
| 自动刷新尚未实现 | 案例更新可能导致陈旧引用 | `refresh_case_prototypes.py`、锁、缓存状态机和并发测试通过 |

## 5. 不在本计划首轮内强行完成的内容

- 不首轮制作全部 36 条原型；
- 不为案例再复制一套 Prototype PPTX / PNG；
- 不在 API 未确认前修改 VI Build 代码；
- 不把只有 PNG 的案例包装成可编辑案例；
- 不用 Prompt 变长代替运行时状态机和验收证据。

## 6. 完成判定

只有 Track A、B、C 都达到通过标准，并完成至少一套端到端 smoke regression，才允许解除前置阻塞项。完整 MVP 仍须按主实施计划完成三套跨领域 baseline/upgraded 回归；长期资产仍可继续扩充，但不需要重做基础架构。
