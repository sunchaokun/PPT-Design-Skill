# 视觉能力升级执行计划

状态：实施中

## 目标

在不触碰 `main`、不覆盖既有工作树改动的前提下，先交付可信的 MVP 工程底座，再逐步接入视觉资产。首轮完成条件是：

1. 六份 Draft 2020-12 Schema 可加载，并有标准库校验器和跨文件引用校验；
2. 案例索引、路径安全、freshness 状态和并发刷新可运行；
3. Router / P01 / Reviewer 状态可持久化并支持恢复；
4. 现有案例的事实状态准确，缺证据者保持 `BLOCKED`；
5. 通过自动化测试、运行时检查和至少一套 smoke regression。

## 工作包

| 顺序 | 工作包 | 交付物 | 完成证据 |
|---|---|---|---|
| 1 | 契约和索引 | `skill/schemas/`、三个 index、字段约束 | Schema 加载、非法 fixture 失败 |
| 2 | 资产校验 | `validate_visual_pack.py`、`inspect_visual_assets.py` | 路径、引用、Schema、可选 SVG 检查通过 |
| 3 | 案例刷新 | `refresh_case_prototypes.py`、缓存状态机、锁 | unchanged / changed / concurrent 三种测试 |
| 4 | 工作流状态 | `run_visual_workflow.py`、task 状态文件、`validate_runtime_trace.py` | route、P01、暂停、恢复、失败回退和 gate 一致性测试 |
| 5 | 案例证据 | 6～8 条有效 Prototype 或明确 `BLOCKED` | PPTX、预览、Recipe、object map、许可证据、provenance；`audit_case_outputs.py`、`audit_prototype_provenance.py` |
| 6 | 视觉资产 | 首轮 2 Pack、4 Family | 真实案例原型和 Pack 校验 |
| 7 | 回归验收 | 技术、科研、品牌/建筑三套回归、`regression-manifest.json`、`regression-readiness.md`、`audit_visual_review_records.py` | baseline/upgraded 文件对、页数一致、Review 路径一致、逐页评分和复核；`validate_acceptance_record.py`、`audit_regression_pairs.py` |
| 8 | 安装交付 | 安装器、README、完整测试 | runtime、installer（非零失败码）、pytest、diff check |

## 状态和降级

- `PASS`：证据完整，可以进入下一阶段或候选集；
- `NEEDS_REVISION`：有可定位的局部或方向性问题，不能宣布完成；
- `BLOCKED`：缺少源文件、许可、模板决策或 API 证据，必须排除出正式候选集；
- 快速 FreeStyle 可以不使用 Prototype，但必须记录降级原因；交付级 Build 不允许使用 stale / invalid Prototype。

案例输出审计只证明 PPTX/PDF/PNG 的文件和页数闭环，输出 `render_ready` 不等于视觉评分、provenance 或许可证通过；缺少完整预览的案例继续保持 `BLOCKED`。

Acceptance Record 还必须通过 `python skill/scripts/validate_acceptance_record.py <record.json>`：六项固定 criteria、逐项证据、总分求和、PASS 阈值以及 readability/editability 下限均由脚本复核。

## 分支与提交策略

本次工作从 `origin/master` 派生隔离分支，提交只包含升级相关文档、Schema、脚本、索引和测试。既有删除和其他未提交文件不自动纳入提交。完成验证后使用显式 `git push origin HEAD:master` 更新远端 `master`，不向 `main` 推送。
