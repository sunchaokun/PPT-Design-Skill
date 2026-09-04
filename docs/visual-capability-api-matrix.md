# VI Build API compatibility matrix

核验环境：Windows / Python 3.12 / `pptx-designer 1.0.0b10`，核验日期：2026-09-04。

| 名称 | 实际位置 | 签名/输入 | 当前结论 |
|---|---|---|---|
| `extract_design_context` | `pptx_designer` 顶层导出 | `(pptx_path: str) -> dict[str, Any]` | VI Build 模板上下文主入口 |
| `extract_design_dna` | `pptx_designer` 顶层导出 | `(pptx_path: str) -> dict[str, Any]` | legacy compatibility projection；不作为主入口 |
| `merge_vi_design_context` | `pptx_designer` 顶层导出 | `(template_context, *overrides) -> dict[str, Any]` | 模板锁保护和合并入口 |
| `VIBuildDelivery` | `pptx_designer.enterprise.vi_delivery` | `(presentation, adapter)` | 企业交付对象；需要独立 adapter 行为测试 |

## 已完成的行为核验

对 `examples/new_examplex/louvre_abudhabi/output/louvre_abudhabi_complete.pptx`
执行了 `extract_design_context()`、`extract_design_dna()` 和
`merge_vi_design_context(context, {page_role: content})`：三者均可调用；合并
结果包含 `diagnostics`，且没有 conflict。该测试证明的是导出和基本合并行为，
不是完整模板交付兼容性。

## 未完成 / 阻塞

`VIBuildDelivery` 尚未用最小真实模板和 adapter 完成 `add → finalize` 回归，
因此 VI Build 的完整 MVP 仍为 `BLOCKED`。在该回归完成前，文档只推荐：

```text
extract_design_context()
→ merge_vi_design_context()
→ VIBuildDelivery（待 adapter fixture 验证）
```

`extract_design_dna()` 只作为兼容分析入口保留，不得在新流程中与
`extract_design_context()` 并列作为两条主路径。
