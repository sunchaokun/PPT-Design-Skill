# 案例 Prototype 就绪度盘点

盘点时间：2026-09-04

本报告只记录当前仓库中可观察到的文件状态，不把 PNG 或 Markdown 说明自动视为可编辑 Prototype。合格 Prototype 必须能够从同一案例取得可重开的 PPTX、预览、Recipe、object map、素材许可和脱敏状态。这里明确区分：`build.py` 是可复现源代码，`output/*.pptx` 是已生成的候选源文件；在 provenance 未确认前，不能把 output PPTX 自动当作唯一事实源。

## 当前案例目录

| case_id | 当前可见资产 | 关键缺口 | 当前状态 |
|---|---|---|---|
| `vertical-city-retrofit` | brief、visual direction、page plan、acceptance contract、完整 PNG、`output/*.pptx`、`output/*.pdf` | PPTX provenance、Recipe、object map | `BLOCKED` |
| `car-t-single-cell-paper` | brief、visual direction、page plan、acceptance contract、论文图片素材、完整 PNG、`output/*.pptx`、`output/*.pdf` | PPTX provenance、Recipe、object map | `BLOCKED` |
| `louvre-abudhabi` | brief、visual direction、page plan、README、完整 PNG、`output/*.pptx`、`output/*.pdf`、ASSET_CREDITS | PPTX provenance、Recipe、object map | `BLOCKED` |
| `ai-infrastructure-economics` | brief、visual direction、page plan、acceptance contract、完整 PNG、`output/*.pptx`、`output/*.pdf` | PPTX provenance、Recipe、object map | `BLOCKED` |
| `couture-lipstick-atelier` | brief、visual direction、page plan、README、完整 PNG、`output/*.pptx`、`rendered/*.pdf`、ASSET_CREDITS、图片素材 | PPTX provenance、Recipe、object map、Pack 兼容性 | `BLOCKED` |
| `ai-agent-operating-system` | brief、visual direction、page plan、README、完整 PNG、`output/*.pptx`、`output/*.pdf` | PPTX provenance、Recipe、object map | `BLOCKED` |

## 首轮接入动作

首轮只从上述案例中选择最容易恢复完整证据链的 6～8 个页面，目标是覆盖 2 个 Pack、4 个 Family 和至少 4 种 Page Job。每个页面必须完成：

1. 确认 `build.py`、候选 PPTX、PDF、PNG 之间的 provenance；只有无法确认 provenance 时，才允许重新生成 PPTX，并记录重建原因；
2. 按现有页面真实结构编写 Recipe；
3. 用 `inspect_pptx.py` 和渲染流程确认页面对象边界；
4. 编写 object map，标出原生对象与局部 SVG；
5. 补充素材来源、许可证、分发权限和脱敏结论；
6. 通过逐页评分和 `case-prototype-index.json` 登记。

在以上步骤完成前，案例可以作为人工设计参考，但不能作为运行时 Prototype 候选，也不能计入 6～8 条首轮合格记录。

可先运行 `python skill/scripts/audit_case_outputs.py` 检查输出文件与页数闭环。该命令不会把 `render_ready` 误报为许可证、provenance 或视觉评分通过；若后续案例缺少完整 PNG，仍必须保持 `BLOCKED`。
