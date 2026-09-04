# 回归验收就绪度

盘点时间：2026-09-04

## 当前结论

三套首轮回归目标已经登记，但都保持 `BLOCKED`：

| regression_id | 领域 | upgraded | 阻塞原因 |
|---|---|---|---|
| `technical-infrastructure` | 技术 | `ai_infrastructure_economics` | 缺少可追溯 baseline PPTX 和控制变量记录 |
| `scientific-evidence` | 科研 | `car_t_single_cell_paper` | 缺少可追溯 baseline PPTX 和控制变量记录 |
| `brand-architecture` | 品牌/建筑 | `louvre_abudhabi` | 缺少可追溯 baseline PPTX 和控制变量记录 |

Git 历史审计未在 `origin/master` 或 `origin/main` 找到这三套目标 PPTX 的历史版本，因此不能把当前 upgraded 文件复制或回退后冒充 baseline。当前案例源码是 Visual Grammar / Pack 版本，不能反向推断旧版视觉结果。

## 已确定的补齐方式

三套 baseline 将作为独立生成物重新制作，而不是复制现有 PPTX：

1. 沿用同一案例的 brief、事实、数据、素材、页数和输出尺寸；
2. 使用记录明确的 `pptx-designer` 版本和固定主题参数；
3. 采用中性、未接入 Visual Grammar / Rendering Pack 的基础布局，确保 baseline 与 upgraded 是不同的实际构建结果；
4. 在独立目录保存 baseline 的 `build.py`、PPTX、PDF、PNG 和控制变量记录；
5. 由 `audit_regression_pairs.py` 检查文件独立性和页数，再进行逐页评分。

在上述独立 baseline 生成前，三套回归继续保持 `BLOCKED`；不得使用复制、重命名或回退当前 upgraded 文件的方式解锁。

## 解阻条件

每套回归必须补齐独立的 baseline PPTX，并记录相同的 brief、内容、素材、模板、seed、`pptx-designer` 版本、页面数量和输出尺寸。完成后运行：

```powershell
python skill/scripts/audit_regression_pairs.py
```

审计通过后，再按页生成 Acceptance Record，由主 Reviewer 评分，第二 Reviewer 复核 P01、首屏和低分项。`render_ready` 只能证明文件与页数闭环，不能替代 baseline、视觉评分或许可证审查。
