# SKILL.md 压缩评估（保守版）

## 逐块重新审视

核心问题：**LLM 读 SKILL.md 时，哪些信息是做决策时真的需要的？** 不是"有没有用"，而是"读不到会不会做错"。

| 块 | 行数 | 之前建议 | 重新评估 | 理由 | 新建议 |
|----|------|----------|----------|------|--------|
| Frontmatter | 11 | 保留 | 保留 | — | 保留 |
| 标题+python-pptx引用 | 5 | 保留 | 保留 | LLM 写代码时必须知道 API 范围 | 保留 |
| 执行工作流 | 39 | 保留 | 保留 | 核心流程 | 保留 |
| Content Design Rules | 14 | 保留 | 保留 | 直接影响 content.json 质量 | 保留 |
| When to Activate | 8 | 保留 | 保留 | — | 保留 |
| **Dual-Mode Architecture** | 48 | →15 | →**30** | CLI 示例是 LLM 调用管线时的参考，不能全删。但两个模式的示例可以各保留 2 行关键命令 | 压缩 18 行 |
| **Python API** | 29 | →5 | →**15** | `generate_ppt()` 和 `fetch_image()` 的签名是 LLM 写代码的关键参考，不能只留 import。但示例可以精简 | 压缩 14 行 |
| 4-Phase Pipeline | 6 | 9→8 | 保留 | — | 保留 |
| Design Atoms | 10 | →8 | →8 | "Status" 列无信息量 | 压缩 2 行 |
| 10 Diagram Types | 14 | 保留 | 保留 | LLM 选图表类型时需要 | 保留 |
| **Image Engines** | 71 | →15 | →**30** | 引擎表+模式说明必须保留。Standalone Image 的 CLI/Python 示例可以精简但参数表有用（LLM 需要知道 width/height 默认值） | 压缩 41 行 |
| Animation System | 6 | 保留 | 保留 | — | 保留 |
| Enterprise Project Structure | 15 | 保留 | 保留 | LLM 需要知道目录结构 | 保留 |
| content.json Format | 35 | 保留 | 保留 | **最核心** — LLM 写 content.json 必须知道格式 | 保留 |
| brand.json Format | 23 | 保留 | 保留 | LLM 写 brand.json 必须知道格式 | 保留 |
| Page Revision Syntax | 15 | 保留 | 保留 | — | 保留 |
| **.env Configuration** | 14 | 删除 | **删除** | 运行时自动读取，LLM 从不需要手动配置 .env | 删除 14 行 |
| **CLI Reference** | 67 | 删除 | **→10** | 完整参数表确实不需要，但 LLM 需要知道关键参数名（style/content/enterprise/image 分组），保留分组标题+参数名列表（无描述） | 压缩 57 行 |
| **Component Library** | 129 | →60 | →**80** | Selection Strategy 和 content.json 用法必须保留。Database Management 和 Coordinate Normalization 是实现细节可删 | 压缩 49 行 |
| **Design Quality Standards** | 31 | 删除 | **→8** | 完整模块列表不需要，但 LLM 需要知道系统能力概要（如"CJK 字体、5 级字号、暗色发光"等关键能力） | 压缩 23 行 |
| Key Constraints | 18 | 保留 | 保留 | API 签名约束防止 LLM 写错代码 | 保留 |
| **Dependencies** | 6 | 删除 | **保留** | LLM 需要知道 ui-ux-pro-max 是依赖 | 保留 |

## 压缩汇总

| 操作 | 行数 | 块 |
|------|------|-----|
| 删除 | 14 | .env Configuration |
| 压缩 18 行 | Dual-Mode Architecture（48→30） |
| 压缩 14 行 | Python API（29→15） |
| 压缩 2 行 | Design Atoms（10→8） |
| 压缩 41 行 | Image Engines（71→30） |
| 压缩 57 行 | CLI Reference（67→10） |
| 压缩 49 行 | Component Library（129→80） |
| 压缩 23 行 | Design Quality Standards（31→8） |
| **总节省** | **218 行** | |

| | 当前行数 | 压缩后 | 新增设计指导 | 最终 |
|---|---|---|---|---|
| SKILL.md | 634 | 416 | +735 | **~1151** |

<details>
<summary>之前过度压缩的对比</summary>

之前建议删 384 行（47%），实际只应删 218 行（34%）。

过度压缩的部分：
- Python API 不能只留 5 行 — `generate_ppt()` 签名是 LLM 调用时的关键参考
- CLI Reference 不能全删 — LLM 需要知道参数分组
- Image Engines 不能只留 15 行 — 参数默认值（width/height）是决策依据
- Design Quality Standards 不能全删 — 关键能力概要帮助 LLM 知道系统能做什么
- Dependencies 不能删 — ui-ux-pro-max 依赖关系需要声明

</details>

## 各块压缩具体方案

### Dual-Mode Architecture（48→30）

删掉：每个模式下重复的命令变体（natural language / exact atom / design dials 分三行），只保留每种模式 1 个典型命令 + 关键参数说明。

### Python API（29→15）

删掉：每个函数的多行示例。保留：`generate_ppt()` 签名 + `fetch_image()` 签名 + 返回值格式。

### Image Engines（71→30）

删掉：Standalone Image 的 CLI 子命令示例（6 行）、Python API 重复示例（10 行）、Parameters 表（15 行）、Returns 格式（1 行）。
保留：引擎表 + 模式说明 + cache-first + standalone 一行说明 + fetch_image() 签名。

### CLI Reference（67→10）

删掉：每个参数的描述文字。保留：分组名 + 参数名列表。

```markdown
## CLI Quick Reference
- **Style**: --style, --palette, --fonts, --decoration, --layout-variant, --mood, --style-seed
- **Content**: --slides, --content, --variance, --motion, --density
- **Images**: --image-mode, --fetch-images, --llm-provider, --llm-api-key
- **Enterprise**: --project, --business-mode, --review, --pages, --history, --component-library
- **Beautify**: --beautify, --beautify-mode
- **Output**: --persist, --dry-run, -o
```

### Component Library（129→80）

删掉：Database Management（14 行）、Coordinate Normalization（7 行）、Design Quality Optimization 的 4 个子节中重复 Selection Priority 的部分（~15 行）、Node count matching tips 中冗余说明（~13 行）。

### Design Quality Standards（31→8）

删掉：完整模块列表（21 行）。保留：一行概要。

```markdown
## Design Capabilities
5-level typography · OKLCH 11-shade color · gradient overlay · 5-level elevation · CJK font pairing · adaptive margins · 10 decorations · 4 hero patterns · corner radius scale · two-column bullets · noise texture · progress bar
```
