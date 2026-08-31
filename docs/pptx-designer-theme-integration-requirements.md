# ppt-design-skill 修订清单：接入新版 pptx-designer

状态：执行版（2026-08-31）

## 1. 新版库已提供的能力

`E:\pptx-designer` 已提供以下能力，skill 不得重复实现：

- `generate_ppt(theme=...)`：接收完整主题上下文；
- `validate_resolved_theme()`：在 FreeStyle 渲染前验证完整主题契约；
- `Presentation(theme=..., strict_theme=True)`：提供 presentation 级主题继承
  并可拒绝不完整的普通 Build Mode 主题；
- slide 级主题覆盖和显式局部覆盖；
- 颜色语义角色、字体、装饰和布局变体的解析；
- text、shapes、SVG、layout、images、charts、cards 等 helper 的主题继承；
- `normalize_design_context()`、`merge_design_context()`、
  `merge_vi_design_context()` 和 `VIBuildSession`；
- BuildSpec 与 VI Build 的模板约束、资源和验收上下文；
- 主题来源、seed、fallback 和部分未应用效果的诊断信息。

`theme` 参数只接收 `ThemeComposer.compose()` 产生的 complete resolved theme。
Theme Lock 是 skill 层的设计状态，不能直接传入 `generate_ppt()`。如果同时传入
完整主题和 `style`、palette atoms 或 `style_seed` 等发现参数，库会发出 warning
并在结果中记录 ignored arguments。

主题、Build Spec、VI Context 相关回归测试已通过。当前 Python 环境必须确认
实际导入的是升级后的库，而不是旧版 `site-packages`。

## 2. skill 必须修订的内容

### 2.1 主题选择与锁定

在 task-init 阶段增加主题锁定结果，至少记录：

```yaml
theme_lock:
  version: 1
  status: active
  style: dark-tech
  seed: 17
  visual_thesis: "一句话视觉主张"
  audience_promise: "观众应感受到什么"
  visual_grammar: {}
  forbidden_patterns: []
```

规则：

- 主题只在整稿级别选择一次；
- LLM 负责解释主题如何影响构图、密度、焦点和节奏；
- 不要每页重新选主题或手工重建一份颜色字典；
- 局部覆盖必须有页面语义理由。
- Theme Lock 必须与 resolved theme 分开保存：前者保存设计决策，后者保存
  `ThemeComposer.compose()` 的完整返回值和其 fingerprint。

### 2.2 主题版本必须是项目唯一事实源

主题锁定不是写入 prompt 后就永久有效，而是项目中的可更新状态。必须遵循：

1. 在项目目录保存一个当前主题文件，例如
   `ppt_tasks/<task>/theme-lock.yaml`；
2. 每次用户明确要求改变风格、字体、色彩、装饰或布局语言时，创建新的
   `version`，完整替换当前文件，不在旧版本上隐式叠加；
3. 生成脚本只读取当前 resolved-theme 文件；不得把主题 token、`C` 或 `typo`
   长期硬编码在页面脚本中。局部 `C` 覆盖只能由当前 resolved theme 派生，并写明
   页面语义理由；
4. 每次生成把 `theme_version`、resolved-theme fingerprint、主题来源、seed、
   包版本和模块路径写入生成结果及 QA 记录；
5. 旧版本保留为历史记录，但不能作为后续生成的默认输入；
6. 如果用户只修改内容，不修改风格，继续使用当前版本，不重新随机组合主题。

推荐的状态流转：

```text
theme-v1 active
      ↓ 用户确认风格修订
theme-v1 superseded → theme-v2 active
      ↓ 后续生成
只允许读取 theme-v2
```

### 2.3 生成前必须执行主题一致性检查

每次生成前，skill 必须检查：

- task-init、当前主题文件、页面计划和生成脚本引用的主题版本是否一致；
- `generate_ppt(theme=...)` 或 `Presentation(theme=..., strict_theme=True)`
  使用的是否为当前 resolved theme；
- resolved-theme fingerprint、主题版本、实际导入的包版本和模块路径是否一致；
- 局部 `C` 或 typography 覆盖是否由当前 resolved theme 派生且具备页面语义理由；
- 用户最新修订是否已经明确确认，未确认的候选方向不得成为 active 版本。

发现版本不一致时，停止生成并先同步主题，不允许用旧主题继续产出。

### 2.4 FreeStyle 接入

统一使用：

```python
from pptx_designer import generate_ppt, validate_resolved_theme
from pptx_designer.renderer.theme import ThemeComposer

resolved_theme = ThemeComposer().compose(style="dark-tech", seed=17)
validate_resolved_theme(resolved_theme)
result = generate_ppt(content=content, theme=resolved_theme, output=output_path)
```

skill 需要在生成结果中保留 `theme_context`，并检查主题来源、seed、fallback
和应用诊断。

### 2.5 Build Mode 接入

创建 presentation 时传入主题：

```python
prs = Presentation(theme=resolved_theme, strict_theme=True)
```

之后优先调用公开 helper，让其继承 presentation 主题；只有确有页面语义时
才显式传入 `C`、`typo` 或局部样式。skill 不需要为每个 helper 编写主题版，
也不需要新增组件层。

### 2.6 VI Build Mode 接入

主题不能绕过模板 DNA。生成前按以下顺序合并：

```text
模板上下文 + resolved theme + 页面约束
                         ↓
           merge_vi_design_context()
                         ↓
       conflicts 为空才进入 VIBuildSession / BuildSpec
```

skill 必须保留模板锁定的 logo、字体、边距、颜色和固定装饰。使用通用
`merge_design_context()` 不构成 VI 保护；必须调用 `merge_vi_design_context()`，
并将任何冲突作为预检失败。若需例外的品牌变更，必须先形成用户明确确认的新版
template context，并在其中审慎修改或移除对应 lock 后重新执行受保护合并；不能把
被拒绝的 context 直接传入生成。

## 3. SKILL.md 和 references 需要修改的文件

- `skill/SKILL.md`：更新三种模式的主题接入方式、主题版本和生成前一致性检查；
- `skill/references/public-api.md`：补充 `theme`、`Presentation(theme=...)`、
  design context 和 VI Build 的公开用法；
- `skill/references/workflow.md`：加入主题版本更新、当前版本读取和生成前一致性检查；
- `skill/references/template-brand.md`：明确 VI Build 中主题与模板 DNA 的合并顺序；
- `skill/references/qa-and-delivery.md`：增加主题兑现检查，不能只检查换色；
- `skill/references/install-and-runtime.md`：增加新版 `pptx-designer` 导入路径检查。

## 4. 新的生成与验收要求

每个交付任务必须：

1. 锁定主题并记录 seed；
2. 按 FreeStyle、Build Mode 或 VI Build Mode 的对应入口传入主题；
3. 运行结构检查；
4. 执行 PPTX → PDF → PNG；
5. 检查主题是否影响了视觉语言，而不只是颜色：构图、焦点、密度、节奏、
   字体层级、图表和装饰；
6. 发现问题时先判断是 LLM 设计决策、主题接入、渲染器还是导出链路；
7. 重新生成并复查 PNG。

## 5. 明确不修改的内容

- 不新增大量主题专属组件；
- 不复制竞争对手的固定页面模板；
- 不新增第二套 DSL 或 API 包装；
- 不强制每页使用高级效果；
- 不让库替 LLM 决定叙事、构图和页面重点。
