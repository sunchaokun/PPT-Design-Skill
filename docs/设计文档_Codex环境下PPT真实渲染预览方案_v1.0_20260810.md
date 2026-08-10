# 设计方案：Codex 环境下 PPT 真实渲染预览

> 状态：待审查（v1.1 已修正审查发现的问题）
> 日期：2026-08-10
> 版本：v1.1

---

## 一、问题根因（已确认，非推测）

**Codex 沙箱机制**：Codex 通过 `codex-command-runner.exe` 以**隔离用户**（`CodexSandboxOffline/Online`）+ 受限令牌执行命令（见 `.codex/.sandbox/setup_marker.json`）。

**结果**：沙箱进程无法访问交互式桌面会话 → PowerPoint COM 启动失败，报 **WinError 1312**（`ERROR_NO_SUCH_LOGON_SESSION` "指定的登录会话不存在"）。

**性质**：这是 Codex 的**安全设计**，不是 bug。绕过沙箱去启动 PowerPoint 是错误方向（破坏隔离、不可复现、违反平台安全模型）。当前 `render_preview` 只依赖 PowerPoint COM，一旦 1312 就完全无法出图 → Codex 看不到任何渲染结果，无法做视觉审查。

## 二、方案目标

在 Codex 沙箱（及任何无交互桌面的环境）内，获得**真实的 PPT 渲染效果图**（非线框、非分析图），供 Codex 做视觉审查。同时保留现有交互式环境下的 PowerPoint COM 高质量渲染。

## 三、架构：多引擎渲染链（选型 LibreOffice headless）

```
render_preview(pptx)
   │
   ├─① PowerPoint COM   ← 交互式桌面环境（当前已可用，保真度最高）
   ├─② LibreOffice headless ← Codex 沙箱 / CI / 无桌面环境【新增核心】
   └─③ 明确报错 + 安装指引  ← 全部失败时
```

**选型理由**：LibreOffice headless（`soffice --headless --convert-to pdf`）是纯后台进程，不依赖交互式桌面会话，可在受限令牌下运行。它是 LibreOffice 官方支持的服务器/CI 模式，行业标准方案。渲染效果与 PowerPoint 高度接近（字体、布局、图表都能正确呈现）。

## 四、实现细节

### 4.1 依赖安装

| 依赖 | 作用 | 安装方式 |
|------|------|---------|
| LibreOffice | pptx→PDF 真实渲染 | `winget install TheDocumentFoundation.LibreOffice`（~350MB） |
| poppler (pdftoppm) | PDF→PNG 高质量栅格化 | `winget install oschwartz10612.Poppler` 或 Python `pdf2image`+poppler |

**PDF→PNG 路径（修正 v1.1）**：
- **路径 A（首选）**：`pdftoppm`（poppler，`winget install oschwartz10612.Poppler`）直接转 PNG，分辨率可控、快
- **路径 B（备选）**：Python `pdf2image`（底层仍调 poppler，`pip install pdf2image`）

> **⚠️ v1.1 审查修正**：原方案的"路径 C：Pillow 纯 Python 兜底"**不成立**——已实测本机 Pillow 的 PDF 解码为 `False`（Pillow 读 PDF 依赖 Ghostscript 且非默认编译）。因此 PDF→PNG **必须**依赖 poppler（pdftoppm 或 pdf2image）。这是方案的关键修正：LibreOffice + poppler 是硬依赖，二者缺一不可，不存在纯 Python 兜底。

### 4.2 `build_helpers.py` 新增 `preview()` 一站式函数（LLM 友好）

**核心设计原则**：LLM 在 build.py 里已习惯 `from ppt_pro_max.build_helpers import *`。CLI（`python -m ...`）需要 LLM 记忆参数、拼接命令、解析输出，极易反复试错。因此**预览能力必须暴露为 build_helpers 内的一个函数**，与 `ai_image()` 同风格，让 LLM 写完 build.py 直接调用。

```python
def preview(pptx_path, out_dir=None, engine=None, open_in_browser=False):
    """Render a .pptx to PNG previews for visual inspection (one call).

    Wraps render_preview.render_preview() so build.py can see real layout
    renders without opening PowerPoint manually.

    Args:
        pptx_path: Path to the saved .pptx.
        out_dir: Output dir for PNGs + index.html (default: <pptx>/preview/<stem>/).
        engine: 'powerpoint' | 'libreoffice' | None=auto (falls back gracefully).
                In Codex sandbox / headless env, use engine='libreoffice'.
        open_in_browser: Open the HTML contact sheet (interactive env only).

    Returns:
        dict {"pngs": [Path, ...], "html": Path, "engine": str, "warnings": [str]}
    """
```

- `preview()` 从 `render_preview` re-export，LLM `import *` 后直接可用（与 `ai_image`/`fetch_image` 同一模式）
- 调用方（SKILL.md / AGENTS.md）引导 LLM：**写 build.py 末尾自动加 `preview('output.pptx')`**，Codex 环境自动落到 LibreOffice，无需 LLM 记忆 CLI

### 4.3 `render_preview.py` 改动

- `detect_engine()` 扩展：`powerpoint` / `libreoffice` / `none`，按可用性返回
- `render_preview()` 改为**引擎优先级自动降级**：
  1. 显式 `engine=` 指定时，失败即报错（不静默降级，避免掩盖问题）
  2. `engine=None` 自动时：尝试 PowerPoint → 失败 → LibreOffice → 失败 → 抛带安装指引的错误
- `_render_libreoffice()` 健壮性增强：
  - 增加 `--headless --norestore` 参数，避免会话恢复
  - 检测 pdftoppm/pdf2image/Pillow 三种转换路径
  - 输出 PDF 名称标准化（LibreOffice 可能生成 `<原名>.pdf`）
  - 超时处理（大文件）
- 新增 `_diagnose()`：COM 失败时给出可读原因（1312 → "无交互桌面，建议 Codex 环境用 engine='libreoffice'"）

### 4.4 CLI 改动（保留，作为备选）

```bash
# Codex 环境（推荐显式指定，避免浪费时间试 COM）
python -m ppt_pro_max.render_preview output.pptx --engine libreoffice

# 自动（交互式桌面会用 COM，沙箱会自动落到 LibreOffice）
python -m ppt_pro_max.render_preview output.pptx
```

`--engine` 增加 `libreoffice` 选项。CLI 仍保留，但 LLM 首选 `preview()` 函数。

### 4.5 SKILL.md / AGENTS.md 文档

- AGENTS.md 的 Render Preview 命令段：注明 **Codex 环境必须/推荐用 `--engine libreoffice`**，解释 1312 原因
- SKILL.md 的调试工作流：加一条 "Codex 环境预览用 LibreOffice headless" + `preview()` 函数用法

## 五、测试计划

| 测试 | 内容 | 方式 |
|------|------|------|
| 单元测试 | `detect_engine()` 三分支 (mock 检测) | mock，CI 可跑 |
| 单元测试 | `render_preview()` 引擎降级链 (mock) | mock，CI 可跑 |
| 单元测试 | `_render_libreoffice()` 调用 soffice 正确参数 (mock subprocess) | mock，CI 可跑 |
| 单元测试 | `preview()` 可从 build_helpers star-import + 参数透传 (mock) | mock，CI 可跑 |
| 集成测试 | 真实 LibreOffice 渲染（若环境已装）→ 断言 PNG 尺寸/非空 | `pytest.mark.skipif` 无 soffice 时跳过 |
| 回归 | 现有 test_render_preview.py 全部通过 | 本地 |

## 六、风险与边界（v1.1 更新）

| 风险 | 应对 | 审查结论 |
|------|------|---------|
| Codex 沙箱内 soffice 也被阻止 | 实测验证；若被禁，降级到方案 C | ⚠️ 未验证，实施第一步必须实测 |
| **PDF→PNG 依赖 poppler（修正）** | pdftoppm 或 pdf2image 必须装一个 | ✅ 已确认 Pillow 无法兜底 |
| LibreOffice 渲染字体差异（中文字体缺失） | Windows 系统字体 `C:\Windows\Fonts` 通常沙箱可读；已实测 msyh/simsun 存在 | ✅ 低风险 |
| poppler 未装 | 方案内置安装提示；或 `pip install pdf2image` | 需引导 |
| 首次安装 LibreOffice 较大 | 一次性成本，winget 静默安装 | 需用户确认 |

**方案 C（若 soffice 在沙箱也被禁）**：改由**真实环境预渲染**——在你有桌面的会话跑一次 `render_preview` 生成 PNG 到 `output/preview/`，Codex 直接读已生成的 PNG（沙箱能读 writableRoots 内的图）。此方案零新增依赖，但要求每次改版后先在桌面环境渲染一次。

## 七、交付物清单

1. `src/ppt_pro_max/build_helpers.py` — 新增 `preview()` 一站式函数（re-export render_preview）
2. `src/ppt_pro_max/render_preview.py` — LibreOffice 引擎 + 自动降级 + 诊断
3. `tests/test_render_preview.py` / `tests/test_build_helpers.py` — 新增 preview()/降级链/LO 测试
4. `AGENTS.md` / `skill/SKILL.md` — Codex 预览指引（preview() 用法 + 1312 说明）
5. 验证报告 — LibreOffice 真实渲染的 5 页 PNG + 与 PowerPoint 对比
6. **安装脚本/指引** — winget 一键装 LibreOffice + poppler（含 `install_deps` 说明）

---

## 待确认事项（v1.1）

1. **选型**：LibreOffice headless 是否认可？（备选方案 C 是"真实环境预渲染"）
2. **安装范围**：LibreOffice (~350MB) + poppler (~100MB) 两个系统依赖能否接受？
   （⚠️ v1.1 修正：**无纯 Python 兜底**，poppler 是硬依赖）
3. **降级策略**：自动降级 vs 显式指定引擎，你的偏好？
4. **preview() 命名**：`preview(pptx_path, ...)` 这个函数名/签名是否认可？
5. **实施顺序**：建议先装 LibreOffice + poppler 并**实测沙箱内 soffice 可行性**，再写代码——是否同意？
