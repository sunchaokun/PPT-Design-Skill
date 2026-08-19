<div align="center">

# PPT Design Skill

> 基于 build 模式的精细化 PPT 设计系统

**Precision PPT design system with 40,000+ styles, pixel-perfect build-mode control, and AI image generation**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![pptx](https://img.shields.io/badge/python--pptx-1.0.2-green.svg)](https://pypi.org/project/python-pptx/)

在 AI 编码工具中输入：`用Build模式生成一份AI融资路演PPT` → 自动调用 skill → 输出可编辑 .pptx

| Build 设计师模式 ⭐ | FreeStyle 自由模式 | VI Build 企业模式 |
|:---:|:---:|:---:|
| **像素级控制 + 方案对比** | 一句话出PPT | **基于企业模板 VI 精确生成** |
| **比肩高级设计师水平** | 30秒快速生成 | **保留框架页 + build_helpers** |

[English](docs/README_EN.md) | [使用手册](docs/usage-guide.md) | 中文

</div>

---

## ✨ 案例展示

> 5 种风格，5 种场景 — 每个案例包含封面 + 内容页，AI 配图由 Seedream 生成

### 🏢 Professional Modern — 企业融资路演

<img src="docs/showcase/showcase-professional-slide1.jpg" width="45%"/> <img src="docs/showcase/showcase-professional-slide2.jpg" width="45%"/>

*深蓝商务风 · 金色点缀 · 左侧导航栏 · 四宫格数据卡片*

### 🌌 Dark Tech — 科技产品发布

<img src="docs/showcase/showcase-dark-tech-slide1.jpg" width="45%"/> <img src="docs/showcase/showcase-dark-tech-slide2.jpg" width="45%"/>

*赛博朋克风 · 霓虹蓝紫粉 · Consolas 等宽字体 · 三列特性卡片*

### 🏛️ Warm Elegant — 奢侈品牌策略

<img src="docs/showcase/showcase-warm-elegant-slide1.jpg" width="45%"/> <img src="docs/showcase/showcase-warm-elegant-slide2.jpg" width="45%"/>

*金色大理石风 · Georgia 衬线字体 · 居中编辑式排版 · 菱形装饰符*

### 🚀 Vibrant Startup — 创业融资路演

<img src="docs/showcase/showcase-vibrant-startup-slide1.jpg" width="45%"/> <img src="docs/showcase/showcase-vibrant-startup-slide2.jpg" width="45%"/>

*紫粉渐变风 · Segoe UI · 进度条数据可视化 · 半透明统计胶囊*

### 🌿 Nature Calm — 可持续发展报告

<img src="docs/showcase/showcase-nature-calm-slide1.jpg" width="45%"/> <img src="docs/showcase/showcase-nature-calm-slide2.jpg" width="45%"/>

*森林绿风 · 圆形装饰符 · 四列影响卡片 · 左侧窄边栏*

---

## 🚀 快速开始

### 安装为 Skill（推荐）

```bash
git clone https://github.com/sunchaokun/PPT-Design-Skill.git
cd PPT-Design-Skill

# 一键安装（自动覆盖旧版本）
python install.py                     # 自动检测平台
python install.py --platform opencode # 指定平台
python install.py --force             # 强制覆盖已有安装
```

安装后 `ppt_pro_max` 注册为全局 Python 包，任何项目下可直接 `from ppt_pro_max.build_helpers import *`。

支持 13 个平台：OpenCode · Claude Code · Codex · Cursor · Windsurf · Roo Code · Gemini · Trae · Continue · Droid · KiloCode · Augment · Copilot

### 作为 Python 包使用

```bash
pip install --upgrade git+https://github.com/sunchaokun/PPT-Design-Skill.git
```

安装后即可在任何项目直接 import：

```python
from ppt_pro_max.build_helpers import *
from ppt_pro_max import generate_ppt
```

### 更新

```bash
# Skill 安装方式
cd PPT-Design-Skill && git pull && python install.py --force

# pip 安装方式
pip install --force-reinstall --upgrade git+https://github.com/sunchaokun/PPT-Design-Skill.git
```

### 在 AI 编码工具中使用

安装后，在 OpenCode / Claude Code / Codex 中直接输入：

```
用Build模式生成一份AI融资路演PPT，dark cyberpunk风格
```

AI 会自动加载 skill 并生成 .pptx 文件。

### Build Script — 逐页精确控制（推荐 ⭐）

Build 模式是**推荐的主力模式**——用 `build_helpers` 工具箱逐像素控制每一页、每个元素，3 种结构化方案自动对比，质量最高、确定性最强。

```python
from ppt_pro_max.build_helpers import *

# 设计令牌：暗色 AI 科技主题
C = {
    "primary": "#4F46E5", "accent": "#22D3EE", "background": "#0B1020",
    "card_bg": "#141B33", "text_dark": "#E2E8F0", "text_body": "#B6C2D6",
    "font_heading": "Orbitron", "font_body": "JetBrains Mono", "font_cjk": "微软雅黑",
}
T = Typography(hero=46, h1=30, h2=20, h3=16, body=14, caption=12, micro=11)

prs = Presentation()
set_widescreen(prs)
set_dark_theme(prs, C)

s = add_slide(prs)
rect(s, 0, 0, 13.333, 7.5, C["background"])
grid_background(s, spacing=0.85, color=C["divider"], alpha=7)
gradient_text(s, 1.0, 2.0, 11.3, 1.5, "让智能的成本，下降一个数量级",
              stops=[("#22D3EE", 0), ("#818CF8", 100)],
              font_size=46, bold=True, font_name="Orbitron", cjk_font="微软雅黑")
kpi_card(s, 1.0, 4.5, 3.5, 1.1, "92", "综合能力指数", C=C, typo=T)
native_chart(s, 5.0, 1.5, 7.5, 4.5, "bar",
             categories=["DeepSeek", "GPT-5", "Claude 4"],
             series=[{"name": "综合能力", "values": [92, 95, 93]}], C=C)

clean_save(prs, "output.pptx")
```

**核心工具箱**：`rect` · `rrect` · `oval` · `text` · `gradient_text` · `kpi_card` · `native_chart` · `donut_chart` · `code_block` · `section_divider` · `neon_border` · `grid_background` · `page_header` · `highlight_cards` · `set_widescreen` · `set_dark_theme` · `clean_save`

**暗色主题支持**：`set_dark_theme(prs, C)` 自动修正主题 `dk1/lt1`，确保默认文字为浅色；`clean_save(prs, path)` 清除 `printerSettings`、修复空 `<a:ln/>`、保证 PPTX 合法性。

### FreeStyle — Agent 驱动 / 一句话生成

**Path A（推荐）— 写 content.json，渲染高质量 deck**：agent 写真实内容 + 每页 `goal`，`generate_ppt` 直通渲染。这是确定性最高、质量最好的 FreeStyle 路径：

```json
{
  "slides": [
    {"goal": "hook", "title": "让智能的成本下降一个数量级", "subtitle": "DeepSeek B 轮融资"},
    {"goal": "problem", "title": "大模型竞争白热化", "bullets": ["推理成本是商业化瓶颈", "参数军备竞赛撞墙"]},
    {"goal": "data", "title": "推理成本对比",
     "chart": {"type": "bar", "categories": ["DeepSeek", "GPT-5"], "series": [{"name": "价格", "values": [0.14, 2.5]}]}},
    {"goal": "section", "title": "融资计划", "section_number": "01"},
    {"goal": "features", "title": "三大护城河", "cards": [{"title": "MoE", "text": "..."}, {"title": "RL 新范式", "text": "..."}]},
    {"goal": "cta", "title": "融资 500 亿元", "subtitle": "..."}
  ]
}
```

```python
from ppt_pro_max import generate_ppt
result = generate_ppt(content_file="content.json", style="dark-tech")  # 预设名 = 确定性输出
```

- 支持 11 种 goal（hook/problem/features/data/code/exercise/diagram/section/testimonials/cta/content）+ 图表/卡片/图形/代码等字段
- 统一版式系统：渐变背景 + 装饰、自适应字号填满版面、图表文字主题化
- 自然语言风格随机、预设名确定；配图用 `--fetch-images --llm-provider seedream`

**Path B — 一句话快速草稿**：

```bash
ppt-design "AI产品融资路演"
ppt-design "融资路演" --style "warm fintech"
ppt-design "产品发布" --style "dark cyberpunk tech"
ppt-design "ESG报告" --style "calm nature"

# AI 配图 + 动画
ppt-design "融资路演" --style "dark cyberpunk" \
  --fetch-images --llm-provider seedream --llm-api-key $ARK_API_KEY \
  --motion 7 --density 6
```

### VI Build — 基于企业模板精确生成

```bash
python -m ppt_pro_max.enterprise.template_analyzer template.pptx > analysis.txt
# 将 analysis.txt 交给 LLM 生成 build.py，然后：
python build.py
```

---

## 🔥 核心特性

| 特性 | 说明 |
|------|------|
| **三模式引擎** | FreeStyle 快速生成 + Build Script 逐页精确控制 + VI Build 企业模板合规 |
| **SVG→PPTX 编译器** | SVG 矢量图直接编译为原生可编辑 PPTX 形状（自由曲线/渐变/clipPath/use） |
| **Designer Mindset** | 内置专业设计师思维框架 — 受众优先、克制装饰、系统化思考 |
| **content.json 直通渲染** | Agent 写真实内容 + 每页 goal → `generate_ppt(content_file=...)` 直接渲染，跳过 StoryPlanner |
| **统一版式系统** | 11 种 goal 版式共享同一设计框架（标题带/内容区/页脚/页码），渐变背景 + 风格装饰 |
| **自适应排版** | 字号随内容量自动缩放 + 垂直居中，稀疏内容不再大段留白 |
| **图表主题化** | 图表系列/文字/网格线全部匹配主题色，深色浅色主题自动适配 |
| **40,000+ 风格组合** | 30 色彩方案 × 25 字体 × 15 装饰 × 12 布局，自然语言指定 `--style` |
| **AI 智能配图** | Seedream / GPT Image / DALL-E / Gemini / Wanx 5 种引擎 + Kimi 增强 |
| **python-pptx 直出** | 完全可编辑 .pptx，356x 快于 HTML→截图方案 |
| **10 种图形引擎** | 流程图 / 漏斗 / 时间线 / SWOT / 矩阵 / 循环 / 表格 / 层级 / 金字塔 / 韦恩 |
| **动画系统** | 12 种切换 + 10 入场 + 8 退场 + 8 强调 + Morph，motion 1-10 映射 |
| **CJK 字体** | 12 种中英文字体配对自动回退 |
| **内置设计数据库** | 192 色彩方案 · 84 风格 · 74 字体搭配 · 161 反模式，BM25 搜索，开箱即用 |

---

## 🏗️ 三模式架构

| | **Build Script** ⭐ | **FreeStyle** | **VI Build** |
|---|---|---|---|
| **场景** | 交付级精确控制 | 快速探索、原型 | 企业 VI 合规 |
| **触发** | `"build mode"` / `"像素级"` | 默认 | 提供 template.pptx |
| **内容** | 手写 build.py | Agent 写 content.json 或一句话 | LLM 读模板生成 build.py |
| **质量** | ★★★★★ | ★★★★（goal 驱动 11 种版式） | ★★★★★ |
| **方案** | 3 种结构化方案 | 3 种风格预览 | 3 种布局方案（同 VI Token） |

> **推荐工作流**：FreeStyle 快速原型 → Build 精细交付（主力模式）

---

## 🎨 设计系统

**自然语言风格** — 描述即生成。自然语言 style 走 **mood 检测 + 内置设计数据库**（BM25 搜索 colors/typography/styles），未指定 seed 时每次运行**随机**选取调色板/字体/装饰，不保证固定映射。要确定性输出，用预设名（`dark-tech`/`professional`/`warm-elegant`）或显式 `--seed`：

```bash
ppt-design "融资路演" --style "warm fintech"       # mood=[warm,fintech]，调色板/字体走设计数据库
ppt-design "产品发布" --style "dark cyberpunk"      # mood=[dark,neon] → 装饰多为 neon-lines，深色霓虹配色
ppt-design "品牌策略" --style "elegant luxury"      # mood=[elegant, luxury] → 玫红调 ux 配色（非金色）
ppt-design "山水诗词" --style "水墨"                # mood=[ink-wash] → 纸感浅色 + seal-stamp 装饰

# 确定性输出：预设名 / 显式原子 / 固定 seed
ppt-design "产品发布" --style "dark-tech"           # 固定 → cyber-neon 调色板 + tech-mono 字体 + neon-lines
ppt-design "融资路演" --palette ocean-blue --fonts clean-corporate --decoration accent-bar
ppt-design "融资路演" --style "warm fintech" --style-seed 42
```

**41 种 mood 关键词**：professional, tech, dark, warm, elegant, luxury, vibrant, startup, nature, calm, minimal, bold, fresh, industrial, fintech, health, education, sustainability, creative, mckinsey, consulting, pastel, retro, government, legal, pharma, realestate, automotive, aviation, energy, telecom, logistics, ink-wash, zen, sci, neon ...

<details>
<summary><strong>📐 设计原子详情</strong></summary>

| 设计原子 | 数量 | 示例 |
|----------|------|------|
| 🎨 色彩方案 | 30 | ocean-blue, cyber-neon, golden-luxury, ink-wash, zen-minimal, sci-paper... |
| ✏️ 字体搭配 | 25 | modern-sans, serif-editorial, tech-mono, ink-wash-serif, sci-serif, tech-display... |
| 🖌️ 装饰风格 | 15 | accent-bar, neon-lines, gold-trim, brush-stroke, seal-stamp, neon-glow, sci-grid, glass-panel... |
| 📐 布局变体 | 12 | standard, centered, sidebar-left, grid-2x2, scroll, ink-wash, sci-dense, hero-image... |

**30 × 25 × 15 × 12 = 135,000 种组合**

叠加内置设计数据库（192 色彩方案 · 84 风格 · 74 字体搭配 · 161 反模式）可达 200,000+

</details>

<details>
<summary><strong>🖼️ 图片引擎</strong></summary>

| 引擎 | 类型 | CLI | 默认模型 |
|------|------|-----|---------|
| `placeholder` | 占位符 | 默认 | — |
| `search` | 搜索下载 | `--image-mode search` | — |
| `seedream` | AI 生成 | `--llm-provider seedream` | `doubao-seedream-4-5-251128` |
| `gpt-image` | AI 生成 | `--llm-provider gpt-image` | `gpt-image-1` |
| `dalle` | AI 生成 | `--llm-provider dalle` | `dall-e-3` |
| `gemini` | AI 生成 | `--llm-provider gemini` | `gemini-2.5-flash-image` |
| `wanx` | AI 生成 | `--llm-provider wanx` | `wanx-v1` |
| `kimi` | 增强搜索 | `--llm-provider kimi` | `kimi-k2-0711-preview` |

所有 AI 引擎内置**缓存优先**，相同图片不重复调用 API。

</details>

<details>
<summary><strong>🏆 28 项设计质量升级</strong></summary>

**Tier 1 — 基础视觉（10 项）**：布局引擎 · 排版比例尺 · OKLCH 色彩深度 · 渐变叠层 · 5 级阴影 · 品牌条智能省略 · 图片调色 · 卡片升级 · 暗色模式修正 · 代码块重设计

**Tier 2 — 排版增强（6 项）**：CJK 字体配对 · 自适应边距 · 徽章系统 · 章节分隔页 · 装饰渲染器 · 布局变体消费

**Tier 3 — 高级视觉（7 项）**：噪点纹理 · 进度条 · 圆角系统 · 渐变线 · 图片遮罩 · 双栏要点 · Hero 4 变体

</details>

<details>
<summary><strong>🌟 高级设计效果（7 大模块）</strong></summary>

| 模块 | 能力 | API |
|------|------|-----|
| **AD-P1 文字效果** | 渐变(10预设) · 描边 · 阴影 · 发光 · 3D · 透明度 · 竖排 · 旋转 · 字间距 | `gradient_text()` / `vertical_text()` / `seal_stamp()` |
| **AD-P2 图片效果** | 异形裁切(圆/六边/菱形) · 双色调 · 灰度 · 22 种艺术效果 · 7 种 Pillow 滤镜 | `circle_image()` / `duotone_image()` / `artistic_image()` |
| **AD-P3 风格扩展** | +5 调色板 · +5 字体(楷书/仿宋/Orbitron) · +5 装饰 · +4 布局 · +5 mood | `--style "水墨"` / `--style "霓虹"` |
| **AD-P4 3D & 图案** | 3D 形状(extrusion+bevel+material) · 31 种图案填充 · 半透明面板 | `shape_3d()` / `pattern_fill()` / `frosted_panel()` |
| **AD-P5 动画扩展** | Morph 切换 · 8 种退场 · 8 种强调 | `exit_animation()` / `emphasis_animation()` |
| **AD-P6 装饰库** | 毛笔分割线 · 印章 · 卷轴框 · 霓虹边框 · 网格背景 · 半透明面板 · 墨点飞溅 | `brush_divider()` / `neon_border()` / `ink_splash()` |
| **AD-P7 模式集成** | mood → 文字效果/图片效果自动映射 · `compose()` 返回效果字段 | `--style "水墨"` 自动触发 |

**代码示例**：

```python
from ppt_pro_max.build_helpers import *

prs = Presentation()
s = add_slide(prs)
gradient_text(s, 1.0, 1.0, 8.0, 1.5, "标题", preset='gold-shine', font_size=44)
circle_image(s, 6.5, 3.0, 1.0, "photo.jpg")
shape_3d(s, 1.0, 3.5, 3.0, 2.0, depth=15.0, material='metal')
frosted_panel(s, 5.0, 3.0, 6.0, 3.0, tint='#1A1A3A', alpha=20)
brush_divider(s, 1.0, 5.0, 6.0, color='#2C2C2C')
seal_stamp(s, 11.0, 5.5, 0.8, "印", rotation=-15)
prs.save("output.pptx")
```

</details>

---

## 🧪 测试

### 运行测试

```bash
# 从项目根目录（conftest.py 自动将 src/ 加入 sys.path，确保加载 V2 源码而非旧版安装包）
python -m pytest tests/ -q

# 仅运行新增的核心严格测试
python -m pytest tests/test_design_search.py tests/test_ui_ux_adapter.py \
  tests/test_build_qa_v2.py tests/test_theme_colors.py -q
```

> **⚠️ 源码优先级**：系统可能已安装旧版 `ppt_pro_max`（site-packages）。pytest 由 `tests/conftest.py` 的 `sys.path.insert` 保证加载 `src/` 下的 V2 源码。但**手动运行 `python build.py` 或 `python -m ppt_pro_max` 时**，需确保加载的是 V2：
> ```powershell
> $env:PYTHONPATH = "src"; python build.py
> # 或安装为开发包
> pip install -e .
> ```

### 已知跳过（7 个，数据依赖问题，非代码缺陷）

```bash
python -m pytest tests/ -q --ignore=tests/test_group_audit.py \
  --ignore=tests/test_image_fetcher.py --ignore=tests/test_pptx_capabilities.py \
  --ignore=tests/test_xml_extraction.py --ignore=tests/test_analyze_template.py
```

| 跳过文件 | 原因 |
|----------|------|
| `test_group_audit.py` | 依赖已移除的组件库数据 |
| `test_image_fetcher.py` | 模型版本不匹配（doubao-seedream-4-5 vs 5-0） |
| `test_pptx_capabilities.py` / `test_xml_extraction.py` | 需要特定的 .pptx 样本文件 |
| `test_analyze_template.py` | 依赖已移除的 `analyze_template` 模块 |

### 覆盖矩阵（1479 tests, 0 failures）

| 模块 | 测试文件 | 覆盖 |
|------|----------|------|
| **内置设计数据库** | `test_design_search.py` (50) | BM25 搜索、domain 检测、设计系统生成、dial 解析 |
| **ui_ux_adapter 契约** | `test_ui_ux_adapter.py` (19) | 4 个消费者的 API 依赖字段、优雅降级 |
| **BuildQA 三级判定** | `test_build_qa_v2.py` (18) | 装饰性出血 vs 内容溢出、严重度边界、报告一致性 |
| **主题色写入** | `test_theme_colors.py` (15) | C→clrScheme 映射、持久化回环、XML 良构 |
| **设计质量升级** | `test_design_quality.py` (95) + `test_design_integration.py` (25) | 28 项 DQ 升级 |
| **图形引擎** | `test_diagram_engine.py` (75) + `test_shape_functions.py` (96) | 10 种图形、形状工厂 |
| **高级设计效果** | `test_3d_pattern_frosted.py` (38) + `test_text_effects*.py` (65) + `test_animation_expansion.py` (33) + `test_decoration_library.py` (45) | 3D/图案/文字/动画/装饰 |
| **原生图表** | `test_native_chart.py` (52) + `test_chart_renderer.py` (20) | 图表类型、样式、主题化 |
| **LLM 配置适配** | `test_llm_config_adapter.py` (84) | 13 平台 LLM 配置 |
| **Build helpers** | `test_build_helpers.py` (32) + `test_build_helpers_integration.py` (30) | 颜色解析、形状、数据组件 |
| **其他** | 其余 40+ 文件 | 规划/决策/内容/渲染/提取器/图片 |

---

## License

MIT
