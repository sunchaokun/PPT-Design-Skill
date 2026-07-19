<div align="center">

# PPT Design Skill

**三种模式 · 叙事驱动 · AI 配图 · 40,000+ 风格组合 · ui-ux-pro-max 设计智能 · 6 种 AI 图片引擎**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![pptx](https://img.shields.io/badge/python--pptx-1.0.2-green.svg)](https://pypi.org/project/python-pptx/)

| FreeStyle 自由模式 | Build 设计师模式 | VI Build 企业模式 |
|:---:|:---:|:---:|
| 一句话出PPT | **像素级控制 + 方案对比** | **基于企业模板 VI 精确生成** |
| 30秒快速生成 | **python-pptx 精确构建** | **保留框架页 + build_helpers** |

适配 OpenCode · Claude Code · Codex · Cursor

依赖 ui-ux-pro-max（192 色彩方案 · 84 风格 · 74 字体 · 161 反模式）

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

## 🔥 核心特性

| 特性 | 说明 |
|------|------|
| **三模式引擎** | FreeStyle 快速生成 + Build Script 逐页精确控制 + VI Build 基于企业模板生成 |
| **叙事引擎** | 3 种策略（YC Seed Deck / Product Demo / Sales Pitch）+ Duarte Sparkline 情绪弧线 |
| **ui-ux-pro-max 设计智能** | 192 色彩方案 · 84 风格 · 74 字体 · 161 反模式 — 必需依赖，自动匹配行业/场景 |
| **40,000+ 风格组合** | 25 色彩方案 × 20 字体搭配 × 10 装饰风格 × 8 布局变体 + 35 种 mood |
| **自然语言风格** | 描述风格即生成：`--style "warm fintech"` / `--style "dark cyberpunk"` |
| **28 项设计质量升级** | OKLCH 色彩深度 · 阴影层级 · 渐变叠层 · 进度条 · 圆角系统 · CJK 字体配对 · 噪点纹理 · 双栏要点 · 4 种 Hero 布局 · 章节分隔页 · 徽章系统 · 渐变线 · 图片遮罩 · 装饰渲染器 · 代码块重设计 · 卡片升级 · 自适应边距 · 排版比例尺 |
| **10 种图形引擎** | 流程图/漏斗/时间线/SWOT/矩阵/循环/表格/层级/金字塔/韦恩 |
| **Build Script 模式** | 10 种页面模板 + Design Token 系统 + 生成后自动检查，交付级质量 |
| **VI Build 模式** | `analyze_template.py` 提取模板 VI → LLM 生成 build.py → `build_helpers` 精确构建，企业 VI 合规 |
| **python-pptx 直出** | 完全可编辑 .pptx，356x 快于 HTML→截图方案 |
| **12 母版布局** | 13.333"×7.5" 16:9 精确坐标 |
| **AI 智能配图** | Seedream / GPT Image / DALL-E / Gemini / Wanx — 5 种生成引擎 + Kimi 增强 |
| **动画系统** | 12 种切换 + 10 种入场动画，motion 1-10 智能映射 |
| **代码块/练习** | 深色代码块 + 语言徽章 + 练习徽章 + 步骤列表 — 教育场景标配 |
| **CJK 字体** | 12 种 CJK 配对自动回退（Microsoft YaHei / STSong / SimHei 等） |

---

## 🚀 快速开始

### 安装

```bash
git clone https://github.com/sunchaokun/PPT-Design-Skill.git
cd PPT-Design-Skill

# 一键安装 — 自动检测 AI 平台 + 安装 skill + pip 依赖 + ui-ux-pro-max
python install.py

# 手动安装
pip install -e .

# 单独安装 ui-ux-pro-max（必需依赖）
npx ui-ux-pro-max-cli init --ai opencode
```

### FreeStyle — 一句话生成

```bash
# 最简用法
ppt-design "AI产品融资路演"

# 自然语言风格 — 40,000+ 组合
ppt-design "融资路演" --style "warm fintech"
ppt-design "产品发布" --style "dark cyberpunk tech"
ppt-design "ESG报告" --style "calm nature"

# AI 配图 + 风格 + 动画
ppt-design "融资路演" --style "dark cyberpunk" \
  --fetch-images --llm-provider seedream --llm-api-key $ARK_API_KEY \
  --motion 7 --density 6
```

### VI Build — 基于企业模板精确生成

```bash
# Step 1: 分析企业模板
python -m ppt_pro_max.analyze_template template.pptx > analysis.txt

# Step 2: 将 analysis.txt 交给 LLM，生成 build.py

# Step 3: 运行生成
python build.py
```

> 保留框架页（封面/目录/封底）+ `build_helpers` 精确构建新增页，详见 [使用手册 §5](docs/usage-guide.md#5-vi-build-模式--基于模板-vi-的精确生成)

### Build Script — 逐页精确控制（推荐模式）

```python
# build.py — 直接用 python-pptx 逐页精确控制
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
sl = prs.slides.add_slide(prs.slide_layouts[6])

# 每个元素精确指定 x, y, w, h, font, size, color
tb = sl.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(8), Inches(1.5))
run = tb.text_frame.paragraphs[0].add_run()
run.text = "标题"
run.font.name = "Space Grotesk"  # run-level 字体，PowerPoint 认
run.font.size = Pt(52)
run.font.bold = True

prs.save("output/presentation.pptx")
```

> 10 种页面模板 + Design Token 系统 + 生成后检查脚本，详见 [使用手册 §4](docs/usage-guide.md#4-build-script-模式--精确控制)

---

## 🏗️ 三模式架构

### FreeStyle Pipeline

```
用户输入 → 叙事规划 → 设计决策 → 内容生成 → PPT渲染 → .pptx
```

适用：快速探索、风格实验、个人演示

### Build Script 模式

```
build.py → Design Token → 页面模板 → python-pptx → .pptx → check.py 验证
               ↓              ↓            ↓
           颜色/字体      10种模板     逐元素精确控制
           一键切换主题   x/y/w/h     run-level 字体
```

适用：**最终交付、无模板约束的精确控制、质量保证**

### VI Build 模式

```
template.pptx → analyze_template.py → LLM 生成 build.py → build_helpers → .pptx
                      ↓                       ↓                  ↓
                  提取 VI Token          copy_decorations    kpi_card / bar_chart
                  色/字/布局            copy_logo           page_header
```

适用：**企业 VI 合规、必须遵守品牌模板的交付**

> **推荐工作流**：FreeStyle 原型 → VI Build（有企业模板）或 Build Script（无模板）精细交付

### 项目目录结构

**Build Script 模式：**

```
my-project/
├── build.py            # 核心构建脚本
├── images/             # 本地图片素材
└── output/             # 生成输出
    └── presentation.pptx
```

**VI Build 模式：**

```
my-project/
├── build.py            # LLM 生成的构建脚本（使用 build_helpers）
├── template.pptx       # 企业模板（VI 来源）
├── analysis.txt        # 模板分析输出（LLM 输入）
├── images/             # 本地图片素材
└── output/             # 生成输出
    └── presentation.pptx
```

---

## 🖼️ 图片引擎

| 引擎 | 类型 | CLI | 默认模型 |
|------|------|-----|---------|
| `placeholder` | 占位符 | 默认 | — |
| `search` | 搜索下载 | `--image-mode search` | — |
| `seedream` | AI 生成 | `--llm-provider seedream` | `doubao-seedream-5-0-260128` |
| `gpt-image` | AI 生成 | `--llm-provider gpt-image` | `gpt-image-1` |
| `dalle` | AI 生成 | `--llm-provider dalle` | `dall-e-3` |
| `gemini` | AI 生成 | `--llm-provider gemini` | `gemini-2.5-flash-image` |
| `wanx` | AI 生成 | `--llm-provider wanx` | `wanx-v1` |
| `kimi` | 增强搜索 | `--llm-provider kimi` | `kimi-k2-0711-preview` |

### Seedream 可用模型

| 模型 | 说明 |
|------|------|
| `doubao-seedream-5-0-260128` | **默认**，Seedream 5.0 |
| `doubao-seedream-5-0-pro-260628` | Seedream 5.0 Pro |
| `doubao-seedream-4-5-251128` | Seedream 4.5 |

所有 AI 引擎内置**缓存优先**，相同图片不重复调用 API。

---

## 🎨 设计系统 — 40,000+ 风格组合

| 设计原子 | 数量 | 示例 |
|----------|------|------|
| 🎨 色彩方案 | 25（+ 192 ui-ux-pro-max） | ocean-blue, cyber-neon, golden-luxury, wine-burgundy, monochrome-dark... |
| ✏️ 字体搭配 | 20（+ 74 ui-ux-pro-max） | modern-sans, serif-editorial, tech-mono, elegant-serif, contrast-mix... |
| 🖌️ 装饰风格 | 10 | accent-bar, neon-lines, gold-trim, gradient-bar, circle-accent, sidebar-nav... |
| 📐 布局变体 | 8 | standard, centered, sidebar-left, grid-2x2, wide-cards, full-width... |

**25 × 20 × 10 × 8 = 40,000 种组合**（叠加 ui-ux-pro-max 可达 200,000+）

### 自然语言风格

```bash
ppt-design "融资路演" --style "warm fintech"          # → ocean-blue + clean-corporate + accent-bar
ppt-design "产品发布" --style "dark cyberpunk"         # → cyber-neon + tech-mono + neon-lines
ppt-design "品牌策略" --style "elegant luxury"         # → golden-luxury + elegant-serif + gold-trim
ppt-design "ESG报告" --style "calm nature"             # → sage-calm + humanist-sans + circle-accent
ppt-design "创业路演" --style "bold startup vibrant"   # → royal-purple + bold-sans + gradient-bar
```

### 35 种 mood 关键词

除 5 种预置主题外，支持 35 种 mood 关键词自动匹配：

professional, tech, dark, warm, elegant, luxury, vibrant, startup, nature, calm, minimal, bold, fresh, industrial, fintech, health, education, sustainability, creative, international, cream, frosted, mckinsey, consulting, pastel, retro, government, legal, pharma, realestate, automotive, aviation, energy, telecom, logistics

### 预置主题

---

## 🏆 设计质量升级 — 28 项专业级提升

v0.9.0 引入 VI Build 模式 + 28 项设计质量升级：

### Tier 1 — 基础视觉（10 项）

| # | 升级 | 说明 |
|---|------|------|
| 1.1 | **布局引擎** | `LayoutEngine` + `Rect` + `ContentLayout`，统一坐标计算 |
| 1.2 | **排版比例尺** | `TypeScale` 数据类，密度/模式自适应字号比例 |
| 1.3 | **OKLCH 色彩深度** | 色彩感知均匀的 9 级色阶 + 5 级 alpha 层级 |
| 1.4 | **渐变叠层** | `add_gradient_overlay()` 取代纯色遮罩，封面/CTA 更柔和 |
| 1.5 | **5 级阴影层级** | `ELEVATION_SCALE` 从 subtle→floating，层次分明 |
| 1.6 | **品牌条智能省略** | 约 1/3 页面省略品牌色条，避免视觉疲劳 |
| 1.7 | **图片调色** | `grade_image_to_palette()` 将图片色调统一到品牌色系 |
| 1.8 | **卡片升级** | 标题 20pt / 正文 14pt，`featured` 卡片带渐变条 |
| 1.9 | **暗色模式修正** | OKLCH 亮度判定（0.299R+0.587G+0.114B），告别误判 |
| 1.10 | **代码块重设计** | 始终深色背景 #1E293B + 独立语言徽章 + muted 文字色 |

### Tier 2 — 排版增强（6 项）

| # | 升级 | 说明 |
|---|------|------|
| 2.1 | **CJK 字体配对** | 12 种中英文字体组合（如 Space Grotesk + Microsoft YaHei），自动回退 |
| 2.2 | **自适应边距** | presenting / reading / balanced 三种密度对应不同边距 |
| 2.3 | **徽章系统** | `add_badge()` — 全大写、3 种变体（default/solid/outline）、CJK 宽度适配 |
| 2.4 | **章节分隔页** | `goal="section"` → 超大编号 + 标题 + 渐变线，清晰分章 |
| 2.5 | **装饰渲染器** | `DecorationRenderer` 统一渲染 10 种装饰风格 |
| 2.6 | **布局变体消费** | `render_slide()` 读取 content_margin_left / title_alignment / decoration_style |

### Tier 3 — 高级视觉（7 项）

| # | 升级 | 说明 |
|---|------|------|
| 3.1 | **噪点纹理** | 每套 PPT 独立种子的高斯噪点，增添质感 |
| 3.2 | **进度条** | 细条底部进度指示，替代简单页码 |
| 3.3 | **圆角系统** | 4 级圆角（sm=4 / md=8 / lg=12 / pill=50）+ `add_rounded_rect()` |
| 3.4 | **渐变线** | `add_gradient_line()` — alpha 渐隐，标题下划线更精致 |
| 3.5 | **图片遮罩** | `add_masked_image()` — 圆角矩形框 + 0.15" 内边距 |
| 3.6 | **双栏要点** | 6+ 条要点自动分两栏 + 垂直分隔线 |
| 3.7 | **Hero 4 变体** | gradient / split-left / bottom-fade / asymmetric 4 种封面布局 |

---

## License

MIT
