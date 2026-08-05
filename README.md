<div align="center">

# PPT Design Skill

> 给 OpenCode / Claude Code / Codex / Cursor 的 PPT 生成技能

**一句话 → 专业 .pptx · 40,000+ 风格 · AI 配图 · 完全可编辑**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![pptx](https://img.shields.io/badge/python--pptx-1.0.2-green.svg)](https://pypi.org/project/python-pptx/)

在 AI 编码工具中输入：`生成一份AI融资路演PPT` → 自动调用 skill → 输出可编辑 .pptx

| FreeStyle 自由模式 | Build 设计师模式 | VI Build 企业模式 |
|:---:|:---:|:---:|
| 一句话出PPT | **像素级控制 + 方案对比** | **基于企业模板 VI 精确生成** |
| 30秒快速生成 | **python-pptx 精确构建** | **保留框架页 + build_helpers** |

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

# 一键安装 — 自动检测平台 + 安装 skill + pip 注册 Python 包 + 依赖
python install.py                     # 自动检测
python install.py --platform opencode # 指定平台
```

安装后 `ppt_pro_max` 注册为全局 Python 包，任何项目下可直接 `from ppt_pro_max.build_helpers import *`。

支持 13 个平台：OpenCode · Claude Code · Codex · Cursor · Windsurf · Roo Code · Gemini · Trae · Continue · Droid · KiloCode · Augment · Copilot

### 作为 Python 包使用

```bash
pip install git+https://github.com/sunchaokun/PPT-Design-Skill.git
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
pip install --force-reinstall git+https://github.com/sunchaokun/PPT-Design-Skill.git
```

### 在 AI 编码工具中使用

安装后，在 OpenCode / Claude Code / Codex 中直接输入：

```
生成一份AI融资路演PPT，dark cyberpunk风格
```

AI 会自动加载 skill 并生成 .pptx 文件。

### FreeStyle — 一句话生成

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
python -m ppt_pro_max.analyze_template template.pptx > analysis.txt
# 将 analysis.txt 交给 LLM 生成 build.py，然后：
python build.py
```

### Build Script — 逐页精确控制

```python
from ppt_pro_max.build_helpers import *

prs = Presentation()
s = add_slide(prs)
hero_slide(s, '标题', '副标题', C=C, typo=TYPOGRAPHY['mckinsey'])
# ... 每个元素精确控制 x, y, w, h, font, size, color
prs.save("output/presentation.pptx")
```

---

## 🔥 核心特性

| 特性 | 说明 |
|------|------|
| **三模式引擎** | FreeStyle 快速生成 + Build Script 逐页精确控制 + VI Build 企业模板合规 |
| **40,000+ 风格组合** | 30 色彩方案 × 25 字体 × 15 装饰 × 12 布局，自然语言指定 `--style` |
| **AI 智能配图** | Seedream / GPT Image / DALL-E / Gemini / Wanx 5 种引擎 + Kimi 增强 |
| **python-pptx 直出** | 完全可编辑 .pptx，356x 快于 HTML→截图方案 |
| **10 种图形引擎** | 流程图 / 漏斗 / 时间线 / SWOT / 矩阵 / 循环 / 表格 / 层级 / 金字塔 / 韦恩 |
| **动画系统** | 12 种切换 + 10 入场 + 8 退场 + 8 强调 + Morph，motion 1-10 映射 |
| **CJK 字体** | 12 种中英文字体配对自动回退 |
| **5,500+ 组件库** | SmartArt/GroupShape 模板，SQLite 索引，按类别/节点数匹配 |

---

## 🏗️ 三模式架构

| | **FreeStyle** | **Build Script** | **VI Build** |
|---|---|---|---|
| **场景** | 快速探索、原型 | 交付级精确控制 | 企业 VI 合规 |
| **触发** | 默认 | `"build mode"` / `"像素级"` | 提供 template.pptx |
| **内容** | AI 自动生成 | 手写 build.py | LLM 读模板生成 build.py |
| **质量** | ★★★ | ★★★★★ | ★★★★★ |
| **方案** | 3 种风格预览 | 3 种结构化方案 | 3 种布局方案（同 VI Token） |

> **推荐工作流**：FreeStyle 原型 → Build / VI Build 精细交付

---

## 🎨 设计系统

**自然语言风格** — 描述即生成：

```bash
ppt-design "融资路演" --style "warm fintech"       # → ocean-blue + clean-corporate + accent-bar
ppt-design "产品发布" --style "dark cyberpunk"      # → cyber-neon + tech-mono + neon-lines
ppt-design "品牌策略" --style "elegant luxury"      # → golden-luxury + elegant-serif + gold-trim
ppt-design "山水诗词" --style "水墨"                # → ink-wash 调色板 + 楷书 + 笔触装饰
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

叠加 ui-ux-pro-max（192 色彩方案 · 84 风格 · 74 字体 · 161 反模式）可达 200,000+

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

## License

MIT
