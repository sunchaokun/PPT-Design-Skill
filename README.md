<div align="center">

# PPT Design Skill

**一句话生成专业级 .pptx 演示文稿**

三模式引擎 · 叙事驱动 · 品牌合规 · AI 配图 · **40,000+ 风格组合** · **28 项设计质量升级** · **Build Script 精确控制**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![pptx](https://img.shields.io/badge/python--pptx-1.0.2-green.svg)](https://pypi.org/project/python-pptx/)

适配 OpenCode · Claude Code · Codex · Cursor

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
| **三模式引擎** | FreeStyle 快速生成 + Enterprise 品牌合规 + Build Script 逐页精确控制 |
| **叙事引擎** | 3 种策略（YC Seed Deck / Product Demo / Sales Pitch）+ Duarte Sparkline 情绪弧线 |
| **40,000+ 风格组合** | 25 色彩方案 × 20 字体搭配 × 10 装饰风格 × 8 布局变体 |
| **自然语言风格** | 描述风格即生成：`--style "warm fintech"` / `--style "dark cyberpunk"` |
| **28 项设计质量升级** | OKLCH 色彩深度 · 阴影层级 · 渐变叠层 · 进度条 · 圆角系统 · CJK 字体配对 · 噪点纹理 · 双栏要点 · 4 种 Hero 布局 · 章节分隔页 · 徽章系统 · 渐变线 · 图片遮罩 · 装饰渲染器 · 代码块重设计 · 卡片升级 · 自适应边距 · 排版比例尺 |
| **10 种图形引擎** | 流程图/漏斗/时间线/SWOT/矩阵/循环/表格/层级/金字塔/韦恩 |
| **Build Script 模式** | 10 种页面模板 + Design Token 系统 + 生成后自动检查，交付级质量 |
| **品牌视觉设计** | 品牌色自动映射：背景色、accent 竖条、品牌色标题/正文、LOGO 定位 |
| **页面修订** | `--pages` 增删改查：删除/交换/移动/插入，内容完整保留 |
| **版本管理** | v1 → v2 → v3 自动编号，meta.json 记录每页 goal/title |
| **python-pptx 直出** | 完全可编辑 .pptx，356x 快于 HTML→截图方案 |
| **12 母版布局** | 13.333"×7.5" 16:9 精确坐标 |
| **AI 智能配图** | Seedream / GPT Image / DALL-E / Wanx — 4 种生成引擎 + Kimi 增强 |
| **动画系统** | 12 种切换 + 10 种入场动画，motion 1-10 智能映射 |
| **代码块/练习** | 深色代码块 + 语言徽章 + 练习徽章 + 步骤列表 — 教育场景标配 |
| **CJK 字体** | 12 种 CJK 配对自动回退（Microsoft YaHei / STSong / SimHei 等） |
| **QA 门禁** | 5 项自动质量检查 + `--review` 方案确认 |

---

## 🚀 快速开始

### 安装

```bash
git clone https://github.com/sunchaokun/PPT-Design-Skill.git
cd PPT-Design-Skill

# 一键安装 — 自动检测 AI 平台 + 安装 skill + pip 依赖
python install.py

# 手动安装
pip install -e .
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

### Enterprise — 品牌合规 + 版本管理

```bash
# 1. 创建项目目录
mkdir my-pitch && cd my-pitch

# 2. 放入品牌资产（均可选）
#    template.pptx / brand.json / content.json / logo.png / images/

# 3. 生成
ppt-design "AI Platform" --project . --density 6 --motion 5

# 4. 修订页面
ppt-design "" --project . --pages "-3,2<>5"

# 5. 查看历史
ppt-design "" --project . --history
```

### Build Script — 逐页精确控制（交付级质量）

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

> 10 种页面模板 + Design Token 系统 + 生成后检查脚本，详见 [使用手册 §5](docs/usage-guide.md#5-build-script-模式--精确控制)

---

## 🏗️ 三模式架构

### FreeStyle Pipeline

```
用户输入 → 叙事规划 → 设计决策 → 内容生成 → PPT渲染 → .pptx
```

适用：快速探索、风格实验、个人演示

### Enterprise Pipeline

```
项目目录 → 资产扫描 → 品牌合并 → 内容解析 → 视觉设计 → 渲染+动画 → 版本管理 → .pptx
               ↓          ↓          ↓          ↓
          template   brand.json  content.json  accent bar
          logo.png   颜色/字体    diagram      品牌色覆盖
          images/    LOGO定位     code/exercise  页脚/水印
```

适用：企业合规、团队协作、多版本迭代

### Build Script 模式

```
build.py → Design Token → 页面模板 → python-pptx → .pptx → check.py 验证
              ↓              ↓            ↓
          颜色/字体      10种模板     逐元素精确控制
          一键切换主题   x/y/w/h     run-level 字体
```

适用：**最终交付、精确控制、质量保证**

> **推荐工作流**：FreeStyle 原型 → Enterprise 内容填充 → Build Script 精细交付

### 项目目录结构

```
my-project/
├── template.pptx      # 可选：品牌模板
├── brand.json         # 可选：品牌规范
├── content.json       # 可选：真实内容
├── logo.png           # 可选：公司 LOGO
├── images/            # 可选：图片池
└── output/            # 自动生成
    ├── v1/presentation.pptx + meta.json
    └── v2/presentation.pptx + meta.json
```

---

## 📋 content.json — 精确控制内容

```json
{
  "meta": {"title": "Acme Corp — Series B Pitch"},
  "slides": [
    {"goal": "hook", "title": "The Future of AI", "subtitle": "Acme Corp", "image": "images/hero.png"},
    {"goal": "problem", "title": "The Problem", "bullets": ["75% fail", "Data silos"]},
    {"goal": "solution", "title": "Our Solution", "bullets": ["Unified gateway"], "image": "images/product.png"},
    {"goal": "features", "title": "Key Features", "cards": [
      {"title": "Fast", "text": "Sub-100ms"},
      {"title": "Secure", "text": "SOC2+HIPAA"}
    ]},
    {"goal": "market", "title": "Market", "diagram": {"type": "funnel", "data": {"items": [{"text": "TAM $120B"}, {"text": "SOM $8B"}]}}},
    {"goal": "code_demo", "title": "Quick Start", "code": {"language": "python", "source": "from acme import AIPlatform\nplatform = AIPlatform(key='x')"}},
    {"goal": "exercise", "title": "Try It", "exercise": {"instructions": "Deploy in 5 min", "duration": "5 min", "steps": ["Sign up", "Deploy"]}},
    {"goal": "cta", "title": "Join Us", "subtitle": "contact@acme.ai"}
  ]
}
```

> 完整字段说明见 [使用手册](docs/usage-guide.md#5-contentjson-内容格式)

---

## 🎨 brand.json — 品牌视觉规范

```json
{
  "colors": {
    "primary": "#2563EB",
    "accent": "#F97316",
    "foreground": "#1A1A2E",
    "muted-foreground": "#94A3B8",
    "background": "#FFFFFF",
    "muted": "#F1F5F9"
  },
  "fonts": {"heading": "Calibri", "body": "Calibri"},
  "logo": {"position": "top_right", "width_inches": 1.2, "skip_slides": ["hook"]},
  "footer": {
    "show_page_number": true,
    "page_number_format": "第 {n} 页 / 共 {total} 页"
  }
}
```

> 完整字段说明见 [使用手册](docs/usage-guide.md#6-brandjson-品牌格式)

---

## ✏️ 页面修订 — 增删改查

```bash
# 删除第 3 页
ppt-design "" --project . --pages "-3"

# 交换第 2 和第 5 页
ppt-design "" --project . --pages "2<>5"

# 将第 10 页移到第 3 页后
ppt-design "" --project . --pages "10>3"

# 在第 6 页后插入空白页
ppt-design "" --project . --pages "+6"

# 组合操作
ppt-design "" --project . --pages "-3,2<>5,10>3,+6"
```

> 完整语法见 [使用手册](docs/usage-guide.md#7-页面修订语法)

---

## 🖼️ 图片引擎

| 引擎 | 类型 | CLI | 默认模型 |
|------|------|-----|---------|
| `placeholder` | 占位符 | 默认 | — |
| `search` | 搜索下载 | `--image-mode search` | — |
| `seedream` | AI 生成 | `--llm-provider seedream` | `doubao-seedream-5-0-260128` |
| `gpt-image` | AI 生成 | `--llm-provider gpt-image` | `gpt-image-1` |
| `dalle` | AI 生成 | `--llm-provider dalle` | `dall-e-3` |
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
| 🎨 色彩方案 | 25 | ocean-blue, cyber-neon, golden-luxury, wine-burgundy, monochrome-dark... |
| ✏️ 字体搭配 | 20 | modern-sans, serif-editorial, tech-mono, elegant-serif, contrast-mix... |
| 🖌️ 装饰风格 | 10 | accent-bar, neon-lines, gold-trim, gradient-bar, circle-accent, sidebar-nav... |
| 📐 布局变体 | 8 | standard, centered, sidebar-left, grid-2x2, wide-cards, full-width... |

**25 × 20 × 10 × 8 = 40,000 种组合**

### 自然语言风格

```bash
ppt-design "融资路演" --style "warm fintech"          # → ocean-blue + clean-corporate + accent-bar
ppt-design "产品发布" --style "dark cyberpunk"         # → cyber-neon + tech-mono + neon-lines
ppt-design "品牌策略" --style "elegant luxury"         # → golden-luxury + elegant-serif + gold-trim
ppt-design "ESG报告" --style "calm nature"             # → sage-calm + humanist-sans + circle-accent
ppt-design "创业路演" --style "bold startup vibrant"   # → royal-purple + bold-sans + gradient-bar
```

### 预置主题

| 主题 | 色彩 | 字体 | 装饰 | 布局 |
|------|------|------|------|------|
| Professional | midnight-navy | clean-corporate | accent-bar | sidebar-left |
| Dark Tech | cyber-neon | tech-mono | neon-lines | wide-cards |
| Warm Elegant | golden-luxury | serif-editorial | gold-trim | centered |
| Vibrant Startup | neon-gradient | bold-sans | gradient-bar | grid-2x2 |
| Nature Calm | forest-green | humanist-sans | circle-accent | sidebar-left |

---

## 🏆 设计质量升级 — 28 项专业级提升

v0.7.0 引入 28 项设计质量升级，分三个层级全面提升 PPT 视觉品质：

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

## 📁 项目结构

```
PPT-Design-Skill/
├── pyproject.toml
├── install.py                        # 一键安装器
├── SKILL.md                          # AI skill 定义
├── AGENTS.md                         # 项目指令
├── docs/
│   ├── README_EN.md                  # 英文 README
│   ├── usage-guide.md                # 完整使用手册（含 Build Script 模式）
│   └── showcase/                     # 案例截图
├── examples/                         # 案例 PPT
│   ├── showcase-professional.pptx
│   ├── showcase-dark-tech.pptx
│   ├── showcase-warm-elegant.pptx
│   ├── showcase-vibrant-startup.pptx
│   └── showcase-nature-calm.pptx
├── src/ppt_pro_max/
│   ├── __init__.py                   # generate_ppt() API
│   ├── cli.py                        # ppt-design CLI
│   ├── enterprise/                   # Enterprise Pipeline
│   │   ├── pipeline.py               # 主编排器
│   │   ├── precision_renderer.py     # 统一渲染器（8 种 goal 布局 + 28 项设计升级）
│   │   ├── brand_spec.py             # 品牌规范
│   │   ├── content_parser.py         # 内容解析（content.json + README.md）
│   │   ├── image_matcher.py          # 图片匹配 + 尺寸分类 + AI 提示词
│   │   ├── proposal_generator.py     # 方案预览（2-3 种风格）
│   │   ├── slide_extractor.py        # PPT 内容提取（美化模式）
│   │   ├── component_library.py      # 组件库（SQLite 索引 + 去重）
│   │   ├── component_renderer.py     # 组件渲染桥接
│   │   ├── page_revision.py          # 页面修订引擎
│   │   ├── density_profile.py        # 密度配置
│   │   └── ...
│   ├── renderer/
│   │   ├── ppt_renderer.py           # FreeStyle 渲染器
│   │   ├── diagram_engine.py         # 10 种图形引擎
│   │   ├── diagram/                  # 图形实现
│   │   │   ├── flowchart.py / funnel.py / timeline.py / swot.py
│   │   │   ├── matrix.py / cycle.py / table.py
│   │   │   ├── hierarchy.py / pyramid.py / venn.py
│   │   │   └── connector_router.py / text_measurer.py / ...
│   │   ├── animation.py              # 12 切换 + 10 入场
│   │   ├── theme_composer.py         # 40,000+ 风格组合 + 35 种 mood
│   │   ├── typography.py             # 排版比例尺（TypeScale）
│   │   ├── color_system.py           # OKLCH 色彩深度 + alpha 层级
│   │   ├── elevation.py              # 5 级阴影层级
│   │   ├── layout_engine.py          # 布局引擎 + 自适应边距
│   │   ├── image_processor.py        # 图片调色 + 噪点纹理
│   │   ├── decoration_renderer.py    # 10 种装饰风格渲染
│   │   ├── visual_effects.py         # 渐变/发光/字间距
│   │   ├── image_fetcher.py          # 5 引擎 + 缓存
│   │   └── ...
│   ├── planner/story_planner.py      # 叙事规划
│   ├── decider/design_decider.py     # 设计决策
│   └── content/content_generator.py  # 内容生成
├── tests/                            # 824 个测试
└── e2e-test-project/                 # E2E 测试项目
```

## License

MIT
