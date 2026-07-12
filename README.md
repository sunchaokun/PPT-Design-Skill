<div align="center">

# PPT Design Skill

**一句话生成专业级 .pptx 演示文稿**

双模式引擎 · 叙事驱动 · 品牌合规 · AI 配图 · **40,000+ 风格组合**

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
| **双模式引擎** | FreeStyle 快速生成 + Enterprise 品牌合规 + 版本管理 + 页面修订 |
| **叙事引擎** | 3 种策略（YC Seed Deck / Product Demo / Sales Pitch）+ Duarte Sparkline 情绪弧线 |
| **40,000+ 风格组合** | 25 色彩方案 × 20 字体搭配 × 10 装饰风格 × 8 布局变体 |
| **自然语言风格** | 描述风格即生成：`--style "warm fintech"` / `--style "dark cyberpunk"` |
| **10 种图形引擎** | 流程图/漏斗/时间线/SWOT/矩阵/循环/表格/层级/金字塔/韦恩 |
| **品牌视觉设计** | 品牌色自动映射：背景色、accent 竖条、品牌色标题/正文、LOGO 定位 |
| **页面修订** | `--pages` 增删改查：删除/交换/移动/插入，内容完整保留 |
| **版本管理** | v1 → v2 → v3 自动编号，meta.json 记录每页 goal/title |
| **python-pptx 直出** | 完全可编辑 .pptx，356x 快于 HTML→截图方案 |
| **12 母版布局** | 13.333"×7.5" 16:9 精确坐标 |
| **AI 智能配图** | Seedream / GPT Image / DALL-E / Wanx — 4 种生成引擎 + Kimi 增强 |
| **动画系统** | 12 种切换 + 10 种入场动画，motion 1-10 智能映射 |
| **代码块/练习** | 深色代码块 + 练习徽章 + 步骤列表 — 教育场景标配 |
| **CJK 字体** | 东亚字体自动回退链（Microsoft YaHei / STSong） |
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

---

## 🏗️ 双模式架构

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

## 📁 项目结构

```
PPT-Design-Skill/
├── pyproject.toml
├── install.py                        # 一键安装器
├── SKILL.md                          # AI skill 定义
├── AGENTS.md                         # 项目指令
├── docs/
│   ├── README_EN.md                  # 英文 README
│   ├── usage-guide.md                # 完整使用手册
│   └── showcase/                     # 案例截图
├── src/ppt_pro_max/
│   ├── __init__.py                   # generate_ppt() API
│   ├── cli.py                        # ppt-design CLI
│   ├── enterprise/                   # Enterprise Pipeline
│   │   ├── pipeline.py               # 主编排器
│   │   ├── enterprise_renderer.py    # 企业渲染器
│   │   ├── brand_spec.py             # 品牌规范
│   │   ├── content_parser.py         # 内容解析
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
│   │   ├── theme_composer.py         # 40,000+ 风格组合
│   │   ├── image_fetcher.py          # 5 引擎 + 缓存
│   │   └── ...
│   ├── planner/story_planner.py      # 叙事规划
│   ├── decider/design_decider.py     # 设计决策
│   └── content/content_generator.py  # 内容生成
├── tests/                            # 412 个测试
└── e2e-test-project/                 # E2E 测试项目
```

## License

MIT
