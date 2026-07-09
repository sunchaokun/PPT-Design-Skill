# PPT-Design-Skill

**一句话生成专业级 .pptx 演示文稿** — 叙事驱动、设计智能、文本完全可编辑、图片自动配图。

适配 OpenCode、Claude Code、Codex 等主流 AI 编程平台。

[English](#english) | 中文

---

## 案例展示

| # | 案例名称 | 策略 | 主题 | 页数 | 预览 |
|---|----------|------|------|------|------|
| 1 | AI 创业融资路演 | YC Seed Deck | Dark Tech | 10 | [下载](examples/01-ai-startup-investor-pitch.pptx) |
| 2 | SaaS 产品展示 | Product Demo | Professional Modern | 8 | [下载](examples/02-saas-product-demo.pptx) |
| 3 | 高端销售方案 | Sales Pitch | Warm Elegant | 7 | [下载](examples/03-premium-sales-pitch.pptx) |
| 4 | 科技创业发布会 | YC Seed Deck (12p) | Vibrant Startup | 12 | [下载](examples/04-tech-startup-launch.pptx) |
| 5 | 可持续发展报告 | YC Seed Deck | Nature Calm | 10 | [下载](examples/05-sustainability-report.pptx) |
| 6 | 深科技产品深潜 | Product Demo | Dark Tech | 8 | [下载](examples/06-deep-tech-product-dive.pptx) |

## 核心特性

- **叙事引擎** — 3 种演示策略（YC Seed Deck / Product Demo / Sales Pitch），自动规划页面结构和情绪弧线（Duarte Sparkline）
- **设计智能** — 复用 ui-ux-pro-max 的 5100+ 条设计知识库，上下文感知的布局/色彩/字体决策
- **python-pptx 直出** — 输出完全可编辑的 .pptx 文件，356x 快于 HTML→截图方案
- **12 母版布局** — 精确英寸坐标（13.333"×7.5" 16:9），覆盖 95% 演示场景
- **9 种文案公式** — PAS / FAB / AIDA / Social Proof / Cost of Inaction / Proof Stack...
- **5 套预置主题** — Professional Modern / Dark Tech / Warm Elegant / Vibrant Startup / Nature Calm
- **智能配图** — 4 种图片模式：占位符 / Unsplash搜索 / AI生成(DALL-E/通义万相) / Kimi K2.6增强搜索
- **设计拨盘** — variance / motion / density 三维调控
- **CJK 字体** — 东亚字体自动回退链（Microsoft YaHei / STSong）
- **QA 门禁** — 5 项自动质量检查（页数/标题/字号/一致性/占位符）

## 快速开始

### 安装

```bash
pip install -e .
# 可选：安装搜索引擎支持
pip install -e ".[search]"
```

### CLI

```bash
# 一句话生成 PPT
ppt-design "AI产品融资路演"

# 指定策略和主题
ppt-design "SaaS产品展示" --strategy "Product Demo" --theme "dark-tech"

# 带图片搜索
ppt-design "融资路演" --image-mode search --unsplash-key YOUR_KEY

# Kimi K2.6 增强搜索关键词 → 自动下载高质量图片
ppt-design "融资路演" --image-mode enhance --llm-provider kimi --llm-api-key sk-xxx

# AI 生成图片（DALL-E / 通义万相）
ppt-design "产品介绍" --image-mode generate --llm-provider dalle --llm-api-key sk-xxx

# 从 JSON 文件读取真实内容
ppt-design "融资路演" --content pitch-data.json

# 设计拨盘
ppt-design "融资路演" --variance 8 --motion 6 --density 7

# 仅输出设计决策（调试）
ppt-design "融资路演" --dry-run

# 持久化设计系统
ppt-design "融资路演" --persist
```

### Python API

```python
from ppt_pro_max import generate_ppt

# 最简用法
result = generate_ppt("AI产品融资路演")
print(f"生成: {result['output_path']}, {result['page_count']}页")

# 完整配置
result = generate_ppt(
    query="AI产品融资路演",
    strategy="YC Seed Deck",
    theme="dark-tech",
    slides=12,
    image_mode="enhance",
    image_config={
        "llm_provider": "kimi",
        "llm_api_key": "sk-xxx",
    },
    variance=7,
    motion=5,
    density=6,
    persist=True,
    output="my-pitch.pptx",
)
```

### 内容 JSON 格式

```json
{
  "company": "Acme AI",
  "product": "AI Marketing Platform",
  "tagline": "Your AI marketing team. Always on.",
  "metrics": {"users": "10K+", "retention": "95%", "growth": "3x", "arr": "$2M"},
  "pain_points": [
    {"title": "Content Overload", "desc": "Need 10x content with same headcount"},
    {"title": "Tool Fatigue", "desc": "15+ tools that don't talk to each other"}
  ],
  "chart_data": {
    "mrr": {"labels": ["Sep","Oct","Nov","Dec"], "values": [5,12,28,45]}
  }
}
```

## 四层架构

```
用户输入 → Phase 1: 叙事规划 → Phase 2: 设计决策 → Phase 3: 内容生成 → Phase 4: PPT渲染 → .pptx
                 ↓                  ↓                  ↓                  ↓
          story_planner      design_decider     content_generator     ppt_renderer
          (策略+情绪弧线)    (布局+色彩+字体)    (文案公式+图片)       (python-pptx直出)
```

| 层 | 模块 | 职责 |
|----|------|------|
| 叙事层 | `planner/story_planner.py` | 策略选择 → 页面结构 → 情绪弧线 → Duarte Sparkline |
| 设计层 | `decider/design_decider.py` | 布局选择 → 色彩处理 → 排版规格 → 图表类型 → 过渡选择 |
| 内容层 | `content/content_generator.py` | 文案公式 → 数据配置 → 图片关键词 → 模板变量填充 |
| 表达层 | `renderer/ppt_renderer.py` | 主题映射 → 12母版布局 → 图表渲染 → 图片插入 → 动画过渡 → QA检查 |

## 图片模式

| 模式 | 说明 | CLI 参数 | 需要 |
|------|------|----------|------|
| `placeholder` | 灰色占位框 + 关键词提示 | 默认 | 无 |
| `search` | Unsplash/Pexels 搜索下载 | `--image-mode search` | API key |
| `generate` | DALL-E / 通义万相 AI 生成 | `--image-mode generate` | API key |
| `enhance` | Kimi K2.6 优化关键词 → 搜索下载 | `--image-mode enhance` | Moonshot API key |

Kimi K2.6 作为文本模型，会根据每页的 goal 和 emotion 生成更精准的 Unsplash 搜索关键词，然后下载高质量图片插入 PPT。

## 5 套预置主题

| 主题 | 风格 | 背景 | 主色 | 适用场景 |
|------|------|------|------|----------|
| Professional Modern | 现代商务 | 白色 | #2563EB | 企业汇报、产品演示 |
| Dark Tech | 深色科技 | #0F172A | #6366F1 | 科技创业、技术路演 |
| Warm Elegant | 温暖雅致 | #FFFBEB | #92400E | 高端销售、品牌展示 |
| Vibrant Startup | 活力创业 | #F8FAFC | #7C3AED | 创业路演、融资展示 |
| Nature Calm | 自然沉稳 | #ECFDF5 | #059669 | ESG报告、可持续发展 |

## 与 ui-ux-pro-max 的关系

**独立仓库 + 依赖引用**模式：

- **调用** ui-ux-pro-max 的搜索引擎和设计知识库（5100+ 条 CSV）
- **不修改** ui-ux-pro-max 的任何代码
- **新增** PPT 专属代码和数据
- 通过**适配层**隔离上游 API 变更

## 项目结构

```
PPT-Design-Skill/
├── pyproject.toml
├── SKILL.md                          # AI skill 定义
├── src/ppt_pro_max/
│   ├── __init__.py                   # generate_ppt() API
│   ├── cli.py                        # ppt-design CLI
│   ├── adapters/                     # 适配层
│   │   ├── ui_ux_adapter.py          # ui-ux-pro-max 适配
│   │   └── slide_search_adapter.py   # slide_search 适配
│   ├── planner/story_planner.py      # Phase 1: 叙事规划
│   ├── decider/design_decider.py     # Phase 2: 设计决策
│   ├── content/content_generator.py  # Phase 3: 内容生成
│   ├── renderer/
│   │   ├── ppt_renderer.py           # Phase 4: PPT 渲染
│   │   ├── theme_mapper.py           # 主题映射 + CJK 字体
│   │   ├── layout_registry.py        # 12 母版布局
│   │   ├── chart_builder.py          # 图表构建器
│   │   ├── image_fetcher.py          # 图片获取（4种模式）
│   │   └── effects.py               # 阴影/辉光/渐变
│   └── qa/qa_gates.py               # 5项质量检查
├── data/ppt/                         # PPT 专属数据
├── examples/                         # 6个案例 PPT
└── tests/                            # 42 个测试
```

## License

MIT

---

<a id="english"></a>

# PPT-Design-Skill

**Generate professional .pptx presentations from a single sentence** — narrative-driven, design-intelligent, fully editable text, auto-fetched images.

Compatible with OpenCode, Claude Code, Codex and other AI coding platforms.

## Showcase

| # | Example | Strategy | Theme | Pages | Download |
|---|---------|----------|-------|-------|----------|
| 1 | AI Startup Investor Pitch | YC Seed Deck | Dark Tech | 10 | [Download](examples/01-ai-startup-investor-pitch.pptx) |
| 2 | SaaS Product Demo | Product Demo | Professional Modern | 8 | [Download](examples/02-saas-product-demo.pptx) |
| 3 | Premium Sales Pitch | Sales Pitch | Warm Elegant | 7 | [Download](examples/03-premium-sales-pitch.pptx) |
| 4 | Tech Startup Launch | YC Seed Deck (12p) | Vibrant Startup | 12 | [Download](examples/04-tech-startup-launch.pptx) |
| 5 | Sustainability Report | YC Seed Deck | Nature Calm | 10 | [Download](examples/05-sustainability-report.pptx) |
| 6 | Deep Tech Product Dive | Product Demo | Dark Tech | 8 | [Download](examples/06-deep-tech-product-dive.pptx) |

## Features

- **Narrative Engine** — 3 presentation strategies (YC Seed Deck / Product Demo / Sales Pitch) with emotion arcs (Duarte Sparkline)
- **Design Intelligence** — Reuses ui-ux-pro-max's 5100+ design knowledge base for context-aware layout/color/typography decisions
- **python-pptx Direct** — Outputs fully editable .pptx files, 356x faster than HTML→screenshot approaches
- **12 Master Layouts** — Precise inch coordinates (13.333"×7.5" 16:9) covering 95% of presentation scenarios
- **9 Copy Formulas** — PAS / FAB / AIDA / Social Proof / Cost of Inaction / Proof Stack...
- **5 Preset Themes** — Professional Modern / Dark Tech / Warm Elegant / Vibrant Startup / Nature Calm
- **Smart Images** — 4 image modes: placeholder / Unsplash search / AI generate (DALL-E/Wanx) / Kimi K2.6 enhanced search
- **Design Dials** — variance / motion / density 3-axis control
- **CJK Fonts** — East Asian font fallback chain (Microsoft YaHei / STSong)
- **QA Gates** — 5 automated quality checks (page count / titles / font sizes / consistency / placeholders)

## Quick Start

### Install

```bash
pip install -e .
# Optional: search engine support
pip install -e ".[search]"
```

### CLI

```bash
# Generate PPT from a single sentence
ppt-design "AI startup investor pitch"

# Specify strategy and theme
ppt-design "SaaS product demo" --strategy "Product Demo" --theme "dark-tech"

# Fetch images from Unsplash
ppt-design "investor pitch" --image-mode search --unsplash-key YOUR_KEY

# Kimi K2.6 enhances search keywords → downloads high-quality images
ppt-design "investor pitch" --image-mode enhance --llm-provider kimi --llm-api-key sk-xxx

# AI-generated images (DALL-E / Wanx)
ppt-design "product intro" --image-mode generate --llm-provider dalle --llm-api-key sk-xxx

# Load real content from JSON
ppt-design "investor pitch" --content pitch-data.json

# Design dials
ppt-design "investor pitch" --variance 8 --motion 6 --density 7

# Dry-run (design decisions only)
ppt-design "investor pitch" --dry-run

# Persist design system
ppt-design "investor pitch" --persist
```

### Python API

```python
from ppt_pro_max import generate_ppt

# Minimal usage
result = generate_ppt("AI startup investor pitch")
print(f"Generated: {result['output_path']}, {result['page_count']} pages")

# Full configuration
result = generate_ppt(
    query="AI startup investor pitch",
    strategy="YC Seed Deck",
    theme="dark-tech",
    slides=12,
    image_mode="enhance",
    image_config={
        "llm_provider": "kimi",
        "llm_api_key": "sk-xxx",
    },
    variance=7,
    motion=5,
    density=6,
    persist=True,
    output="my-pitch.pptx",
)
```

### Content JSON Format

```json
{
  "company": "Acme AI",
  "product": "AI Marketing Platform",
  "tagline": "Your AI marketing team. Always on.",
  "metrics": {"users": "10K+", "retention": "95%", "growth": "3x", "arr": "$2M"},
  "pain_points": [
    {"title": "Content Overload", "desc": "Need 10x content with same headcount"},
    {"title": "Tool Fatigue", "desc": "15+ tools that don't talk to each other"}
  ],
  "chart_data": {
    "mrr": {"labels": ["Sep","Oct","Nov","Dec"], "values": [5,12,28,45]}
  }
}
```

## 4-Layer Architecture

```
Input → Phase 1: Story Planning → Phase 2: Design Decisions → Phase 3: Content Generation → Phase 4: PPT Rendering → .pptx
              ↓                        ↓                          ↓                           ↓
       story_planner            design_decider            content_generator            ppt_renderer
       (strategy+emotion arc)  (layout+color+typography) (copy formulas+images)      (python-pptx direct)
```

| Layer | Module | Responsibility |
|-------|--------|----------------|
| Narrative | `planner/story_planner.py` | Strategy selection → Page structure → Emotion arc → Duarte Sparkline |
| Design | `decider/design_decider.py` | Layout selection → Color treatment → Typography specs → Chart type → Transitions |
| Content | `content/content_generator.py` | Copy formulas → Data configuration → Image keywords → Template variable filling |
| Expression | `renderer/ppt_renderer.py` | Theme mapping → 12 master layouts → Chart rendering → Image insertion → Animations → QA checks |

## Image Modes

| Mode | Description | CLI | Required |
|------|-------------|-----|----------|
| `placeholder` | Gray box with keywords hint | Default | None |
| `search` | Download from Unsplash/Pexels | `--image-mode search` | API key |
| `generate` | AI-generated via DALL-E / Wanx | `--image-mode generate` | API key |
| `enhance` | Kimi K2.6 enhances keywords → search | `--image-mode enhance` | Moonshot API key |

Kimi K2.6 is a text LLM — it generates better Unsplash search keywords based on each slide's goal and emotion, then downloads high-quality images for insertion.

## 5 Preset Themes

| Theme | Style | Background | Primary | Use Case |
|-------|-------|------------|---------|----------|
| Professional Modern | Clean corporate | White | #2563EB | Business reports, product demos |
| Dark Tech | Futuristic dark | #0F172A | #6366F1 | Tech startups, technical pitches |
| Warm Elegant | Warm premium | #FFFBEB | #92400E | Premium sales, brand showcases |
| Vibrant Startup | Bold energetic | #F8FAFC | #7C3AED | Startup pitches, fundraising |
| Nature Calm | Organic calm | #ECFDF5 | #059669 | ESG reports, sustainability |

## Relationship with ui-ux-pro-max

**Independent repo + dependency reference** model:

- **Calls** ui-ux-pro-max's search engine and design knowledge base (5100+ CSV entries)
- **Does NOT modify** any ui-ux-pro-max code
- **Adds** PPT-specific code and data
- **Adapter layer** isolates upstream API changes

## Project Structure

```
PPT-Design-Skill/
├── pyproject.toml
├── SKILL.md                          # AI skill definition
├── src/ppt_pro_max/
│   ├── __init__.py                   # generate_ppt() API
│   ├── cli.py                        # ppt-design CLI
│   ├── adapters/                     # Adapter layer
│   │   ├── ui_ux_adapter.py          # ui-ux-pro-max adapter
│   │   └── slide_search_adapter.py   # slide_search adapter
│   ├── planner/story_planner.py      # Phase 1: Story planning
│   ├── decider/design_decider.py     # Phase 2: Design decisions
│   ├── content/content_generator.py  # Phase 3: Content generation
│   ├── renderer/
│   │   ├── ppt_renderer.py           # Phase 4: PPT rendering
│   │   ├── theme_mapper.py           # Theme mapping + CJK fonts
│   │   ├── layout_registry.py        # 12 master layouts
│   │   ├── chart_builder.py          # Chart builder
│   │   ├── image_fetcher.py          # Image fetching (4 modes)
│   │   └── effects.py               # Shadow/glow/gradient
│   └── qa/qa_gates.py               # 5 quality checks
├── data/ppt/                         # PPT-specific data
├── examples/                         # 6 showcase PPTs
└── tests/                            # 42 tests
```

## License

MIT
