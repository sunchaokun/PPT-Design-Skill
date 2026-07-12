# PPT Pro Max 使用手册

> 版本 v0.3.0 | 完整功能参考

---

## 目录

1. [两种模式](#1-两种模式)
2. [快速开始](#2-快速开始)
3. [FreeStyle 模式](#3-freestyle-模式)
4. [Enterprise 模式](#4-enterprise-模式)
5. [content.json 内容格式](#5-contentjson-内容格式)
6. [brand.json 品牌格式](#6-brandjson-品牌格式)
7. [页面修订语法](#7-页面修订语法)
8. [10 种图形类型](#8-10-种图形类型)
9. [设计系统 — 40,000+ 风格组合](#9-设计系统--40000-风格组合)
10. [动画系统](#10-动画系统)
11. [密度控制](#11-密度控制)
12. [图片引擎](#12-图片引擎)
13. [Python API](#13-python-api)
14. [CLI 完整参数](#14-cli-完整参数)

---

## 1. 两种模式

PPT Pro Max 提供两种生成模式：

| | FreeStyle | Enterprise |
|---|---|---|
| **适用场景** | 快速生成、风格探索 | 企业合规、团队协作、版本管理 |
| **触发方式** | 默认（无 `--project`） | `--project <目录>` |
| **内容来源** | AI 自动生成 | content.json 精确控制 |
| **品牌规范** | 风格原子组合 | brand.json + 模板分析 |
| **版本管理** | 无 | v1 → v2 → v3 ... |
| **页面修订** | 不支持 | `--pages` 增删改查 |
| **确认流程** | 不支持 | `--review` 方案确认 |

---

## 2. 快速开始

### 安装

```bash
git clone https://github.com/sunchaokun/PPT-Design-Skill.git
cd PPT-Design-Skill
pip install -e .
```

### 一句话生成

```bash
# FreeStyle — 自动生成内容和设计
ppt-design "AI startup investor pitch"

# Enterprise — 从项目目录生成，品牌合规
ppt-design "AI Platform" --project ./my-project
```

---

## 3. FreeStyle 模式

### 自然语言风格

```bash
ppt-design "融资路演" --style "warm fintech"
ppt-design "产品发布" --style "dark cyberpunk tech"
ppt-design "品牌策略" --style "elegant luxury"
ppt-design "ESG报告" --style "calm nature sustainability"
ppt-design "创业路演" --style "bold startup vibrant"
```

### 精确原子控制

```bash
ppt-design "融资路演" \
  --palette wine-burgundy \
  --fonts elegant-serif \
  --decoration gold-trim \
  --layout-variant centered
```

### 设计拨盘

```bash
ppt-design "融资路演" --variance 8 --motion 6 --density 7
```

| 拨盘 | 范围 | 效果 |
|------|------|------|
| `--variance` | 1-10 | 设计变化程度，越高越独特 |
| `--motion` | 1-10 | 动画强度，1=无动画，10=最强 |
| `--density` | 1-10 | 内容密度，1=大字少量，10=密集信息 |

### AI 配图

```bash
# Seedream（推荐，字节豆包）
ppt-design "融资路演" --fetch-images --llm-provider seedream --llm-api-key $ARK_API_KEY

# GPT Image
ppt-design "产品介绍" --fetch-images --llm-provider gpt-image --llm-api-key $OPENAI_API_KEY

# 指定模型
ppt-design "融资路演" --fetch-images --llm-provider seedream --llm-model doubao-seedream-5-0-260128 --llm-api-key $ARK_API_KEY
```

### 预置主题（向后兼容）

```bash
ppt-design "融资路演" --theme professional
ppt-design "产品发布" --theme dark-tech
ppt-design "品牌策略" --theme warm-elegant
ppt-design "创业路演" --theme vibrant-startup
ppt-design "ESG报告" --theme nature-calm
```

---

## 4. Enterprise 模式

### 项目目录结构

```
my-project/
├── template.pptx      # 可选：品牌模板
├── brand.json         # 可选：品牌规范
├── content.json       # 可选：真实内容
├── logo.png           # 可选：公司LOGO
├── images/            # 可选：图片池
│   ├── product.png
│   └── team.jpg
└── output/            # 自动生成
    ├── v1/
    │   ├── presentation.pptx
    │   └── meta.json
    └── v2/
        └── ...
```

### 基本用法

```bash
# 1. 初始生成
ppt-design "融资路演" --project ./my-project --density 6 --motion 5

# 2. 预览方案（不生成文件）
ppt-design "融资路演" --project ./my-project --dry-run

# 3. 确认流程
ppt-design "融资路演" --project ./my-project --review

# 4. 查看版本历史
ppt-design "" --project ./my-project --history

# 5. 使用不同内容文件
ppt-design "融资路演" --project ./my-project --content ./v3-content.json

# 6. 基于旧版本修订
ppt-design "融资路演" --project ./my-project --from-version 1 --density 8
```

### 业务模式

```bash
ppt-design "融资路演" --project ./my-project --business-mode pitch
ppt-design "课程大纲" --project ./my-project --business-mode education
ppt-design "培训手册" --project ./my-project --business-mode training
ppt-design "年度报告" --project ./my-project --business-mode report
```

| 模式 | 叙事策略 |
|------|---------|
| `pitch` | 默认策略（YC Seed Deck 风格） |
| `education` | Education Course — 知识点 → 练习 |
| `training` | Training Workshop — 目标 → 演练 |
| `report` | Business Report — 数据 → 洞察 |

### 页面修订

```bash
# 删除第 3 页
ppt-design "" --project ./my-project --pages "-3"

# 交换第 2 页和第 5 页
ppt-design "" --project ./my-project --pages "2<>5"

# 将第 10 页移到第 3 页后面
ppt-design "" --project ./my-project --pages "10>3"

# 在第 6 页后插入空白页
ppt-design "" --project ./my-project --pages "+6"

# 组合操作：删除 3、交换 2和5、移动 10到3
ppt-design "" --project ./my-project --pages "-3,2<>5,10>3"
```

> 详细语法见 [第 7 节](#7-页面修订语法)

### 版本管理

Enterprise Pipeline 自动管理版本：

```
output/
├── v1/               # 第一次生成
│   ├── presentation.pptx
│   └── meta.json     # 包含每页 goal/title
├── v2/               # 第二次生成
│   ├── presentation.pptx
│   └── meta.json
└── v3/               # 第三次生成
    └── ...
```

- `--output-version N`：指定版本号（覆盖已有版本）
- `--from-version N`：基于 N 版 meta.json 上下文修订
- `--history`：查看所有版本

### 视觉设计

Enterprise Pipeline 自动为每页应用品牌视觉：

| 页面类型 | 视觉效果 |
|---------|---------|
| **Hook / CTA** | 品牌主色全屏覆盖（90% 不透明度），浅色文字 |
| **内容页** | 品牌背景色 + 左侧 accent 竖条（accent 色）+ 底部 muted 横条 |
| **所有页面** | 标题用 foreground 色，副标题用 muted-foreground 色，正文用 foreground 色 |

---

## 5. content.json 内容格式

### 完整示例

```json
{
  "meta": {
    "title": "Acme Corp — Series B Pitch",
    "author": "Investor Relations"
  },
  "slides": [
    {
      "goal": "hook",
      "title": "The Future of Enterprise AI",
      "subtitle": "Acme Corp — 2026 Product Launch",
      "image": "images/hero.png"
    },
    {
      "goal": "problem",
      "title": "The Problem",
      "bullets": [
        "75% of AI projects fail to reach production",
        "Data silos create 3x more integration work",
        "Legacy systems cannot scale"
      ]
    },
    {
      "goal": "solution",
      "title": "Our Solution",
      "subtitle": "One Platform. Every Model.",
      "bullets": [
        "Unified inference gateway",
        "Auto-scaling engine",
        "One-click deployment"
      ],
      "image": "images/product.png"
    },
    {
      "goal": "features",
      "title": "Key Features",
      "cards": [
        {"title": "Fast", "text": "Sub-100ms inference latency"},
        {"title": "Secure", "text": "SOC2 + HIPAA compliant"},
        {"title": "Scalable", "text": "0 to 10K requests/sec"}
      ]
    },
    {
      "goal": "market",
      "title": "Market Opportunity",
      "diagram": {
        "type": "funnel",
        "data": {
          "items": [
            {"text": "TAM $120B"},
            {"text": "SAM $45B"},
            {"text": "SOM $8B"}
          ]
        }
      }
    },
    {
      "goal": "architecture",
      "title": "System Architecture",
      "diagram": {
        "type": "flowchart",
        "data": {
          "nodes": [
            {"id": "src", "text": "Data Sources"},
            {"id": "etl", "text": "ETL Pipeline"},
            {"id": "model", "text": "Model Registry"},
            {"id": "serve", "text": "Inference Engine"}
          ],
          "connectors": [
            ["src", "etl"],
            ["etl", "model"],
            ["model", "serve"]
          ]
        }
      }
    },
    {
      "goal": "code_demo",
      "title": "Quick Start",
      "code": {
        "language": "python",
        "source": "from acme import AIPlatform\nplatform = AIPlatform(api_key='key')\nresult = platform.predict(data)"
      }
    },
    {
      "goal": "exercise",
      "title": "Try It Yourself",
      "exercise": {
        "instructions": "Deploy your first endpoint in 5 minutes",
        "duration": "5 min",
        "steps": [
          "Sign up at acme.ai",
          "Create a new project",
          "Click 'Deploy'"
        ]
      }
    },
    {
      "goal": "swot",
      "title": "Strategic Analysis",
      "diagram": {
        "type": "swot",
        "data": {
          "strengths": ["Proprietary tech", "Strong team"],
          "weaknesses": ["Early stage"],
          "opportunities": ["AI boom"],
          "threats": ["Big Tech entry"]
        }
      }
    },
    {
      "goal": "cta",
      "title": "Let's Build Together",
      "subtitle": "contact@acme.ai"
    }
  ]
}
```

### 字段说明

| 字段 | 类型 | 说明 |
|------|------|------|
| `goal` | str | 页面目标，默认 `"content"`。hook/cta 获得全屏品牌色覆盖 |
| `title` | str | 页面标题 |
| `subtitle` | str | 副标题 |
| `bullets` | list[str] | 要点列表 |
| `image` | str | 图片路径（相对项目目录或绝对路径） |
| `cards` | list[dict] | 卡片列表，每张含 `title`/`text`/可选 `image` |
| `diagram` | dict | 图形定义，含 `type` 和 `data`（见第 8 节） |
| `code` | str 或 dict | 代码块：字符串或 `{"source": "...", "language": "python"}` |
| `exercise` | dict | `{"instructions": str, "duration": str, "steps": list[str]}` |
| `notes` | str | 演讲者备注（v0.4.0） |
| `links` | list[dict] | 超链接（v0.4.0） |
| `chart` | dict | 数据图表（v0.4.0） |

### 图形数据格式别名

content.json 支持自然格式，Pipeline 自动转换：

| 图形类型 | content.json 写法 | 内部转换 |
|---------|------------------|---------|
| funnel | `"items": [{"text": "..."}]` | → `stages: [{"label": "..."}]` |
| timeline | `"items": [{"text": "..."}]` | → `events: [{"label": "..."}]` |
| swot | `"strengths": [...]` | → `quadrants: [{label: "Strengths", items: [...]}]` |
| pyramid | `"items": [{"text": "..."}]` | → `levels: [{"label": "..."}]` |
| flowchart | `"nodes": [{"text": "..."}]` | → `nodes: [{"label": "..."}]` |

---

## 6. brand.json 品牌格式

### 完整示例

```json
{
  "colors": {
    "primary": "#2563EB",
    "on-primary": "#FFFFFF",
    "secondary": "#64748B",
    "accent": "#F97316",
    "foreground": "#1A1A2E",
    "on-primary": "#FFFFFF",
    "muted-foreground": "#94A3B8",
    "border": "#E2E8F0",
    "background": "#FFFFFF",
    "muted": "#F1F5F9"
  },
  "fonts": {
    "heading": "Calibri",
    "body": "Calibri"
  },
  "logo": {
    "position": "top_right",
    "width_inches": 1.2,
    "skip_slides": ["hook"]
  },
  "layout_mapping": {
    "hook": 0,
    "problem": 1,
    "solution": 1,
    "cta": 0
  },
  "footer": {
    "show_page_number": true,
    "page_number_format": "第 {n} 页 / 共 {total} 页",
    "page_number_position": "bottom_right",
    "show_footer_text": true,
    "footer_text": "Confidential — ACME Corp",
    "footer_position": "bottom_center"
  },
  "watermark": {
    "text": "DRAFT",
    "opacity": 0.15,
    "rotation": -45,
    "skip_pages": [1]
  },
  "dark_mode": false
}
```

### 颜色角色

| 角色 | 用途 |
|------|------|
| `primary` | 主色，hook/cta 页全屏覆盖，图形节点默认填充 |
| `on-primary` | 主色上的文字色 |
| `secondary` | 辅助色 |
| `accent` | 强调色，左侧 accent 竖条 |
| `foreground` | 正文/标题文字色 |
| `muted-foreground` | 副标题/次要文字色 |
| `background` | 页面背景色 |
| `muted` | 底部横条/浅色背景 |
| `border` | 边框色 |

### Logo 配置

| 字段 | 默认值 | 说明 |
|------|--------|------|
| `position` | `"top_right"` | `top_right` / `top_left` / `bottom_right` |
| `width_inches` | `1.0` | LOGO 宽度（英寸） |
| `skip_slides` | `[]` | 跳过 LOGO 的页面 goal 列表 |

### 页脚配置

| 字段 | 默认值 | 说明 |
|------|--------|------|
| `show_page_number` | `false` | 显示页码 |
| `page_number_format` | `"{n}"` | `{n}`=当前页，`{total}`=总页数 |
| `page_number_position` | `"bottom_right"` | `bottom_left` / `bottom_center` / `bottom_right` |
| `show_footer_text` | `false` | 显示页脚文本 |
| `footer_text` | `""` | 页脚文字 |
| `footer_position` | `"bottom_center"` | 同上位置 |
| `font_size_pt` | `10` | 字号 |
| `skip_pages` | `[]` | 跳过页码的页号（封面自动跳过） |

### 水印配置

| 字段 | 默认值 | 说明 |
|------|--------|------|
| `text` | `"CONFIDENTIAL"` | 水印文字 |
| `opacity` | `0.15` | 不透明度（0.0-1.0） |
| `rotation` | `-45` | 旋转角度 |
| `font_size_pt` | `72` | 字号 |
| `skip_pages` | `[1]` | 跳过的页号 |

### 合并规则

当同时存在 `template.pptx` 和 `brand.json` 时：

- **颜色/字体**：brand.json 逐键覆盖模板分析结果
- **logo / footer / watermark**：brand.json 优先
- **layout_mapping**：brand.json 优先
- **template_layouts**：始终来自模板分析
- **dark_mode**：始终来自模板分析

---

## 7. 页面修订语法

### 操作类型

| 语法 | 操作 | 示例 | 说明 |
|------|------|------|------|
| `N` | 修改 | `3` | 标记第 N 页需修改 |
| `N-M` | 范围修改 | `2-5` | 标记第 2-5 页需修改 |
| `+N` | 插入 | `+6` | 在第 6 页后插入空白页 |
| `-N` | 删除 | `-8` | 删除第 8 页 |
| `N>M` | 移动 | `3>5` | 将第 3 页移到第 5 页后 |
| `N<>M` | 交换 | `3<>7` | 交换第 3 页和第 7 页 |

### 组合操作

用逗号分隔多个操作：

```bash
# 删除第 3 页，交换第 2 和第 5 页，将第 10 页移到第 3 页后
--pages "-3,2<>5,10>3"

# 删除第 9、10 页，交换第 2 和第 7 页，交换第 11 和第 15 页
--pages "-9,-10,2<>7,11<>15"
```

### 执行顺序

1. **交换** (swap) — 先执行，不影响后续索引
2. **移动** (move) — 交换后执行
3. **删除** (delete) — 按页码从大到小删除
4. **插入** (insert) — 最后执行

### 页码说明

- 所有页码为 **1-based**，指原始文档页码
- 删除操作不影响后续操作的页码引用（均基于原始页码）
- 页码超出范围会报错

---

## 8. 10 种图形类型

### 通用格式

```json
{
  "diagram": {
    "type": "<图形类型>",
    "data": { <类型专用数据> }
  }
}
```

### 8.1 流程图 (flowchart)

```json
{
  "type": "flowchart",
  "data": {
    "nodes": [
      {"id": "start", "text": "开始"},
      {"id": "process", "text": "处理"},
      {"id": "end", "text": "结束"}
    ],
    "connectors": [["start", "process"], ["process", "end"]],
    "direction": "horizontal"
  }
}
```

- `direction`：可选 `"horizontal"` 或 `"vertical"`，默认自动（≤6 节点横排）

### 8.2 漏斗图 (funnel)

```json
{
  "type": "funnel",
  "data": {
    "items": [
      {"text": "TAM $120B"},
      {"text": "SAM $45B"},
      {"text": "SOM $8B"}
    ]
  }
}
```

### 8.3 时间线 (timeline)

```json
{
  "type": "timeline",
  "data": {
    "items": [
      {"text": "Q1: Beta", "position": "start"},
      {"text": "Q2: GA", "position": "middle"},
      {"text": "Q3: EU", "position": "middle"},
      {"text": "Q4: IPO", "position": "end"}
    ]
  }
}
```

### 8.4 SWOT 分析 (swot)

```json
{
  "type": "swot",
  "data": {
    "strengths": ["技术领先", "团队优秀"],
    "weaknesses": ["早期阶段"],
    "opportunities": ["AI 浪潮", "合规需求"],
    "threats": ["巨头入场"]
  }
}
```

### 8.5 矩阵图 (matrix)

```json
{
  "type": "matrix",
  "data": {
    "rows": [{"label": "产品A"}, {"label": "产品B"}],
    "cols": [{"label": "功能"}, {"label": "价格"}, {"label": "服务"}],
    "cells": [["强", "中", "强"], ["中", "强", "弱"]]
  }
}
```

### 8.6 循环图 (cycle)

```json
{
  "type": "cycle",
  "data": {
    "nodes": [
      {"text": "部署"},
      {"text": "测量"},
      {"text": "扩展"},
      {"text": "推荐"}
    ]
  }
}
```

### 8.7 表格 (table)

```json
{
  "type": "table",
  "data": {
    "headers": ["特性", "基础版", "专业版"],
    "rows": [
      ["用户数", "5", "无限"],
      ["API 调用", "1K/月", "无限"]
    ]
  }
}
```

### 8.8 层级图 (hierarchy)

```json
{
  "type": "hierarchy",
  "data": {
    "nodes": [
      {"id": "ceo", "text": "CEO"},
      {"id": "cto", "text": "CTO", "parent": "ceo"},
      {"id": "cfo", "text": "CFO", "parent": "ceo"},
      {"id": "vp1", "text": "VP Eng", "parent": "cto"}
    ]
  }
}
```

### 8.9 金字塔图 (pyramid)

```json
{
  "type": "pyramid",
  "data": {
    "items": [
      {"text": "网络效应"},
      {"text": "数据壁垒"},
      {"text": "技术壁垒"},
      {"text": "品牌壁垒"}
    ]
  }
}
```

### 8.10 韦恩图 (venn)

```json
{
  "type": "venn",
  "data": {
    "sets": [
      {"label": "AI 基础设施"},
      {"label": "合规治理"},
      {"label": "开发者体验"}
    ]
  }
}
```

---

## 9. 设计系统 — 40,000+ 风格组合

**25 色彩方案 × 20 字体搭配 × 10 装饰风格 × 8 布局变体 = 40,000 种组合**

### 25 色彩方案

| 名称 | 主色 | 背景 | 暗色 |
|------|------|------|------|
| `ocean-blue` | #1E5AB8 | #FFFFFF | - |
| `midnight-navy` | #1E3A5F | #0A1E3D | Y |
| `cyber-neon` | #6366F1 | #060B18 | Y |
| `neon-gradient` | #8B5CF6 | #120C1E | Y |
| `golden-luxury` | #C99A4E | #FAF3E8 | - |
| `rose-gold` | #B76E79 | #FFF5F5 | - |
| `forest-green` | #1B5E20 | #F5F7F0 | - |
| `sage-calm` | #5B7553 | #F4F7F1 | - |
| `sunset-warm` | #D97706 | #FFFBEB | - |
| `terracotta` | #C4704B | #FBF5F0 | - |
| `cherry-red` | #DC2626 | #FEF2F2 | - |
| `royal-purple` | #7C3AED | #F5F3FF | - |
| `arctic-frost` | #0EA5E9 | #F0F9FF | - |
| `slate-minimal` | #475569 | #F8FAFC | - |
| `charcoal-bold` | #1F2937 | #111827 | Y |
| `coral-energy` | #F97316 | #FFF7ED | - |
| `teal-fresh` | #0D9488 | #F0FDFA | - |
| `indigo-deep` | #4338CA | #EEF2FF | - |
| `copper-industrial` | #B87333 | #2D2D2D | Y |
| `monochrome` | #18181B | #FAFAFA | - |
| `monochrome-dark` | #D4D4D8 | #09090B | Y |
| `lavender-dream` | #8B5CF6 | #FAF5FF | - |
| `mint-fresh` | #10B981 | #ECFDF5 | - |
| `wine-burgundy` | #7F1D1D | #1C1010 | Y |
| `sky-bright` | #0284C7 | #F0F9FF | - |

### 20 字体搭配

| 名称 | 标题字体 | 正文字体 |
|------|---------|---------|
| `modern-sans` | Inter | Inter |
| `geometric-sans` | Space Grotesk | Inter |
| `bold-sans` | Poppins | Inter |
| `clean-corporate` | Calibri | Calibri |
| `serif-editorial` | Georgia | Georgia |
| `elegant-serif` | Playfair Display | Inter |
| `literary-serif` | Lora | Inter |
| `tech-mono` | Consolas | Consolas |
| `mono-clean` | Consolas | Inter |
| `swiss-style` | Arial | Arial |
| `humanist-sans` | Segoe UI | Segoe UI |
| `friendly-round` | Nunito | Inter |
| `sharp-modern` | Montserrat | Inter |
| `classic-formal` | Times New Roman | Times New Roman |
| `contrast-mix` | Playfair Display | Calibri |
| `tech-contrast` | Space Grotesk | Consolas |
| `warm-mix` | Georgia | Calibri |
| `startup-mix` | Poppins | Segoe UI |
| `minimal-mix` | Inter | Calibri |
| `editorial-mix` | Georgia | Segoe UI |

### 10 装饰风格

| 名称 | 效果 |
|------|------|
| `accent-bar` | 简洁强调条 + 下划线 |
| `neon-lines` | 霓虹发光线条 |
| `gold-trim` | 金色装饰线 |
| `minimal-dots` | 圆点项目符号 |
| `diamond-bullets` | 菱形项目符号 |
| `gradient-bar` | 渐变强调条 |
| `circle-accent` | 圆形装饰元素 |
| `sidebar-nav` | 左侧导航栏 |
| `no-decoration` | 无装饰 |
| `full-bleed-overlay` | 全出血图 + 暗色叠层 |

### 8 布局变体

| 名称 | 特点 |
|------|------|
| `standard` | 标准左对齐，0.9" 边距 |
| `centered` | 居中编辑式，2.0" 边距 |
| `sidebar-left` | 左侧导航栏 + 主内容区 |
| `sidebar-right` | 右侧统计/引用栏 |
| `wide-cards` | 宽卡片布局 |
| `grid-2x2` | 2×2 度量卡片网格 |
| `asymmetric` | 不对称交错布局 |
| `full-width` | 全宽边缘到边缘 |

### 风格关键词

通过 `--style` 使用自然语言，系统自动匹配最合适的设计原子：

| 关键词 | 匹配倾向 |
|--------|---------|
| professional/corporate | ocean-blue + clean-corporate + accent-bar |
| tech/dark/cyberpunk | cyber-neon + tech-mono + neon-lines |
| warm/elegant/luxury | golden-luxury + elegant-serif + gold-trim |
| startup/vibrant/bold | neon-gradient + bold-sans + gradient-bar |
| nature/calm/sustainability | forest-green + humanist-sans + circle-accent |
| minimal/clean | slate-minimal + modern-sans + no-decoration |
| fintech/finance | ocean-blue + clean-corporate + sidebar-nav |
| education/learning | ocean-blue + clean-corporate + minimal-dots |

---

## 10. 动画系统

### 12 种幻灯片切换

| 类型 | 效果 |
|------|------|
| `fade` | 淡入淡出 |
| `push` | 推入 |
| `wipe` | 擦除 |
| `split` | 分裂 |
| `cover` | 覆盖 |
| `dissolve` | 溶解 |
| `wheel` | 轮辐（4 辐） |
| `wedge` | 楔形展开 |
| `blinds` | 百叶窗 |
| `checker` | 棋盘格 |
| `comb` | 梳状 |
| `random` | 随机 |

### 10 种入场动画

| 类型 | 效果 |
|------|------|
| `appear` | 瞬间出现 |
| `fly_in` | 从下方飞入 |
| `fly_in_left` | 从左侧飞入 |
| `fly_in_right` | 从右侧飞入 |
| `fly_in_top` | 从上方飞入 |
| `fly_in_bottom` | 从下方飞入 |
| `fade_in` | 淡入 |
| `zoom_in` | 缩放进入 |
| `float_up` | 上浮 |
| `grow_turn` | 旋转放大 |
| `bounce` | 弹跳进入 |

### Motion 等级对应

| Motion | 切换速度 | 入场动画 |
|--------|---------|---------|
| 1-2 | 慢速 | 无 |
| 3-5 | 中速 | fade_in |
| 6-10 | 快速 | fly_in |

---

## 11. 密度控制

`--density` 1-10 控制每页信息量：

| 等级 | 标题字号 | 正文字号 | 最大要点数 | 最大字符数 | 图片比例 |
|------|---------|---------|-----------|-----------|---------|
| 1 | 36pt | 16pt | 3 | 40 | 50% |
| 2 | 34pt | 15pt | 4 | 50 | 45% |
| 3 | 32pt | 14pt | 5 | 60 | 40% |
| 4 | 30pt | 13pt | 5 | 70 | 38% |
| 5 | 28pt | 12pt | 6 | 80 | 35% |
| 6 | 26pt | 11pt | 7 | 90 | 32% |
| 7 | 24pt | 10pt | 8 | 100 | 30% |
| 8 | 22pt | 9pt | 10 | 120 | 25% |
| 9 | 20pt | 8pt | 12 | 140 | 22% |
| 10 | 18pt | 7pt | 15 | 160 | 20% |

- 超长要点自动截断并加省略号
- 超量要点自动丢弃
- 图片宽度按比例缩放

---

## 12. 图片引擎

### 5 种图片模式

| 模式 | 说明 | 需要 API Key |
|------|------|-------------|
| `placeholder` | 渐变占位图（默认） | 否 |
| `search` | Unsplash/Pexels 搜索 | 是 |
| `generate` | AI 图片生成 | 是 |
| `enhance` | Kimi 优化关键词后搜索 | 是 |
| `auto` | 先 AI 生成，失败则搜索 | 是 |

### 5 种 AI 生成引擎

| 引擎 | CLI 参数 | 默认模型 | 环境变量 |
|------|---------|---------|---------|
| **Seedream** | `--llm-provider seedream` | `doubao-seedream-5-0-260128` | `ARK_API_KEY` |
| **GPT Image** | `--llm-provider gpt-image` | `gpt-image-1` | `OPENAI_API_KEY` |
| **DALL-E 3** | `--llm-provider dalle` | `dall-e-3` | `OPENAI_API_KEY` |
| **通义万相** | `--llm-provider wanx` | `wanx-v1` | `DASHSCOPE_API_KEY` |
| **Kimi** | `--llm-provider kimi` | `kimi-k2-0711-preview` | `MOONSHOT_API_KEY` |

### Seedream 可用模型

| 模型名 | 说明 |
|--------|------|
| `doubao-seedream-5-0-260128` | **默认**，Seedream 5.0 |
| `doubao-seedream-5-0-pro-260628` | Seedream 5.0 Pro |
| `doubao-seedream-4-5-251128` | Seedream 4.5 |

```bash
# 使用 Seedream 5.0 Pro
ppt-design "路演" --fetch-images --llm-provider seedream --llm-model doubao-seedream-5-0-pro-260628 --llm-api-key $ARK_API_KEY
```

### 缓存机制

所有图片引擎内置**缓存优先**：相同关键词+尺寸不会重复调用 API，节省费用。

---

## 13. Python API

```python
from ppt_pro_max import generate_ppt

# FreeStyle
result = generate_ppt("AI startup pitch", style="dark cyberpunk", motion=7)
print(result["output_path"])

# Enterprise
result = generate_ppt(
    "Investor Pitch",
    project="./my-project",
    content_file="./content.json",
    density=6,
    motion=5,
)
print(f"v{result['version']}: {result['num_slides']} slides")

# Page revision
result = generate_ppt("", project="./my-project", pages="-3,2<>5")

# Version history
result = generate_ppt("", project="./my-project", history=True)
for v in result["versions"]:
    print(f"v{v['version']}: {v['meta']['num_slides']} slides")
```

### 返回值

**FreeStyle:**
```python
{"output_path": "...", "page_count": 10, "strategy": "...", "theme": "..."}
```

**Enterprise:**
```python
{"pipeline": "enterprise", "project": "...", "output_path": "...", "version": 1, "num_slides": 12}
```

---

## 14. CLI 完整参数

```
ppt-design [query] [options]

位置参数:
  query                  演示主题（--history 模式可省略）

风格选项:
  --strategy STR         覆盖叙事策略
  --theme THEME          预置主题名称
  --style STYLE          自然语言风格
  --palette PALETTE      色彩方案名称
  --fonts FONTS          字体搭配名称
  --decoration DEC       装饰风格名称
  --layout-variant LV    布局变体名称
  --mood MOOD            风格关键词
  --style-seed N         风格随机种子（可复现）
  --slides N             覆盖页数
  --content FILE         内容 JSON 文件
  --variance 1-10        设计变化程度
  --motion 1-10          动画强度
  --density 1-10         内容密度

图片选项:
  --image-mode MODE      placeholder/search/generate/enhance
  --fetch-images         快捷搜索图片
  --unsplash-key KEY     Unsplash API Key
  --pexels-key KEY       Pexels API Key
  --llm-provider PROV    seedream/gpt-image/dalle/wanx/kimi
  --llm-api-key KEY      LLM API Key
  --llm-base-url URL     LLM API 基础 URL
  --llm-model MODEL      LLM 模型名称

输出选项:
  --persist              保存设计系统为 MASTER.md
  --dry-run              仅输出设计决策
  -o, --output PATH      输出文件路径

企业选项:
  --project DIR          项目目录（触发 Enterprise Pipeline）
  --business-mode MODE   pitch/education/training/report
  --review               启用方案确认
  --review-file PATH     方案输出 JSON 文件
  --output-version N     指定版本号
  --from-version N       基于指定版本修订
  --pages OPS            页面操作（如 "3,5 +6 -8 3>5 3<>7"）
  --history              查看版本历史
```
