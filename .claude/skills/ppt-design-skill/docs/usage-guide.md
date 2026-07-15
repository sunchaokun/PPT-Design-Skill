# PPT-Design-Skill 使用手册

> 版本 v0.7.0 | 完整功能参考

---

## 目录

1. [三种模式](#1-三种模式)
2. [快速开始](#2-快速开始)
3. [FreeStyle 模式](#3-freestyle-模式)
4. [Enterprise 模式](#4-enterprise-模式)
5. [Build Script 模式 — 精确控制](#5-build-script-模式--精确控制)
6. [content.json 内容格式](#6-contentjson-内容格式)
7. [brand.json 品牌格式](#7-brandjson-品牌格式)
8. [组件库数据库](#8-组件库数据库)
9. [页面修订语法](#9-页面修订语法)
10. [10 种图形类型](#10-10-种图形类型)
11. [设计系统 — 40,000+ 风格组合](#11-设计系统--40000-风格组合)
12. [设计质量升级 — 28 项专业级提升](#12-设计质量升级--28-项专业级提升)
13. [动画系统](#13-动画系统)
14. [密度控制](#14-密度控制)
15. [图片引擎](#15-图片引擎)
16. [Python API](#16-python-api)
17. [CLI 完整参数](#17-cli-完整参数)
18. [最佳实践与质量保证](#18-最佳实践与质量保证)

---

## 1. 三种模式

PPT-Design-Skill 提供三种生成模式，从快速到精确：

| | FreeStyle | Enterprise | Build Script |
|---|---|---|---|
| **适用场景** | 快速生成、风格探索 | 企业合规、团队协作 | **交付级质量、逐页精确控制** |
| **触发方式** | 默认（无 `--project`） | `--project <目录>` | 手写 `build.py` 脚本 |
| **内容来源** | AI 自动生成 | content.json 精确控制 | 代码中硬编码每页内容 |
| **品牌规范** | 风格原子组合 | brand.json + 模板分析 | Design Token 字典 |
| **布局控制** | 自动匹配母版 | 自动匹配母版 | **逐元素 x/y/w/h 精确定位** |
| **字体控制** | 主题级 | 主题级 | **run-level 逐字设置** |
| **版本管理** | 无 | v1 → v2 → v3 ... | Git 管理 build.py |
| **页面修订** | 不支持 | `--pages` 增删改查 | 直接改代码 |
| **质量上限** | ★★★ | ★★★ | ★★★★★ |

> **选择建议**：探索阶段用 FreeStyle，团队协作用 Enterprise，**最终交付用 Build Script**。

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

> 详细语法见 [第 9 节](#9-页面修订语法)

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

## 5. Build Script 模式 — 精确控制

### 为什么需要 Build Script？

FreeStyle 和 Enterprise 模式通过 pipeline 自动渲染，布局由 LayoutRegistry 的母版模板决定。这种方式适合快速生成，但**无法精确控制每个元素的位置、大小、字体**——而这正是专业 PPT 的核心要求。

Build Script 模式直接使用 python-pptx API，**逐页、逐元素**控制 PPT 的每一个细节，效果可媲美手动设计。

### 核心原则

1. **逐元素定位**：每个文本框、图片、形状都指定 x, y, w, h，不依赖自动布局
2. **run-level 字体**：字体属性设在 `run.font` 上，不设在 `p.font` 上（PowerPoint 忽略 paragraph 级别）
3. **Pillow 预裁剪**：图片先用 Pillow center-crop 到目标比例，再插入，绝不拉伸变形
4. **全幅图+遮罩**：封面/案例页用全幅背景图 + 半透明暗色矩形 + 文字叠加
5. **生成后必检**：每次生成后跑检查脚本，验证元素数量、字体、尺寸

### 项目结构

```
my-project/
├── build.py            # 核心构建脚本
├── brand.json          # 可选：Design Token 参考
├── content.json        # 可选：内容数据参考
├── images/             # 本地图片素材
└── output/             # 生成输出
    └── presentation.pptx
```

### 最小示例

```python
"""build.py — 最小 Build Script 示例"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

# ── Slide 1: Cover ──
sl = prs.slides.add_slide(blank)

# 背景矩形
bg = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
bg.fill.solid()
bg.fill.fore_color.rgb = RGBColor(0x06, 0x0B, 0x18)
bg.line.fill.background()

# 标题 — run-level 字体设置
tb = sl.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(8), Inches(1.5))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
run = p.add_run()
run.text = "My Presentation"
run.font.name = "Space Grotesk"    # ← run-level，PowerPoint 认
run.font.size = Pt(52)             # ← run-level
run.font.color.rgb = RGBColor(0xF8, 0xFA, 0xFC)
run.font.bold = True

prs.save("output/presentation.pptx")
```

### 完整工具函数库

以下工具函数覆盖 90% 的 PPT 构建需求，可直接复制到 build.py 中使用：

#### Design Token 系统

```python
# ── Design Token — 一处修改，全局生效 ──
TOKENS = {
    'cyber-neon': {
        'bg': '#060B18', 'primary': '#6366F1', 'accent': '#22D3EE',
        'fg': '#F8FAFC', 'muted': '#64748B', 'dim': '#334155',
        'card_bg': '#0D152A', 'card_bd': '#1A2A4A',
        'font_h': 'Space Grotesk', 'font_b': 'Consolas',
    },
    'professional': {
        'bg': '#FFFFFF', 'primary': '#2563EB', 'accent': '#F97316',
        'fg': '#1E293B', 'muted': '#64748B', 'dim': '#94A3B8',
        'card_bg': '#F1F5F9', 'card_bd': '#E2E8F0',
        'font_h': 'Inter', 'font_b': 'Inter',
    },
    'warm-elegant': {
        'bg': '#FAF3E8', 'primary': '#C99A4E', 'accent': '#B8860B',
        'fg': '#1A1A2E', 'muted': '#8B7355', 'dim': '#A0937D',
        'card_bg': '#FFF8F0', 'card_bd': '#E8D5B7',
        'font_h': 'Playfair Display', 'font_b': 'Inter',
    },
}
C = TOKENS['cyber-neon']  # ← 切换主题只改这一行
```

#### 文本函数

```python
def _rgb(h):
    """十六进制颜色 → RGBColor"""
    return RGBColor.from_string(h.lstrip('#'))

def _add_text(slide, text, x, y, w, h, font=None, size=20, color=None, bold=False, align='left'):
    """添加单行文本 — run-level 字体设置"""
    font = font or C['font_h']
    color = color or C['fg']
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}[align]
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.color.rgb = _rgb(color)
    run.font.bold = bold
    return tb

def _add_multiline(slide, lines, x, y, w, h, font=None, size=14, color=None, bold=False, align='left', spacing=6):
    """添加多行文本"""
    font = font or C['font_b']
    color = color or C['fg']
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}[align]
        run = p.add_run()
        run.text = line
        run.font.name = font
        run.font.size = Pt(size)
        run.font.color.rgb = _rgb(color)
        run.font.bold = bold
        p.space_after = Pt(spacing)
    return tb
```

#### 图片函数（Pillow 预裁剪）

```python
from PIL import Image as PILImage
import tempfile, hashlib

def _add_image(slide, path, x, y, w, h):
    """添加图片 — Pillow center-crop 到目标比例，绝不拉伸"""
    if not os.path.isfile(path):
        return
    img = PILImage.open(path)
    iw, ih = img.size
    box_ratio = w / h
    img_ratio = iw / ih
    if img_ratio > box_ratio:
        cw, ch = int(ih * box_ratio), ih
        cl, ct = (iw - cw) // 2, 0
    else:
        cw, ch = iw, int(iw / box_ratio)
        cl, ct = 0, (ih - ch) // 2
    cropped = img.crop((cl, ct, cl + cw, ct + ch))
    td = os.path.join(tempfile.gettempdir(), 'ppt-build-crops')
    os.makedirs(td, exist_ok=True)
    cp = os.path.join(td, hashlib.md5(f'{path}:{w}x{h}'.encode()).hexdigest() + '.png')
    if not os.path.exists(cp):
        cropped.save(cp, 'PNG')
    slide.shapes.add_picture(cp, Inches(x), Inches(y), Inches(w), Inches(h))
```

#### 形状函数

```python
def _add_rect(slide, x, y, w, h, fill, border=None):
    """添加矩形"""
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = _rgb(fill)
    if border:
        sh.line.color.rgb = _rgb(border)
        sh.line.width = Pt(1)
    else:
        sh.line.fill.background()
    return sh

def _add_rounded_rect(slide, x, y, w, h, fill, border=None):
    """添加圆角矩形"""
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = _rgb(fill)
    if border:
        sh.line.color.rgb = _rgb(border)
        sh.line.width = Pt(1)
    else:
        sh.line.fill.background()
    return sh
```

#### 半透明遮罩

```python
from pptx.oxml.ns import qn
from lxml import etree

def _dark_overlay(slide, opacity=0.65):
    """全幅半透明暗色遮罩 — 用于全幅背景图上的文字可读性"""
    ov = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    ov.fill.solid()
    ov.fill.fore_color.rgb = _rgb(C['bg'])
    ov.line.fill.background()
    el = ov._element.find(qn('p:spPr')).find(qn('a:solidFill')).find(qn('a:srgbClr'))
    if el is not None:
        a = etree.SubElement(el, qn('a:alpha'))
        a.set('val', str(int(opacity * 100000)))
```

#### 卡片组件

```python
def _card(slide, x, y, w, h, title, body, accent=None):
    """标准卡片 — 圆角背景 + accent 条 + 标题 + 正文"""
    accent = accent or C['primary']
    _add_rounded_rect(slide, x, y, w, h, C['card_bg'], C['card_bd'])
    _add_rect(slide, x + 0.15, y + 0.15, 0.4, 0.05, accent)  # accent 条
    _add_text(slide, title, x + 0.15, y + 0.35, w - 0.3, 0.4, C['font_h'], 15, accent, True)
    _add_multiline(slide, body.split('\n'), x + 0.15, y + 0.8, w - 0.3, h - 1.0, C['font_b'], 11, C['muted'])
```

### 10 种页面模板

以下是经过调试的最佳布局参数，可直接调用：

#### page_cover — 封面页

```python
def page_cover(prs, title, subtitle, desc, metrics, image_path=None):
    """封面页：全幅图+遮罩+大标题+右侧指标
    metrics: [(value, label), ...]
    """
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    if image_path:
        _add_image(sl, image_path, 0, 0, 13.333, 7.5)
        _dark_overlay(sl, 0.72)
    else:
        _add_rect(sl, 0, 0, 13.333, 7.5, C['bg'])
    _add_text(sl, title, 1.2, 1.5, 8, 1.5, C['font_h'], 52, C['fg'], True)
    _add_text(sl, subtitle, 1.2, 3.2, 8, 0.6, C['font_b'], 22, C['accent'])
    _add_text(sl, desc, 1.2, 4.0, 8, 0.5, C['font_b'], 14, C['muted'])
    for i, (val, label) in enumerate(metrics):
        yy = 2.0 + i * 1.6
        _add_text(sl, val, 9.5, yy, 3, 0.7, C['font_h'], 40, C['accent'], True)
        _add_text(sl, label, 9.5, yy + 0.65, 3, 0.4, C['font_b'], 13, C['muted'])
```

#### page_numbered_list — 编号列表页

```python
def page_numbered_list(prs, title, items, side_image=None):
    """编号列表页：左侧编号卡片 + 右侧配图
    items: [(title, description), ...]
    """
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _add_rect(sl, 0, 0, 0.06, 7.5, C['primary'])
    _add_rect(sl, 0, 7.25, 13.333, 0.25, C['card_bg'])
    _add_text(sl, title, 0.9, 0.5, 11, 0.6, C['font_h'], 28, C['fg'], True)
    _add_rect(sl, 0.9, 1.2, 2, 0.04, C['accent'])
    for i, (t, d) in enumerate(items):
        yy = 1.6 + i * 1.8
        _add_rounded_rect(sl, 0.9, yy, 7, 1.5, C['card_bg'], C['card_bd'])
        _add_text(sl, f'0{i+1}', 1.1, yy + 0.15, 0.6, 0.5, C['font_h'], 24, C['accent'], True)
        _add_text(sl, t, 1.9, yy + 0.15, 5.5, 0.4, C['font_h'], 18, C['fg'], True)
        _add_text(sl, d, 1.9, yy + 0.65, 5.5, 0.7, C['font_b'], 12, C['muted'])
    if side_image:
        _add_image(sl, side_image, 8.3, 1.2, 4.2, 5.3)
```

#### page_process_flow — 流程页

```python
def page_process_flow(prs, title, subtitle, stages):
    """流程页：N 列卡片 + 箭头连接
    stages: [(num, en_name, cn_name, description), ...]
    """
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _add_rect(sl, 0, 0, 0.06, 7.5, C['primary'])
    _add_rect(sl, 0, 7.25, 13.333, 0.25, C['card_bg'])
    _add_text(sl, title, 0.9, 0.5, 11, 0.6, C['font_h'], 28, C['fg'], True)
    _add_rect(sl, 0.9, 1.2, 2, 0.04, C['accent'])
    _add_text(sl, subtitle, 0.9, 1.5, 10, 0.5, C['font_b'], 14, C['muted'])
    n = len(stages)
    col_w = min(2.6, (11.5 - (n - 1) * 0.3) / n)
    for i, (num, en, cn, desc) in enumerate(stages):
        xx = 0.9 + i * (col_w + 0.3)
        _add_rounded_rect(sl, xx, 2.3, col_w, 3.0, C['card_bg'], C['card_bd'])
        _add_rect(sl, xx + 0.15, 2.5, 0.4, 0.05, C['accent'])
        _add_text(sl, num, xx + 0.15, 2.7, 0.5, 0.4, C['font_h'], 16, C['accent'], True)
        _add_text(sl, en, xx + 0.15, 3.2, col_w - 0.3, 0.4, C['font_h'], 15, C['fg'], True)
        _add_text(sl, cn, xx + 0.15, 3.7, col_w - 0.3, 0.4, C['font_b'], 14, C['fg'])
        _add_text(sl, desc, xx + 0.15, 4.3, col_w - 0.3, 0.5, C['font_b'], 11, C['muted'])
        if i < n - 1:
            _add_text(sl, '→', xx + col_w + 0.02, 3.5, 0.26, 0.4, C['font_h'], 20, C['accent'], True, 'center')
```

#### page_3cards — 三列卡片页

```python
def page_3cards(prs, title, cards):
    """三列卡片页
    cards: [(title, body, accent_color), ...]
    """
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _add_rect(sl, 0, 0, 0.06, 7.5, C['primary'])
    _add_rect(sl, 0, 7.25, 13.333, 0.25, C['card_bg'])
    _add_text(sl, title, 0.9, 0.5, 11, 0.6, C['font_h'], 28, C['fg'], True)
    _add_rect(sl, 0.9, 1.2, 2, 0.04, C['accent'])
    for i, (t, body, accent) in enumerate(cards):
        xx = 0.9 + i * 4.1
        _card(sl, xx, 1.6, 3.6, 4.5, t, body, accent)
```

#### page_showcase — 案例展示页

```python
def page_showcase(prs, image_path, style_en, style_cn, atoms, desc):
    """案例展示页：全幅图+遮罩+风格名+设计原子"""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _add_image(sl, image_path, 0, 0, 13.333, 7.5)
    _dark_overlay(sl, 0.55)
    _add_text(sl, style_en, 0.8, 1.2, 6, 0.7, C['font_h'], 18, C['accent'], True)
    _add_text(sl, style_cn, 0.8, 2.0, 6, 1.2, C['font_h'], 44, C['fg'], True)
    _add_text(sl, atoms, 0.8, 3.4, 6, 0.4, C['font_b'], 14, C['muted'])
    _add_text(sl, desc, 0.8, 4.0, 6, 0.4, C['font_b'], 13, C['dim'])
```

#### page_funnel — 漏斗页

```python
def page_funnel(prs, title, items):
    """漏斗页：递减宽度圆角条
    items: [(value, label, width, color), ...]
    """
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _add_rect(sl, 0, 0, 0.06, 7.5, C['primary'])
    _add_rect(sl, 0, 7.25, 13.333, 0.25, C['card_bg'])
    _add_text(sl, title, 0.9, 0.5, 11, 0.6, C['font_h'], 28, C['fg'], True)
    _add_rect(sl, 0.9, 1.2, 2, 0.04, C['accent'])
    for i, (val, label, fw, color) in enumerate(items):
        yy = 1.8 + i * 1.7
        _add_rounded_rect(sl, 0.9, yy, fw, 1.3, C['card_bg'], color)
        _add_text(sl, val, 1.2, yy + 0.15, 3, 0.6, C['font_h'], 32, color, True)
        _add_text(sl, label, 1.2, yy + 0.75, fw - 0.6, 0.4, C['font_b'], 13, C['muted'])
```

#### page_swot — SWOT 页

```python
def page_swot(prs, title, quadrants):
    """SWOT 页：2×2 彩色网格
    quadrants: [(letter, title, items, color, x, y), ...]
    """
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _add_rect(sl, 0, 0, 0.06, 7.5, C['primary'])
    _add_rect(sl, 0, 7.25, 13.333, 0.25, C['card_bg'])
    _add_text(sl, title, 0.9, 0.5, 11, 0.6, C['font_h'], 28, C['fg'], True)
    _add_rect(sl, 0.9, 1.2, 2, 0.04, C['accent'])
    for letter, qtitle, items, color, xx, yy in quadrants:
        _add_rounded_rect(sl, xx, yy, 5.4, 2.6, C['card_bg'], C['card_bd'])
        _add_text(sl, letter, xx + 0.2, yy + 0.15, 0.5, 0.5, C['font_h'], 22, color, True)
        _add_text(sl, qtitle, xx + 0.8, yy + 0.2, 3.5, 0.4, C['font_h'], 16, C['fg'], True)
        _add_multiline(sl, [f'•  {it}' for it in items], xx + 0.2, yy + 0.8, 4.8, 1.5, C['font_b'], 11, C['muted'])
```

#### page_pyramid — 金字塔页

```python
def page_pyramid(prs, title, subtitle, layers, legend=None):
    """金字塔页：递增宽度层 + 可选侧栏
    layers: [(number, label, width, color), ...]
    legend: [(keyword, description), ...] or None
    """
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _add_rect(sl, 0, 0, 0.06, 7.5, C['primary'])
    _add_rect(sl, 0, 7.25, 13.333, 0.25, C['card_bg'])
    _add_text(sl, title, 0.9, 0.5, 11, 0.6, C['font_h'], 28, C['fg'], True)
    _add_rect(sl, 0.9, 1.2, 2, 0.04, C['accent'])
    _add_text(sl, subtitle, 0.9, 1.5, 10, 0.5, C['font_b'], 14, C['muted'])
    for i, (num, label, width, color) in enumerate(layers):
        yy = 2.3 + i * 1.2
        x = 0.9 + (7.5 - width) / 2
        _add_rounded_rect(sl, x, yy, width, 0.9, C['card_bg'], color)
        _add_text(sl, num, x + 0.3, yy + 0.1, 1, 0.5, C['font_h'], 28, color, True)
        _add_text(sl, label, x + 1.5, yy + 0.2, 3, 0.4, C['font_b'], 16, C['fg'])
    if legend:
        _add_rounded_rect(sl, 9.0, 2.3, 3.8, 5.0, C['card_bg'], C['card_bd'])
        _add_text(sl, '风格关键词', 9.3, 2.5, 3, 0.4, C['font_h'], 14, C['accent'], True)
        _add_multiline(sl, [f'•  {k} → {d}' for k, d in legend], 9.2, 3.1, 3.3, 3.5, C['font_b'], 10, C['muted'], spacing=4)
```

#### page_table — 表格页

```python
def page_table(prs, title, headers, rows, col_widths, note=None):
    """表格页：表头 + 交替色行
    headers: [str, ...]
    rows: [[str, ...], ...]
    col_widths: [float, ...]
    """
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _add_rect(sl, 0, 0, 0.06, 7.5, C['primary'])
    _add_rect(sl, 0, 7.25, 13.333, 0.25, C['card_bg'])
    _add_text(sl, title, 0.9, 0.5, 11, 0.6, C['font_h'], 28, C['fg'], True)
    _add_rect(sl, 0.9, 1.2, 2, 0.04, C['accent'])
    col_x = [0.9]
    for w in col_widths[:-1]:
        col_x.append(col_x[-1] + w + 0.1)
    for j, h in enumerate(headers):
        _add_rounded_rect(sl, col_x[j], 1.6, col_widths[j], 0.5, C['primary'])
        _add_text(sl, h, col_x[j] + 0.1, 1.65, col_widths[j] - 0.2, 0.4, C['font_h'], 12, C.get('white', '#FFFFFF'), True, 'center')
    for i, row in enumerate(rows):
        yy = 2.2 + i * 0.8
        bg = C['card_bg'] if i % 2 == 0 else C['bg']
        for j, cell in enumerate(row):
            _add_rounded_rect(sl, col_x[j], yy, col_widths[j], 0.65, bg, C['card_bd'])
            _add_text(sl, cell, col_x[j] + 0.1, yy + 0.1, col_widths[j] - 0.2, 0.4, C['font_b'], 11, C['fg'] if j == 0 else C['muted'], j == 0, 'center' if j == 0 else 'left')
    if note:
        _add_text(sl, note, 0.9, 6.5, 11, 0.4, C['font_b'], 10, C['accent'])
```

#### page_cta — CTA 页

```python
def page_cta(prs, title, subtitle, envs, code_lines, image_path=None):
    """CTA 页：全幅图+安装路径+代码块
    envs: [(env_name, path), ...]
    code_lines: [str, ...]
    """
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    if image_path:
        _add_image(sl, image_path, 0, 0, 13.333, 7.5)
        _dark_overlay(sl, 0.7)
    else:
        _add_rect(sl, 0, 0, 13.333, 7.5, C['bg'])
    _add_text(sl, title, 1.2, 0.8, 6, 1.0, C['font_h'], 44, C['fg'], True)
    _add_text(sl, subtitle, 1.2, 1.8, 6, 0.5, C['font_b'], 16, C['muted'])
    for i, (env, path) in enumerate(envs):
        yy = 2.6 + i * 1.1
        _add_rounded_rect(sl, 1.2, yy, 5.5, 0.85, C['card_bg'], C['card_bd'])
        _add_text(sl, env, 1.5, yy + 0.1, 2.5, 0.35, C['font_h'], 14, C['accent'], True)
        _add_text(sl, path, 1.5, yy + 0.45, 4.8, 0.3, C['font_b'], 10, C['muted'])
    _add_rounded_rect(sl, 7.5, 2.6, 4.8, 3.5, '#1E1E2E', C['primary'])
    _add_text(sl, 'quick start', 7.8, 2.8, 2, 0.3, C['font_b'], 9, C['dim'], True)
    _add_multiline(sl, code_lines, 7.8, 3.2, 4.2, 2.5, C['font_b'], 11, '#CDD6F4', spacing=3)
```

### 完整项目示例

```python
"""build.py — PPT-Design-Skill 产品介绍（14 页）"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image as PILImage
import tempfile, hashlib

# ── Design Token ──
C = {
    'bg': '#060B18', 'primary': '#6366F1', 'accent': '#22D3EE',
    'fg': '#F8FAFC', 'muted': '#64748B', 'dim': '#334155',
    'card_bg': '#0D152A', 'card_bd': '#1A2A4A',
    'font_h': 'Space Grotesk', 'font_b': 'Consolas',
}

# ... (工具函数同上) ...

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

page_cover(prs, "PPT-Design-Skill", "AI 驱动的 PPT 生成引擎",
           "40,000+ 风格组合 · 一句话生成专业演示文稿",
           [("40K+", "风格组合"), ("10", "图形类型"), ("5", "AI引擎")],
           image_path="docs/showcase/showcase-dark-tech-slide1.jpg")

page_numbered_list(prs, "传统 PPT 制作的三大痛点",
    [("手动排版耗时", "一份专业 PPT 平均需要 8 小时"),
     ("风格不一致", "团队协作难以统一视觉标准"),
     ("图形绘制困难", "流程图、漏斗图全靠手动拼凑")],
    side_image="docs/showcase/showcase-professional-slide2.jpg")

# ... 更多页面 ...

prs.save("output/presentation.pptx")
```

### 生成后检查脚本

```python
"""check.py — 生成后自动检查"""
from pptx import Presentation

prs = Presentation("output/presentation.pptx")
errors = []
for i, slide in enumerate(prs.slides, 1):
    for s in slide.shapes:
        if s.has_text_frame:
            for p in s.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.name is None:
                        errors.append(f"Slide {i}: run-level font.name is None")
                    if r.font.size is None:
                        errors.append(f"Slide {i}: run-level font.size is None")
if errors:
    print(f"❌ {len(errors)} issues found:")
    for e in errors[:10]:
        print(f"  {e}")
else:
    print(f"✅ All {len(prs.slides)} slides passed font checks")
```

---

## 6. content.json 内容格式

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
      "goal": "section",
      "title": "Market Analysis",
      "section_number": "01"
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
| `goal` | str | 页面目标，默认 `"content"`。hook/cta 获得全屏品牌色覆盖；`section` 渲染章节分隔页（超大编号+标题+渐变线） |
| `title` | str | 页面标题 |
| `subtitle` | str | 副标题 |
| `section_number` | str | 章节编号（配合 `goal="section"` 使用，如 `"01"`、`"02"`） |
| `bullets` | list[str] | 要点列表 |
| `image` | str | 图片路径（相对项目目录或绝对路径） |
| `cards` | list[dict] | 卡片列表，每张含 `title`/`text`/可选 `image` |
| `diagram` | dict | 图形定义，含 `type` 和 `data`（见第 8 节） |
| `code` | str 或 dict | 代码块：字符串或 `{"source": "...", "language": "python"}` |
| `exercise` | dict | `{"instructions": str, "duration": str, "steps": list[str]}` |
| `notes` | str | 演讲者备注（v0.4.0） |
| `links` | list[dict] | 超链接（v0.4.0） |
| `chart` | dict | 数据图表（v0.4.0） |
| `component_type` | str | 组件类型：`"group"` 或 `"smartart"`，触发组件库匹配注入 |
| `component_category` | str | 组件类别：`process`/`hierarchy`/`infographic`/`swot`/`timeline`/`chart` |
| `component_variant` | str | 可选，组件变体名 |

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

## 7. brand.json 品牌格式

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

## 8. 组件库数据库

组件库是从 PPT 素材中提取的图表模板数据库，支持在生成 PPT 时自动匹配和注入专业图表组件（组织架构图、流程图、时间线、SWOT 等），无需从零绘制。

### 数据库概览

| 属性 | 值 |
|------|---|
| 存储位置 | `component_library/index.db`（SQLite）+ `component_library/storage/`（XML 文件） |
| 当前规模 | 5,560 组件（5,537 group + 23 smartart） |
| 素材来源 | DEF 目录 124 个 PPT + ABC 目录 7 个 PPT |
| 压缩方式 | gzip（6.5× 压缩率） |
| 去重机制 | MD5 checksum（type+category+variant+node_count+XML 前 4KB） |

### 组件分类

| 类别 | 数量 | 说明 |
|------|------|------|
| infographic | 4,101 | 信息图、数据展示 |
| process | 672 | 流程图、步骤图 |
| hierarchy | 548 | 组织架构、层级图 |
| chart | 132 | 数据图表 |
| timeline | 42 | 时间线、里程碑 |
| swot | 39 | SWOT 分析矩阵 |
| features | 2 | 功能特性展示 |
| comparison | 1 | 对比图 |

### 构建数据库

从 PPT 素材目录批量提取组件并入库：

```bash
# 基本用法
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材目录"

# 指定最低节点数（推荐 3，过滤装饰性碎片）
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材目录" --min-node-count 3

# 多个目录依次导入（去重自动跳过已有组件）
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材目录A"
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材目录B"
```

**关键参数**：

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--materials-dir` | 必填 | PPT 素材目录路径（支持子目录递归扫描） |
| `--min-node-count` | 3 | 最低节点数阈值，低于此值的 GroupShape 被过滤 |

**构建输出示例**：

```
[INFO] Found 124 PPT files in E:\素材目录
[INFO] [10/124] added=308 skip=1045 err=0 time=28.1s total=289
[INFO] [20/124] added=488 skip=1511 err=0 time=17.8s total=769
...
[INFO] === BUILD COMPLETE ===
[INFO] Components added: 5341
[INFO] Skipped (dedup/empty): 15149
[INFO] Errors: 0
[INFO] Total time: 139.7s (1.1s/file)
```

### min_node_count 阈值选择

| 阈值 | 效果 | 适用场景 |
|------|------|---------|
| 1 | 包含所有组件（含单个标签、图标+文字对） | 需要最大素材量 |
| 2 | 排除纯装饰单节点，保留标签对 | 需要简单组件 |
| **3（推荐）** | 只保留 3+ 节点的图表级组件 | **质量优先，日常使用** |

**为什么推荐 3**：node_count < 3 的组件本质上是装饰碎片（单个标签、按钮），PrecisionRenderer 已能用 `add_text()`、`add_card()` 等方法从零生成且自动符合品牌规范。组件库的核心价值是复杂图表模板，简单元素从零构建反而更好（品牌一致性、坐标精确）。

### 数据库操作要领

#### 查看数据库状态

```python
from ppt_pro_max.enterprise.component_library import ComponentLibrary

lib = ComponentLibrary("component_library/index.db")

# 总体统计
print(lib.stats())
# {'total': 5560, 'group': 5537, 'smartart': 23}

# 分类目录
catalog = lib.catalog()
for comp_type, categories in catalog.items():
    for cat, info in categories.items():
        print(f"  {comp_type}/{cat}: {info['count']} components, nodes {info['min_nodes']}-{info['max_nodes']}")

# 某类目覆盖度
print(lib.coverage("process"))
# {'3': 45, '4': 120, '5': 89, ...}

lib.close()
```

#### 搜索组件

```python
lib = ComponentLibrary("component_library/index.db")

# 按类型+类别搜索
results = lib.search(type="group", category="process", limit=10)

# 指定节点数精确匹配
results = lib.search(type="group", category="hierarchy", node_count=5, limit=5)

# 带标签过滤
results = lib.search(type="group", category="infographic", tags=["circular"], limit=10)

lib.close()
```

#### 智能匹配

```python
# 根据提取数据自动匹配最合适的组件
element = {"type": "group", "category": "process", "node_count": 4}
matched = lib.match(element)
# 优先精确匹配 node_count，无结果则选最接近的
```

#### 加载组件 XML

```python
# 加载组件的 XML 模板（gzip 自动解压）
xml_parts = lib.load_xml(matched["id"])
# xml_parts = {"group": b"<p:grpSp>...", "img_rId2": b"\x89PNG..."}

# 加载缩略图
thumb = lib.load_thumbnail(matched["id"])
```

#### 删除组件

```python
# 删除单个组件（同时删除存储文件）
lib.remove(component_id)

# 清空整个数据库（删除 component_library/ 目录）
import shutil
shutil.rmtree("component_library", ignore_errors=True)
```

#### 重建数据库

当需要从零重建（如坐标归一化逻辑更新后）：

```bash
# 1. 删除旧数据库
rmdir /s /q component_library

# 2. 重新构建
python -m ppt_pro_max.scripts.build_library --materials-dir "E:\素材目录" --min-node-count 3
```

> **重要**：修改了 `group_extractor.py` 的坐标归一化逻辑后，必须重建数据库。旧数据库中的 XML 是非归一化的，无法正确缩放注入。

### 公开 API 查询

无需直接操作数据库，通过公开 API 查询：

```python
from ppt_pro_max import query_component_library

# 查看完整目录（无参数）
catalog = query_component_library()
# {"group": {"infographic": {"count": 4101, ...}, ...}, ...}

# 搜索组件
results = query_component_library(type="group", category="process", node_count=4)
# [{"id": 123, "type": "group", "category": "process", ...}, ...]

# 指定数据库路径
results = query_component_library(type="group", category="swot", component_library="path/to/index.db")
```

### 在 content.json 中使用组件

在 slides 中指定 `component_type` 和 `component_category`，Pipeline 自动从组件库匹配并注入：

```json
{
  "slides": [
    {
      "goal": "content",
      "title": "项目流程",
      "bullets": ["需求分析", "方案设计", "开发实现", "测试上线"],
      "component_type": "group",
      "component_category": "process"
    },
    {
      "goal": "content",
      "title": "组织架构",
      "bullets": ["CEO", "CTO", "CFO", "VP Engineering"],
      "component_type": "group",
      "component_category": "hierarchy"
    }
  ]
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `component_type` | str | 组件类型：`"group"` 或 `"smartart"` |
| `component_category` | str | 组件类别：`process`/`hierarchy`/`infographic`/`swot`/`timeline`/`chart`/`comparison`/`features` |
| `component_variant` | str | 可选，组件变体名 |

Pipeline 匹配逻辑：
1. 精确匹配 `type + category + node_count`（node_count 由 bullets 数量推断）
2. 无精确匹配时，选同 type+category 中 node_count 最接近的
3. 匹配失败时，回退到 DiagramEngine 或 bullets 列表渲染

### 坐标归一化原理

组件库中的所有坐标已归一化到 0-1 范围，注入时按目标尺寸反归一化：

**提取时（归一化）**：
```
rel = (val - parent_chOff) / parent_chExt    # off/chOff
rel = val / parent_chExt                       # ext/chExt
```

**注入时（反归一化）**：
```
emu = rel × target_emu + target_offset_emu    # 位置
emu = rel × target_emu                         # 尺寸
sz_hundredths = rel_sz × h_emu, clamp [800, 7200]  # 字号（8pt-72pt）
```

每个 GroupShape 使用自身的 chOff/chExt 归一化子元素，嵌套组递归处理。这确保组件可以在任意尺寸的目标区域正确渲染。

### 数据库路径查找顺序

Pipeline 按以下顺序查找数据库：

1. 用户指定路径（`component_library` 参数）
2. 项目目录下 `component_library/index.db`
3. 包目录及上级目录下 `component_library/index.db`

### Beautify 模式与组件库

Beautify full 模式自动使用组件库：

```bash
# full 模式：提取内容 → 推断 component_type → 匹配组件库 → 重建
ppt-design "" --beautify client.pptx --style professional

# light 模式：仅替换颜色/字体/背景，不使用组件库
ppt-design "" --beautify client.pptx --style professional --beautify-mode light
```

---

## 9. 页面修订语法

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

## 10. 10 种图形类型

### 通用格式

```json
{
  "diagram": {
    "type": "<图形类型>",
    "data": { <类型专用数据> }
  }
}
```

### 9.1 流程图 (flowchart)

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

### 9.2 漏斗图 (funnel)

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

### 9.3 时间线 (timeline)

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

### 9.4 SWOT 分析 (swot)

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

### 9.5 矩阵图 (matrix)

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

### 9.6 循环图 (cycle)

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

### 9.7 表格 (table)

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

### 9.8 层级图 (hierarchy)

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

### 9.9 金字塔图 (pyramid)

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

### 9.10 韦恩图 (venn)

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

## 11. 设计系统 — 40,000+ 风格组合

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
| international/global | slate-minimal + modern-sans + accent-bar |
| cream/vintage | golden-luxury + serif-editorial + gold-trim |
| frosted/glass | arctic-frost + geometric-sans + gradient-bar |
| mckinsey/consulting | midnight-navy + clean-corporate + accent-bar |
| pastel/soft | lavender-dream + friendly-round + circle-accent |
| retro/classic | terracotta + literary-serif + diamond-bullets |
| government/official | midnight-navy + classic-formal + accent-bar |
| legal/law | midnight-navy + classic-formal + sidebar-nav |
| pharma/medical | teal-fresh + clean-corporate + minimal-dots |
| realestate/property | sunset-warm + warm-mix + gold-trim |
| automotive/car | charcoal-bold + sharp-modern + neon-lines |
| aviation/aerospace | arctic-frost + geometric-sans + gradient-bar |
| energy/power | copper-industrial + bold-sans + gradient-bar |
| telecom/5g | cyber-neon + tech-mono + neon-lines |
| logistics/supply | slate-minimal + clean-corporate + sidebar-nav |

---

## 12. 设计质量升级 — 28 项专业级提升

v0.7.0 引入 28 项设计质量升级，分三个层级全面提升 PPT 视觉品质。所有升级已集成到 Pipeline，无需额外配置即可生效。

### Tier 1 — 基础视觉（10 项）

| # | 升级 | 说明 | 模块 |
|---|------|------|------|
| 1.1 | **布局引擎** | `LayoutEngine` + `Rect` + `ContentLayout`，统一坐标计算，消除硬编码 | `renderer/layout_engine.py` |
| 1.2 | **排版比例尺** | `TypeScale` 数据类，密度/模式自适应字号比例（presenting/reading/balanced） | `renderer/typography.py` |
| 1.3 | **OKLCH 色彩深度** | 色彩感知均匀的 9 级色阶 `generate_color_scale()` + 5 级 alpha 层级 `ALPHA_LEVELS` | `renderer/color_system.py` |
| 1.4 | **渐变叠层** | `add_gradient_overlay()` 取代纯色遮罩，封面/CTA 页面更柔和自然 | `precision_renderer.py` |
| 1.5 | **5 级阴影层级** | `ELEVATION_SCALE` 从 subtle→floating，卡片/徽章层次分明 | `renderer/elevation.py` |
| 1.6 | **品牌条智能省略** | 约 1/3 页面省略品牌色条，避免视觉疲劳；有进度条时自动跳过底部条 | `precision_renderer.py` |
| 1.7 | **图片调色** | `grade_image_to_palette()` 将图片色调统一到品牌色系，Pillow blend + 缓存 | `renderer/image_processor.py` |
| 1.8 | **卡片升级** | 标题 20pt / 正文 14pt，`featured=True` 卡片带渐变条（22pt 标题） | `precision_renderer.py` |
| 1.9 | **暗色模式修正** | OKLCH 亮度判定 `0.299R+0.587G+0.114B`，告别 7-hex 硬编码误判 | `precision_renderer.py` |
| 1.10 | **代码块重设计** | 始终深色背景 #1E293B + 独立圆角语言徽章 + muted-foreground 文字色 | `precision_renderer.py` |

### Tier 2 — 排版增强（6 项）

| # | 升级 | 说明 | 模块 |
|---|------|------|------|
| 2.1 | **CJK 字体配对** | 12 种中英文字体组合（Space Grotesk+YaHei、Georgia+STSong 等），`add_text()`/`add_multiline()` 自动回退 | `renderer/theme_mapper.py` |
| 2.2 | **自适应边距** | presenting（0.6"）/ reading（1.2"）/ balanced（0.9"）三种密度对应不同边距 | `renderer/layout_engine.py` |
| 2.3 | **徽章系统** | `add_badge()` — 全大写文字、3 种变体（default/solid/outline）、CJK 宽度自适应 | `precision_renderer.py` |
| 2.4 | **章节分隔页** | `goal="section"` → 超大编号 + 标题 + 渐变线，清晰分章 | `precision_renderer.py` |
| 2.5 | **装饰渲染器** | `DecorationRenderer` 统一渲染 10 种装饰风格，连接到 `render_slide()` | `renderer/decoration_renderer.py` |
| 2.6 | **布局变体消费** | `render_slide()` 读取 `content_margin_left` / `title_alignment` / `decoration_style` | `precision_renderer.py` |

### Tier 3 — 高级视觉（7 项）

| # | 升级 | 说明 | 模块 |
|---|------|------|------|
| 3.1 | **噪点纹理** | 每套 PPT 独立种子的高斯噪点 `generate_noise_tile()`，增添质感 | `renderer/image_processor.py` |
| 3.2 | **进度条** | `add_progress_bar()` — 细条底部进度指示，替代简单页码 | `precision_renderer.py` |
| 3.3 | **圆角系统** | 4 级圆角 `CORNER_RADIUS_SCALE`（sm=4 / md=8 / lg=12 / pill=50）+ `add_rounded_rect()` | `precision_renderer.py` |
| 3.4 | **渐变线** | `add_gradient_line()` — alpha 渐隐，标题下划线更精致 | `precision_renderer.py` |
| 3.5 | **图片遮罩** | `add_masked_image()` — 圆角矩形框 + 0.15" 内边距，侧边图更精致 | `precision_renderer.py` |
| 3.6 | **双栏要点** | 6+ 条要点自动分两栏 + 垂直分隔线，信息密度更高 | `precision_renderer.py` |
| 3.7 | **Hero 4 变体** | `_select_hero_pattern()` — gradient / split-left / bottom-fade / asymmetric 4 种封面布局 | `precision_renderer.py` |

### 使用方式

所有 28 项升级**默认启用**，无需额外配置。部分升级可通过参数控制：

```python
from ppt_pro_max import generate_ppt

# 章节分隔页 — 在 content.json 中设置 goal="section"
result = generate_ppt("课程大纲", project="./course", style="professional")

# 进度条 — 自动根据 total_pages 显示
# Pipeline 内部传递 page_index 和 total_pages 给 render_slide()

# 卡片 featured 模式 — 在 content.json cards 中设置
# {"title": "核心功能", "text": "...", "featured": true}
```

### content.json 新增字段

| 字段 | 类型 | 说明 |
|------|------|------|
| `section_number` | str | 章节编号，配合 `goal="section"` 使用（如 `"01"`、`"02"`） |
| `featured` | bool | 卡片高亮模式，标题 22pt + 渐变条（在 `cards` 中使用） |

### CJK 字体配对表

| 英文字体 | CJK 回退字体 | 适用场景 |
|---------|-------------|---------|
| Space Grotesk | Microsoft YaHei | 科技/现代 |
| Inter | Microsoft YaHei | 通用/简洁 |
| Poppins | Microsoft YaHei | 活力/创业 |
| Calibri | Microsoft YaHei | 商务/企业 |
| Georgia | STSong | 编辑/传统 |
| Playfair Display | STSong | 优雅/奢侈 |
| Lora | STSong | 文学/人文 |
| Consolas | SimHei | 技术/代码 |
| Segoe UI | Microsoft YaHei | 人文/友好 |
| Nunito | Microsoft YaHei | 圆润/友好 |
| Montserrat | Microsoft YaHei | 锐利/现代 |
| Times New Roman | SimSun | 经典/正式 |

### 圆角规格

| 级别 | 值 (pt) | 适用场景 |
|------|---------|---------|
| `sm` | 4 | 小型标签、徽章 |
| `md` | 8 | 卡片、代码块 |
| `lg` | 12 | 大型面板、图片框 |
| `pill` | 50 | 胶囊按钮、全圆角 |

### 阴影层级

| 级别 | blur (pt) | distance (pt) | 适用场景 |
|------|-----------|---------------|---------|
| `subtle` | 2 | 1 | 徽章、小卡片 |
| `small` | 4 | 2 | 标准卡片 |
| `medium` | 8 | 4 | 浮动面板 |
| `large` | 16 | 8 | 弹出层 |
| `floating` | 24 | 12 | 模态框 |

---

## 13. 动画系统

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

## 14. 密度控制

`--density` 1-10 控制每页信息量：

| 等级 | 标题字号 | 正文字号 | 最大要点数 | 最大字符数 | 图片比例 |
|------|---------|---------|-----------|-----------|---------|
| 1 | 36pt | 16pt | 3 | 40 | 50% |
| 2 | 34pt | 15pt | 4 | 50 | 45% |
| 3 | 32pt | 14pt | 5 | 60 | 40% |
| 4 | 30pt | 13pt | 5 | 70 | 38% |
| 5 | 28pt | 12pt | 6 | 80 | 35% |
| 6 | 26pt | 11pt | 7 | 90 | 32% |
| 7 | 24pt | 11pt | 8 | 100 | 30% |
| 8 | 22pt | 11pt | 10 | 120 | 25% |
| 9 | 20pt | 11pt | 12 | 140 | 22% |
| 10 | 18pt | 11pt | 15 | 160 | 20% |

- 超长要点自动截断并加省略号
- 超量要点自动丢弃
- 图片宽度按比例缩放

---

## 15. 图片引擎

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

## 16. Python API

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

## 17. CLI 完整参数

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
  --component-library PATH  组件库数据库路径（默认自动查找）

美化选项:
  --beautify PPTX        美化现有 PPT 文件
  --beautify-mode MODE   light（仅换色）或 full（重建+组件注入）
```

---

## 18. 最佳实践与质量保证

### 工作流选择

```
需求明确度低 ─────→ FreeStyle（快速探索）
需求明确度高 ─────→ Enterprise（团队协作）
交付质量要求高 ──→ Build Script（精确控制）
```

### 三阶段交付流程

**阶段一：FreeStyle 原型**

用 FreeStyle 快速生成 3-5 种风格方案，确定方向：

```bash
ppt-design "融资路演" --style "dark cyberpunk" --motion 0
ppt-design "融资路演" --style "warm fintech" --motion 0
ppt-design "融资路演" --style "elegant luxury" --motion 0
```

**阶段二：Enterprise 内容填充**

确定风格后，用 content.json 填充真实内容：

```bash
ppt-design "融资路演" --project ./my-project --style "dark cyberpunk" --density 6
```

**阶段三：Build Script 精细交付**

从 Enterprise 输出中提取内容，改写为 build.py 逐页精确控制：

```bash
python build.py && python check.py
```

### Build Script 五条铁律

#### 铁律一：字体必须设 run-level

```python
# ❌ PowerPoint 忽略 paragraph 级别
p.font.size = Pt(20)

# ✅ 只在 run 级别设置
run.font.size = Pt(20)
run.font.color.rgb = _rgb('#F8FAFC')
run.font.name = 'Space Grotesk'
run.font.bold = True
```

**原因**：`p.font.size` 在 XML 中写入 `<a:defRPr>`，PowerPoint 打开时用 `<a:rPr>` 覆盖。只有 `run.font.*` 写入 `<a:rPr>` 才能保证生效。

#### 铁律二：图片用 Pillow 预裁剪，绝不拉伸

```python
# ❌ add_picture 拉伸变形
slide.shapes.add_picture(path, left, top, width, height)

# ✅ Pillow center-crop 到目标比例后再插入
img = PILImage.open(path)
# ... center crop logic ...
cropped.save(temp_path, 'PNG')
slide.shapes.add_picture(temp_path, left, top, width, height)
```

#### 铁律三：全幅图+渐变叠层是万能公式

封面页和案例展示页的标准模式：

```python
_add_image(sl, path, 0, 0, 13.333, 7.5)   # 全幅背景
add_gradient_overlay(sl, 0.65)              # 65% 渐变叠层（v0.7.0+，替代纯色遮罩）
_add_text(sl, "标题", 0.8, 2.0, 6, 1.2, ...) # 文字叠在上面
```

叠层透明度参考：
- 封面页：0.70-0.75（深叠层，突出文字）
- 案例展示页：0.50-0.60（浅叠层，保留图片细节）
- CTA 页：0.65-0.70

#### 铁律四：Design Token 统一管理

```python
# ✅ 一处修改，全局生效
C = TOKENS['cyber-neon']  # 切换主题只改这一行

# ❌ 硬编码颜色散落各处
_add_text(sl, "标题", ..., color='#6366F1')  # 不要这样写
```

#### 铁律五：生成后必须检查

每次 build 后跑检查脚本，验证：

1. **元素数量**：每页 shapes 数量是否符合预期
2. **字体**：所有 run-level font.name 和 font.size 不为 None
3. **图片尺寸**：全幅图 13.3"×7.5"，卡片内图按比例
4. **文字溢出**：textbox 的 w/h 是否足够容纳内容
5. **视觉审查**：在 PowerPoint 中打开逐页检查

### 内容页布局模式

内容页（非全幅图页）的标准结构：

```
┌──────────────────────────────────────────────────┐
│▌                                                 │ ← 左侧 accent 竖条 (0.06" 宽, ~1/3 页省略)
│  标题 (28pt Space Grotesk, foreground)            │ ← x=0.9, y=0.5
│  ━━━━━━━━━━                                      │ ← 渐变下划线 (alpha 渐隐, add_gradient_line)
│                                                  │
│  ┌─────────┐  ┌─────────┐  ┌─────────┐          │ ← 圆角卡片 (md=8pt, add_rounded_rect)
│  │ 01      │  │ 02      │  │ 03      │          │ ← 卡片标题 20pt / 正文 14pt
│  │ 标题    │  │ 标题    │  │ 标题    │          │ ← featured 卡片: 22pt + 渐变条
│  │ 说明    │  │ 说明    │  │ 说明    │          │
│  └─────────┘  └─────────┘  └─────────┘          │ ← 阴影层级: small (blur=4, dist=2)
│                                                  │
│  ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓│ ← 进度条 (细条, 替代简单页码)
└──────────────────────────────────────────────────┘
```

关键间距参数：

| 元素 | 参数 | 值 |
|------|------|---|
| 左边距 | x | 0.9"（balanced）/ 0.6"（presenting）/ 1.2"（reading） |
| accent 竖条宽度 | w | 0.06"（~1/3 页省略） |
| 标题下划线 | — | 渐变线，alpha 渐隐（`add_gradient_line`） |
| 进度条 | — | 底部细条，替代 muted 横条 |
| 标题字号 | size | 28pt（TypeScale 自适应） |
| 标题字体 | font | font_h + CJK 回退自动配对 |
| 卡片圆角 | corner_radius | md=8pt（`add_rounded_rect`） |
| 卡片阴影 | — | small 级（blur=4, dist=2, `apply_elevation`） |
| 卡片间距 | gap | 0.1-0.4" |

### 暗色主题配色方案

Build Script 模式下推荐使用以下暗色主题 Token：

| Token 名 | 适用场景 | 特点 |
|-----------|---------|------|
| `cyber-neon` | 科技/产品/开发者 | 蓝紫霓虹，Consolas 等宽字体 |
| `midnight-navy` | 商务/融资/报告 | 深蓝稳重，Calibri 字体 |
| `charcoal-bold` | 极简/设计/品牌 | 灰黑极简，Inter 字体 |
| `wine-burgundy` | 奢侈/品牌/策略 | 深红金调，Georgia 字体 |

### 常见问题

**Q: 为什么 PowerPoint 中字体显示为"宋体"？**

A: 字体设在了 paragraph 级别而非 run 级别。必须使用 `run.font.name = "字体名"` 而非 `p.font.name = "字体名"`。

**Q: 图片变形了怎么办？**

A: 使用 `_add_image()` 函数（内置 Pillow 预裁剪），不要直接用 `slide.shapes.add_picture()` 传入不同比例的图片。

**Q: 暗色遮罩下文字看不清？**

A: v0.7.0 起使用渐变叠层 `add_gradient_overlay()` 替代纯色遮罩，过渡更自然。如需手动调整，`_dark_overlay(sl, opacity)` 的 opacity 参数：0.5 = 半透明（图片更清晰），0.8 = 几乎不透明（文字更清晰）。推荐 0.55-0.75。

**Q: CJK 文字（中文/日文/韩文）字体显示异常？**

A: v0.7.0 起内置 12 种 CJK 字体配对（如 Space Grotesk + Microsoft YaHei、Georgia + STSong 等），`add_text()` / `add_multiline()` 自动检测并设置东亚字体回退。无需手动处理。

**Q: Build Script 和 Enterprise 怎么选？**

A: Enterprise 适合团队协作、版本管理、品牌合规。Build Script 适合最终交付、精确控制。两者可以组合使用：先用 Enterprise 生成内容和品牌方案，再用 Build Script 精细化调整。
