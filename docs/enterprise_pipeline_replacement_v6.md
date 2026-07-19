# Enterprise Pipeline 剥离方案 v6：LLM-Skill 路径

> 版本：v6.1（含审查修正） | 日期：2026-07-19
> 核心认知：这是一个 LLM Skill，不是传统软件系统。LLM 是 Bridge，不是代码。

---

## 一、核心认知修正

### 1.1 之前的错误

v3-v5 方案都在用传统软件工程思路：

```
错误路径：脚本提取 → 数据结构(VISpec) → 代码框架(VIBuilder) → 自动渲染
```

试图用代码自动化一切：颜色角色映射、shape 分类、布局选择、装饰复制……
结果：2,420 行新代码 + 大量重构 + 23-28 天工期。

### 1.2 正确路径

这是一个 **LLM Skill**，核心流程是：

```
LLM 调用分析脚本 → LLM 读取文本输出 → LLM 分析设计 → LLM 写 build 脚本 → 输出
```

**LLM 就是 Bridge。** 它读色值输出，自己判断哪个是 primary 哪个是 accent；它看 shape 列表，自己决定复制哪些装饰；它理解用户需求，自己设计页面结构和布局。

### 1.3 这意味着什么

| 传统思路 | Skill 思路 |
|----------|-----------|
| 写 VISpec 数据结构 | 不需要 — LLM 直接读文本 |
| 写颜色角色映射算法 | 不需要 — LLM 自己分配 |
| 写 shape→kind 分类器 | 不需要 — LLM 看输出自己判断 |
| 写 VIBuilder 框架 | 不需要 — LLM 自己写 build 脚本 |
| 写 DNAStyleBridge | 不需要 — LLM 就是 Bridge |
| 保留/重构 Enterprise 模块 | 不需要 — 直接删除 |

**需要做的只有两件事：**
1. 提供好的**分析脚本**，让 LLM 能读懂模板
2. 提供好的**Build 辅助库**，让 LLM 写脚本时有工具可用

---

## 二、实际工作流

### 2.1 用户场景

用户有一个 PPT 模板，想基于模板的 VI 制作新 PPT。

### 2.2 LLM 执行流程

```
Step 1: LLM 调用分析脚本
    python analyze_template.py template.pptx
    ↓
    输出：页面尺寸、每页 shape 列表、色值表、字体表、theme colors

Step 2: LLM 读取输出，分析模板 VI
    "这个模板用绿色系，primary=#2E6504, accent=#7DA92F..."
    "封面有渐变大字+胶囊条+连接线装饰..."
    "正文页有顶部绿色条+logo+标题+分隔线..."

Step 3: LLM 根据用户需求设计 PPT 框架
    "用户要乡村振兴数据报告，需要8页：封面+目录+4个数据页+对比页+封底"
    "数据页用 KPI card + bar chart + donut chart..."

Step 4: LLM 写 build 脚本
    - 读取模板（保留原有页面）
    - 定义色板 C={}（从 Step 2 分析结果填入）
    - 定义辅助函数（从 build_helpers 库导入）
    - 逐页 Build 新页面
    - 保存

Step 5: LLM 调用脚本生成 PPT
    python build_xiangcun.py
    ↓
    输出 PPT
```

### 2.3 与现有 Build 模式的关系

这就是 **Build 模式**，只是多了"先分析模板 VI"这一步。现有 Skill 定义中 Build 模式的流程完全适用，只需要在 Step 2 之前加一个模板分析步骤。

---

## 三、分析脚本

### 3.1 现有脚本

`docs/分析脚本/` 中已有 4 个脚本：

| 脚本 | 功能 | 状态 |
|------|------|------|
| `inspect_template.py` | 输出页面尺寸、shape 列表、文字/字号/颜色 | 可用，需微调输出格式 |
| `extract_colors.py` | 输出所有色值 + theme colors | 可用，需微调输出格式 |
| `build_data_analysis.py` | Build 示例脚本 | 参考模板，不是分析脚本 |
| `verify_output.py` | 验证输出 PPT | 可用 |

### 3.2 需要的改进

**唯一需要做的：合并 inspect_template.py + extract_colors.py 为一个脚本，优化输出格式让 LLM 更容易读取。**

当前输出是散乱的 print，LLM 需要自己拼凑信息。改进为结构化输出：

```python
# analyze_template.py — 合并版
"""
用法：python analyze_template.py template.pptx
输出：结构化文本，供 LLM 读取分析
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from lxml import etree
from collections import Counter

def analyze(pptx_path):
    prs = Presentation(pptx_path)
    W_in = prs.slide_width / 914400
    H_in = prs.slide_height / 914400

    nsmap = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    }

    # ── 1. 基本信息 ──
    print(f"=== TEMPLATE ANALYSIS ===")
    print(f"Size: {W_in:.3f} x {H_in:.3f} inches")
    print(f"Slides: {len(prs.slides)}")
    print(f"Layouts: {len(prs.slide_layouts)}")
    for i, layout in enumerate(prs.slide_layouts):
        print(f"  Layout {i}: {layout.name}")
    print()

    # ── 2. 每页 Shape 清单 ──
    for si, slide in enumerate(prs.slides):
        # 简单页面类型推断
        all_text = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                all_text.append(shape.text_frame.text.strip())
        full_text = ' '.join(all_text)
        if si == 0:
            page_type = 'cover'
        elif any(k in full_text for k in ['目录','contents','agenda','概览','overview']):
            page_type = 'toc'
        elif any(k in full_text for k in ['谢谢','感谢','thank','contact','谢观']):
            page_type = 'back_cover'
        elif len(all_text) <= 3 and any(k in full_text for k in ['PART','章节','第','section']):
            page_type = 'transition'
        else:
            page_type = 'content'

        print(f"--- Slide {si} [{page_type}] ({len(slide.shapes)} shapes) ---")

        # 检测背景图
        bg = slide.background
        bg_fill = bg.fill
        if bg_fill.type is not None:
            print(f"  [background: fill_type={bg_fill.type}]")

        for shape in slide.shapes:
            pos = f"({shape.left/914400:.2f}, {shape.top/914400:.2f})"
            size = f"({shape.width/914400:.2f}x{shape.height/914400:.2f})"

            # shape 类型
            stype = str(shape.shape_type)
            tag = shape._element.tag.split('}')[-1]
            if tag == 'cxnSp':
                stype = 'connector'
            elif tag == 'grpSp':
                stype = 'group'
            elif shape.shape_type == 13:
                stype = 'image'

            info = f"  {stype} {pos} {size}"

            # 文字内容
            if shape.has_text_frame:
                txt = shape.text_frame.text[:80].replace('\n', ' | ')
                info += f' text="{txt}"'
                if shape.text_frame.paragraphs:
                    p0 = shape.text_frame.paragraphs[0]
                    if p0.font.size:
                        info += f' size={p0.font.size.pt}pt'
                    try:
                        if p0.font.color and p0.font.color.type is not None:
                            info += f' color=#{p0.font.color.rgb}'
                    except:
                        pass
                    if p0.font.name:
                        info += f' font={p0.font.name}'

            # 几何类型
            sp = shape._element
            prst = sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom')
            if prst is not None:
                info += f' geom={prst.get("prst", "?")}'
                avLst = prst.find('{http://schemas.openxmlformats.org/drawingml/2006/main}avLst')
                if avLst is not None:
                    adjs = [gd.get('fmla','') for gd in avLst.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}gd')]
                    if adjs:
                        info += f' adj={adjs}'

            # alpha（per-shape）
            for clr in sp.findall('.//a:solidFill/a:srgbClr', nsmap):
                alpha = clr.find('a:alpha', nsmap)
                if alpha is not None:
                    alpha_val = int(alpha.get('val', '100000'))
                    info += f' alpha={alpha_val/1000:.0f}%'

            # 渐变（per-shape）
            for grad in sp.findall('.//a:gradFill', nsmap):
                stops = grad.findall('.//a:gs', nsmap)
                if stops:
                    grad_info = []
                    for gs in stops:
                        pos = gs.get('pos', '?')
                        srgb = gs.find('a:srgbClr', nsmap)
                        if srgb is not None:
                            grad_info.append(f"pos={pos} #{srgb.get('val','')}")
                    if grad_info:
                        info += f' gradient=[{" | ".join(grad_info)}]'

            print(info)
        print()

    # ── 3. 色值统计 ──
    color_counter = Counter()
    for slide in prs.slides:
        for shape in slide.shapes:
            sp = shape._element
            for clr in sp.findall('.//a:solidFill/a:srgbClr', nsmap):
                val = clr.get('val', '')
                if val:
                    color_counter[val] += 1

    print()
    print("=== COLOR FREQUENCY ===")
    for hex_val, count in color_counter.most_common(20):
        print(f"  #{hex_val}: {count} occurrences")

    # ── 4. Theme Colors ──
    print()
    print("=== THEME COLORS ===")
    theme = prs.slide_masters[0].element.findall('.//a:clrScheme/a:*', nsmap)
    for t in theme:
        tag = t.tag.split('}')[1] if '}' in t.tag else t.tag
        srgb = t.find('a:srgbClr', nsmap)
        if srgb is not None:
            print(f"  {tag}: #{srgb.get('val')}")
        else:
            sysclr = t.find('a:sysClr', nsmap)
            if sysclr is not None:
                print(f"  {tag}: sys={sysclr.get('val')} lastClr={sysclr.get('lastClr')}")

    # ── 5. 字体统计 ──
    font_counter = Counter()
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    if p.font.name:
                        font_counter[p.font.name] += 1
                    for r in p.runs:
                        if r.font.name:
                            font_counter[r.font.name] += 1

    print()
    print("=== FONT FREQUENCY ===")
    for font_name, count in font_counter.most_common(10):
        print(f"  {font_name}: {count} occurrences")

if __name__ == '__main__':
    path = sys.argv[1] if len(sys.argv) > 1 else r'template.pptx'
    analyze(path)
```

**工作量：1 天**（合并 + 格式优化 + 测试）

### 3.3 输出示例

```
=== TEMPLATE ANALYSIS ===
Size: 13.333 x 7.500 inches
Slides: 5
Layouts: 7
  Layout 0: Title Slide
  Layout 6: Blank

--- Slide 0 [cover] (8 shapes) ---
  group (0.25, 0.53) (2.28x0.34)
  text (0.96, 1.23) (2.56x2.04) text="乡" size=174pt font=汉仪许静行楷简 gradient=[pos=18000 #F2F45D | pos=56000 #11810F | pos=81000 #0D4609]
  rounded_rect (1.04, 5.13) (3.82x0.55) geom=roundRect adj=['val 16667']
  connector (5.89, 1.26) (0.00x5.66)
  rounded_rect (6.24, 1.26) (6.25x0.73) geom=plaque alpha=20%

--- Slide 1 [content] (12 shapes) ---
  text (0.25, 2.5) (0.5x3.0) text="PART ONE" size=14pt color=#7EAB77 font=汉仪旗黑-55简
  text (0.65, 0.45) (11.38x0.5) text="标题" size=28pt color=#0D4609 font=汉仪旗黑-55简
  ...

=== COLOR FREQUENCY ===
  #2E6504: 23 occurrences
  #7DA92F: 18 occurrences
  #466740: 12 occurrences
  #D4E3AC: 9 occurrences
  #84AF7D: 7 occurrences

=== THEME COLORS ===
  dk1: #0D4609
  lt1: #FFFFFF
  accent1: #2E6504
  accent2: #7DA92F

=== FONT FREQUENCY ===
  汉仪旗黑-55简: 45 occurrences
  汉仪许静行楷简: 8 occurrences
```

LLM 读到这个输出，自然就能理解：
- 绿色系模板，primary=#2E6504，accent=#7DA92F
- Slide 0 是封面，Slide 1 是正文页模板（copy_decorations 的来源）
- 封面有渐变大字、高圆角胶囊条、连接线、半透明 plaque
- 正文页有竖排装饰文字"PART ONE"、标题区、分隔线
- 正文用旗黑 55简，14-16pt
- 有 alpha 透明度（plaque 是 20% 透明）

---

## 四、Build 辅助库

### 4.1 现状

`build_data_analysis.py` 中手写了 `_rect/_rrect/_oval/_text/_multiline/_kpi_card/_bar` 等辅助函数。这些函数是 LLM 写 build 脚本时的"积木"。

### 4.2 需要做的

**把这些辅助函数提取为一个可导入的库，让 LLM 写脚本时 `from ppt_build_helpers import *`。**

```python
# src/ppt_pro_max/build_helpers.py
"""
Build 辅助库 — LLM 写 build 脚本时的工具箱

用法（在 build 脚本中）：
    from ppt_pro_max.build_helpers import *
    C = {'primary': '#2E6504', 'accent': '#7DA92F', ...}
    prs = Presentation(template_path)  # 保留模板原有页面
    s = add_slide(prs)
    page_header(s, '标题', '副标题', C)
    kpi_card(s, 1.0, 1.5, 3.0, 1.35, '12.8亿', '年度产值', '+8.3%', C=C)
    bar_chart(s, 1.0, 3.5, data, C=C)
    prs.save('output.pptx')
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import copy, os


def _resolve_color(val, C):
    """解析颜色：角色名→hex，hex→hex，缺失→fallback"""
    if val is None:
        return '#000000'
    if val.startswith('#'):
        return val
    return (C or {}).get(val, '#000000')


def _rgb(hex_str):
    return RGBColor.from_string(hex_str.lstrip('#'))


def add_slide(prs, layout_index=None):
    """添加空白 slide"""
    if layout_index is not None:
        return prs.slides.add_slide(prs.slide_layouts[layout_index])
    for layout in prs.slide_layouts:
        if 'blank' in layout.name.lower():
            return prs.slides.add_slide(layout)
    return prs.slides.add_slide(prs.slide_layouts[-1])


def rect(slide, left, top, width, height, fill, line=None, C=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(left), Inches(top),
                                    Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(_resolve_color(fill, C))
    if line:
        shape.line.color.rgb = _rgb(_resolve_color(line, C))
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


def rrect(slide, left, top, width, height, fill, line=None, C=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(left), Inches(top),
                                    Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(_resolve_color(fill, C))
    if line:
        shape.line.color.rgb = _rgb(_resolve_color(line, C))
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


def oval(slide, left, top, width, height, fill, line=None, C=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                    Inches(left), Inches(top),
                                    Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(_resolve_color(fill, C))
    if line:
        shape.line.color.rgb = _rgb(_resolve_color(line, C))
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


def text(slide, left, top, width, height, txt, font_size=12,
         color='text_body', bold=False, align='left',
         font_name=None, C=None):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top),
                                      Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = txt
    p.font.size = Pt(font_size)
    color_val = _resolve_color(color, C)
    p.font.color.rgb = _rgb(color_val)
    p.font.bold = bold
    if font_name:
        p.font.name = font_name
    p.alignment = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER,
                   'right': PP_ALIGN.RIGHT}[align]
    return txBox


def multiline(slide, left, top, width, height, lines, font_size=12,
              color='text_body', bold=False, align='left',
              font_name=None, C=None):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top),
                                      Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    color_val = _resolve_color(color, C)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size)
        p.font.color.rgb = _rgb(color_val)
        p.font.bold = bold
        if font_name:
            p.font.name = font_name
        p.alignment = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER,
                       'right': PP_ALIGN.RIGHT}[align]
        p.space_before = Pt(2)
        p.space_after = Pt(2)
    return txBox


def copy_decorations(slide, template_slide, skip_long_text=True, skip_image=True):
    """从模板页复制装饰元素
    skip_long_text: 跳过长文本（>50字符），保留短装饰文字如"PART ONE"
    skip_image: 跳过图片
    """
    for shape in template_slide.shapes:
        if skip_image and shape.shape_type == 13:
            continue
        if skip_long_text and shape.has_text_frame:
            if len(shape.text_frame.text) > 50:
                continue
        el = copy.deepcopy(shape._element)
        slide.shapes._spTree.append(el)


def copy_logo(slide, template_slide, color_hints=None):
    """从模板复制 logo（GroupShape，按颜色匹配或取第一个 group）"""
    for shape in template_slide.shapes:
        if shape.shape_type != 6:
            continue
        if color_hints:
            sp = shape._element
            ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
            for hint in color_hints:
                if sp.find(f'.//{{{ns}}}srgbClr[@val="{hint.lstrip("#")}"]') is not None:
                    el = copy.deepcopy(sp)
                    slide.shapes._spTree.append(el)
                    return
        else:
            el = copy.deepcopy(shape._element)
            slide.shapes._spTree.append(el)
            return


def top_bar(slide, color, width=13.333, height=0.08, C=None):
    """顶部装饰条"""
    color_val = _resolve_color(color, C)
    return rect(slide, 0, 0, width, height, color_val)


def page_header(slide, title, subtitle='', C=None, left=0.65, width=None):
    """页面标题区：标题 + 副标题 + 分隔线"""
    C = C or {}
    cw = width or (13.333 - 2 * left)
    text(slide, left, 0.45, cw, 0.5, title,
         font_size=28, color=C.get('text_dark', '#000000'), bold=True,
         font_name=C.get('font_heading'), C=C)
    if subtitle:
        text(slide, left, 0.95, cw, 0.25, subtitle,
             font_size=11, color=C.get('text_muted', '#666666'),
             font_name=C.get('font_body'), C=C)
    rect(slide, left, 1.25, cw, 0.004, C.get('divider', '#CCCCCC'))


def kpi_card(slide, left, top, width, height, number, label,
             trend='', trend_up=True, C=None):
    """KPI 卡片"""
    C = C or {}
    rrect(slide, left, top, width, height, C.get('white', '#FFFFFF'),
          line=C.get('light', '#DDDDDD'), C=C)
    rect(slide, left, top, width, 0.06, C.get('accent', '#4CAF50'), C=C)
    text(slide, left + 0.2, top + 0.25, width - 0.4, 0.5, number,
         font_size=28, color=C.get('primary', '#1B5E20'), bold=True,
         font_name=C.get('font_heading'), C=C)
    text(slide, left + 0.2, top + 0.75, width - 0.4, 0.3, label,
         font_size=11, color=C.get('text_muted', '#666666'),
         font_name=C.get('font_body'), C=C)
    if trend:
        tc = C.get('primary', '#1B5E20') if trend_up else '#C53030'
        text(slide, left + 0.2, top + 1.05, width - 0.4, 0.25, trend,
             font_size=10, color=tc, bold=True,
             font_name=C.get('font_body'), C=C)


def bar_chart(slide, left, top, data, max_width=5.0, bar_height=0.3, C=None):
    """条形图
    data: [(label, pct_0_to_1, value_str), ...]
    """
    C = C or {}
    bar_colors = [C.get('primary', '#1B5E20'), C.get('accent', '#4CAF50'),
                  C.get('muted', '#81C784'), C.get('light', '#C8E6C9')]
    for i, (label, pct, val) in enumerate(data):
        y = top + i * (bar_height + 0.2)
        rrect(slide, left, y, max_width, bar_height, C.get('bg_tint', '#F5F5F5'), C=C)
        bar_w = max_width * pct
        if bar_w > 0:
            rrect(slide, left, y, bar_w, bar_height,
                  bar_colors[i % len(bar_colors)], C=C)
        text(slide, left - 0.9, y - 0.03, 0.85, bar_height, label,
             font_size=10, color=C.get('text_body', '#333333'), align='right',
             font_name=C.get('font_body'), C=C)
        text(slide, left + max_width + 0.08, y - 0.03, 0.6, bar_height, val,
             font_size=10, color=C.get('text_dark', '#000000'), bold=True,
             font_name=C.get('font_body'), C=C)


def comparison_bars(slide, left, top, metrics, max_width=4.0, C=None):
    """对比条形图（两年对比）
    metrics: [(label, v_old, v_new, pct_old, pct_new), ...]
    """
    C = C or {}
    for label, v_old, v_new, pct_old, pct_new in metrics:
        text(slide, left - 1.1, top - 0.02, 1.0, 0.2, label,
             font_size=11, color=C.get('text_body', '#333333'), bold=True,
             align='right', font_name=C.get('font_body'), C=C)
        rrect(slide, left, top, max_width, 0.18, C.get('bg_tint', '#F5F5F5'), C=C)
        bar_old = max_width * pct_old
        if bar_old > 0:
            rrect(slide, left, top, bar_old, 0.18, C.get('muted', '#81C784'), C=C)
        rrect(slide, left, top + 0.22, max_width, 0.18, C.get('bg_tint', '#F5F5F5'), C=C)
        bar_new = max_width * pct_new
        if bar_new > 0:
            rrect(slide, left, top + 0.22, bar_new, 0.18, C.get('primary', '#1B5E20'), C=C)
        text(slide, left + max_width + 0.1, top - 0.03, 0.8, 0.2, v_old,
             font_size=9, color=C.get('text_muted', '#666666'),
             font_name=C.get('font_body'), C=C)
        text(slide, left + max_width + 0.1, top + 0.19, 0.8, 0.2, v_new,
             font_size=9, color=C.get('text_dark', '#000000'), bold=True,
             font_name=C.get('font_body'), C=C)
        top += 0.55
    return top


def donut_chart(slide, cx, cy, radius, inner_radius, sectors, C=None):
    """环形图（简化版：用重叠圆模拟，最后一个扇区颜色会覆盖前面）
    sectors: [(name, pct_str, color_hex), ...]
    如需精确扇区，LLM 可手写 XML path 或使用组件库
    """
    C = C or {}
    import math
    for name, pct_str, clr in sectors:
        oval(slide, cx - radius, cy - radius, radius * 2, radius * 2, clr, C=C)
    oval(slide, cx - inner_radius, cy - inner_radius,
         inner_radius * 2, inner_radius * 2,
         C.get('background', '#FFFFFF'), C=C)
    text(slide, cx - 0.5, cy - 0.2, 1.0, 0.4, '100%',
         font_size=20, color=C.get('primary', '#1B5E20'), bold=True,
         align='center', font_name=C.get('font_heading'), C=C)
    # 图例
    ly = cy - radius
    lx = cx + radius + 0.5
    for name, pct_str, clr in sectors:
        rrect(slide, lx, ly, 0.2, 0.2, clr, C=C)
        text(slide, lx + 0.3, ly - 0.02, 1.5, 0.25, f'{name}  {pct_str}',
             font_size=11, color=C.get('text_body', '#333333'),
             font_name=C.get('font_body'), C=C)
        ly += 0.35


def highlight_cards(slide, left, top, cards, total_width=12.0, C=None):
    """高亮卡片组
    cards: [(title, description, accent_color), ...]
    """
    C = C or {}
    n = len(cards)
    gap = 0.35
    card_w = (total_width - gap * (n - 1)) / n
    for i, (title, desc, accent) in enumerate(cards):
        x = left + i * (card_w + gap)
        rrect(slide, x, top, card_w, 1.2, C.get('card_bg', '#F9F9F9'),
              line=C.get('light', '#DDDDDD'), C=C)
        rect(slide, x, top, card_w, 0.06, accent, C=C)
        text(slide, x + 0.2, top + 0.18, card_w - 0.4, 0.3, title,
             font_size=12, color=C.get('text_dark', '#000000'), bold=True,
             font_name=C.get('font_heading'), C=C)
        text(slide, x + 0.2, top + 0.52, card_w - 0.4, 0.5, desc,
             font_size=10, color=C.get('text_muted', '#666666'),
             font_name=C.get('font_body'), C=C)
```

**工作量：1-2 天**（从 build_data_analysis.py 提取 + 参数化 + 测试）

### 4.3 设计原则

1. **每个函数都接受 C={} 参数** — LLM 从分析结果构建色板 dict，传入所有函数
2. **颜色参数支持角色名和 hex** — `color='primary'` 或 `color='#2E6504'` 都行
3. **函数名简短** — LLM 写脚本时容易记忆和调用
4. **copy_decorations() 通用** — 从模板页复制所有非内容 shape，适配任何模板

---

## 五、Skill 流程定义

### 5.1 新增 Skill 步骤

在现有 Skill 的 Build 模式流程中，Step 2 之前插入模板分析：

```
Step 1: 用户需求 + 模板
    ↓
Step 1.5: 模板分析（新增）
    LLM 调用 analyze_template.py template.pptx
    LLM 读取输出，提取 VI 信息
    ↓
Step 2: Visual Proposals (Build Mode)
    LLM 用模板 VI 写 build 脚本
    色板 C={} 从分析结果填入
    装饰用 copy_decorations() 从模板复制
    ↓
Step 3-5: 同现有流程
```

### 5.2 LLM 写 build 脚本的模板

```python
# LLM 生成的 build 脚本模板
from pptx import Presentation
from ppt_pro_max.build_helpers import *

# ── 从模板分析结果填入色板 ──
C = {
    'primary': '#2E6504',
    'primary_mid': '#466740',
    'accent': '#7DA92F',
    'muted': '#84AF7D',
    'light': '#D4E3AC',
    'lighter': '#E7F3A6',
    'lightest': '#F2F8D6',
    'background': '#FFFFFF',
    'bg_tint': '#F2F8D6',
    'white': '#FFFFFF',
    'text_dark': '#0D4609',
    'text_body': '#2E6504',
    'text_muted': '#466740',
    'divider': '#84AF7D',
    'card_bg': '#F6FAE8',
    'highlight': '#7DA92F',
    'font_heading': '汉仪许静行楷简',
    'font_body': '汉仪旗黑-55简',
}

# ── 读取模板，保留原有页面 ──
template_path = 'template.pptx'
prs = Presentation(template_path)

# ── 新增页面 ──
s6 = add_slide(prs)
rect(s6, 0, 0, 13.333, 7.5, C['background'], C=C)
top_bar(s6, C['accent'], C=C)
copy_logo(s6, prs.slides[1], color_hints=['#2E6504', '#7DA92F'])
page_header(s6, '乡村振兴核心数据总览', '2024年度关键指标汇总', C=C)

# KPI cards
kpi_w = (12.0 - 0.35 * 3) / 4
kpi_data = [
    ('12.8亿', '年度农业总产值', '同比 +8.3%', True),
    ('1,247万', '新增就业人数', '同比 +15.2%', True),
    ('3,580', '示范村建设数', '同比 +22.6%', True),
    ('96.5%', '基础设施覆盖率', '同比 +3.1%', True),
]
kx = 0.65
for num, label, trend, up in kpi_data:
    kpi_card(s6, kx, 1.55, kpi_w, 1.35, num, label, trend, up, C=C)
    kx += kpi_w + 0.35

# Bar chart
bar_chart(s6, 1.65, 3.55, [
    ('东部地区', 0.92, '460亿'),
    ('中部地区', 0.75, '375亿'),
    ('西部地区', 0.68, '340亿'),
    ('东北地区', 0.48, '240亿'),
], max_width=5.0, C=C)

# ── 保存 ──
output_path = 'output.pptx'
os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
prs.save(output_path)
print(f'Saved: {output_path}')
print(f'Slides: {len(prs.slides)}')
```

---

## 六、实施计划

| Phase | 内容 | 工作量 |
|-------|------|--------|
| **P1** | 合并 inspect+extract 为 analyze_template.py | 1天 |
| **P2** | 提取 build_helpers.py | 1-2天 |
| **P3** | 更新 Skill 定义（Step 1.5 模板分析） | 0.5天 |
| **P4** | 测试（analyze_template + build_helpers + E2E） | 1天 |
| **总计** | | **3-4天** |

---

## 七、与 v5 方案对比

| 维度 | v5（传统思路） | v6（Skill 思路） |
|------|---------------|-----------------|
| 新增代码 | ~2,420 行 | ~470 行 |
| 修改/删除代码 | ~5,100 行重构 | 0 行（纯加法） |
| 新增模块 | vi_analyzer/ + vi_builder/ + vi_spec.py | build_helpers.py + analyze_template.py |
| 核心逻辑 | 代码自动化（VISpec→VIBuilder→渲染） | LLM 自动化（读文本→分析→写脚本） |
| 颜色角色映射 | 算法（~150行） | LLM 自己判断（0行） |
| shape 分类 | 分类器（~200行） | LLM 自己判断（0行） |
| 框架页复刻 | 元素提取+逐元素渲染（~400行） | 保留原页面+copy_decorations（~20行） |
| 工期 | 23-28天 | 3-4天 |
| 通用性 | 需要为每种模板适配 | 任何模板都行（LLM 自适应） |
| 废弃代码清理 | 需要大量清理 | 不涉及 |

---

## 八、风险

| 风险 | 概率 | 影响 | 缓解 |
|------|------|------|------|
| LLM 分析色值出错 | 低 | 色板不对 | 人工检查 C={} 定义 |
| LLM 写的 build 脚本有 bug | 中 | 运行报错 | 运行→报错→LLM 修复→重跑 |
| copy_decorations 复制了不该复制的 | 低 | 装饰多余 | LLM 可指定 skip 条件 |
| build_helpers 函数不够用 | 中 | LLM 需要手写更多代码 | 按需增量添加函数 |

**核心优势：** 所有风险都是"LLM 输出不完美"，而 LLM 输出可以迭代修复。传统代码的 bug 需要开发者改代码，LLM 的"bug"只需要重新跑一遍。

---

## 附录：审查修正记录

### 已修正的 Bug

| Bug | 问题 | 修正 |
|-----|------|------|
| 颜色解析崩溃 | `C.get(fill, fill)` 在角色名缺失时返回角色名字符串，`_rgb()` 崩溃 | 新增 `_resolve_color(val, C)` 函数，缺失角色返回 `#000000` |
| top_bar 崩溃 | `slide._element.getparent()` 返回 None | 移除 XML 遍历，直接用 `width=13.333` 参数（LLM 可从分析结果填入实际值） |
| donut_chart 不精确 | 重叠圆只显示最后一个扇区颜色 | 标注为简化版，LLM 需要精确扇区时可手写 XML path 或用组件库 |

### 已修正的遗漏

| 遗漏 | 问题 | 修正 |
|------|------|------|
| 无页面类型分类 | LLM 不知道哪个 slide 是正文页模板 | analyze_template.py 新增 `[cover/toc/content/transition/back_cover]` 标签 |
| alpha/渐变与 shape 脱节 | alpha 和渐变信息无法关联到具体 shape | 移入 per-shape 循环，作为 shape info 的一部分输出 |
| copy_decorations 跳过装饰文字 | `skip_text=True` 跳过了 "PART ONE" 等装饰文字 | 改为 `skip_long_text=True`，只跳过 >50 字符的长文本 |

### 已知局限（不修正，LLM 可自行处理）

1. **donut_chart 简化版**：用重叠圆模拟，非精确扇区。LLM 需要精确图表时可手写或用组件库。
2. **analyze_template.py 不提取图片文件**：只报告 shape 位置和类型，不提取 image blob。LLM 需要图片时可手写路径或用 ImageFetcher。
3. **build_helpers 与 PrecisionRenderer 功能重叠**：build_helpers 更简单、更适合 LLM 直接调用；PrecisionRenderer 更完整（CJK、渐变、阴影）。两者并存，LLM 按需选择。
4. **design_dna_extractor.py 删除后 PagePlan 等数据类需迁移**：v1.0.0 删除时同步迁移到 `vi_spec.py` 或 `content_parser.py`。
5. **extract_design_dna() 公共 API**：v1.0.0 删除时需在 `__init__.py` 中移除或替换为 `analyze_template()` 的 Python API 包装。
