"""build_helpers — LLM build script toolbox.

Usage in build scripts:
    from ppt_pro_max.build_helpers import *
    C = {'primary': '#2E6504', 'accent': '#7DA92F', ...}
    prs = Presentation(template_path)
    s = add_slide(prs)
    page_header(s, 'Title', 'Subtitle', C)
    kpi_card(s, 1.0, 1.5, 3.0, 1.35, '12.8亿', '年度产值', '+8.3%', C=C)
    prs.save('output.pptx')
"""
from __future__ import annotations

import copy

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE


def _resolve_color(val, C):
    if val is None:
        return '#000000'
    if val.startswith('#'):
        return val
    return (C or {}).get(val, '#000000')


def _rgb(hex_str):
    return RGBColor.from_string(hex_str.lstrip('#'))


def add_slide(prs, layout_index=None):
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
    p.font.color.rgb = _rgb(_resolve_color(color, C))
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
    for shape in template_slide.shapes:
        if skip_image and shape.shape_type == 13:
            continue
        if skip_long_text and shape.has_text_frame:
            if len(shape.text_frame.text) > 50:
                continue
        el = copy.deepcopy(shape._element)
        slide.shapes._spTree.append(el)


def copy_logo(slide, template_slide, color_hints=None):
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
    color_val = _resolve_color(color, C)
    return rect(slide, 0, 0, width, height, color_val)


def page_header(slide, title, subtitle='', C=None, left=0.65, width=None):
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
    C = C or {}
    for name, pct_str, clr in sectors:
        oval(slide, cx - radius, cy - radius, radius * 2, radius * 2, clr, C=C)
    oval(slide, cx - inner_radius, cy - inner_radius,
         inner_radius * 2, inner_radius * 2,
         C.get('background', '#FFFFFF'), C=C)
    text(slide, cx - 0.5, cy - 0.2, 1.0, 0.4, '100%',
         font_size=20, color=C.get('primary', '#1B5E20'), bold=True,
         align='center', font_name=C.get('font_heading'), C=C)
    ly = cy - radius
    lx = cx + radius + 0.5
    for name, pct_str, clr in sectors:
        rrect(slide, lx, ly, 0.2, 0.2, clr, C=C)
        text(slide, lx + 0.3, ly - 0.02, 1.5, 0.25, f'{name}  {pct_str}',
             font_size=11, color=C.get('text_body', '#333333'),
             font_name=C.get('font_body'), C=C)
        ly += 0.35


def highlight_cards(slide, left, top, cards, total_width=12.0, C=None):
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
