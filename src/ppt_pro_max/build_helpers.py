"""build_helpers — LLM build script toolbox.

Usage in build scripts:
    from ppt_pro_max.build_helpers import *
    C = {'primary': '#2E6504', 'accent': '#7DA92F', ...}
    t = TYPOGRAPHY['mckinsey']
    sp = SPACING['mckinsey']
    prs = Presentation(template_path)
    s = add_slide(prs)
    page_header(s, 'Title', 'Subtitle', C, typo=t, spacing=sp)
    kpi_card(s, 1.0, 1.5, 3.0, 1.35, '12.8亿', '年度产值', '+8.3%', C=C, typo=t)
    prs.save('output.pptx')
"""
from __future__ import annotations

import copy

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


def _resolve_color(val, C):
    if val is None:
        return '#000000'
    if val.startswith('#'):
        return val
    return (C or {}).get(val, '#000000')


def _rgb(hex_str):
    return RGBColor.from_string(hex_str.lstrip('#'))


class Typography:
    """Font size scale for a design style.

    Each style defines its own scale so that headings, body, and captions
    maintain consistent visual rhythm across all slides.
    """

    def __init__(self, hero=44, h1=28, h2=20, h3=16, body=12, caption=10, micro=8):
        self.hero = hero
        self.h1 = h1
        self.h2 = h2
        self.h3 = h3
        self.body = body
        self.caption = caption
        self.micro = micro

    def scale(self, level):
        return getattr(self, level, self.body)


class Spacing:
    """Spacing system for a design style.

    Controls margins, padding, gaps, and rhythm so that every component
    has consistent breathing room.
    """

    def __init__(self, page_margin=0.65, section_gap=0.5, card_gap=0.35,
                 card_padding=0.2, line_height=1.4, bar_gap=0.2):
        self.page_margin = page_margin
        self.section_gap = section_gap
        self.card_gap = card_gap
        self.card_padding = card_padding
        self.line_height = line_height
        self.bar_gap = bar_gap


TYPOGRAPHY = {
    'mckinsey': Typography(hero=44, h1=28, h2=20, h3=16, body=12, caption=10, micro=8),
    'cyberpunk': Typography(hero=48, h1=28, h2=18, h3=14, body=11, caption=9, micro=7),
    'creative': Typography(hero=44, h1=28, h2=22, h3=18, body=13, caption=11, micro=9),
    'professional': Typography(hero=44, h1=28, h2=20, h3=16, body=12, caption=10, micro=8),
    'minimal': Typography(hero=40, h1=24, h2=18, h3=14, body=11, caption=9, micro=7),
}

SPACING = {
    'mckinsey': Spacing(page_margin=0.65, section_gap=0.5, card_gap=0.35,
                        card_padding=0.2, line_height=1.4, bar_gap=0.2),
    'cyberpunk': Spacing(page_margin=0.8, section_gap=0.6, card_gap=0.4,
                         card_padding=0.25, line_height=1.3, bar_gap=0.25),
    'creative': Spacing(page_margin=0.8, section_gap=0.6, card_gap=0.4,
                        card_padding=0.25, line_height=1.5, bar_gap=0.25),
    'professional': Spacing(page_margin=0.65, section_gap=0.5, card_gap=0.35,
                            card_padding=0.2, line_height=1.4, bar_gap=0.2),
    'minimal': Spacing(page_margin=1.0, section_gap=0.6, card_gap=0.5,
                       card_padding=0.3, line_height=1.5, bar_gap=0.3),
}


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
         font_name=None, C=None, anchor='top'):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top),
                                      Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    if anchor == 'middle':
        try:
            txBox.text_frame._txBody.bodyPr.set('anchor', 'ctr')
        except Exception:
            pass
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
              font_name=None, C=None, line_spacing=None):
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
        if line_spacing:
            p.space_before = Pt(line_spacing)
            p.space_after = Pt(line_spacing)
        else:
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


def page_header(slide, title, subtitle='', C=None, left=0.65, width=None,
                typo=None, spacing=None):
    C = C or {}
    cw = width or (13.333 - 2 * left)
    t = typo or TYPOGRAPHY.get('mckinsey')
    sp = spacing or SPACING.get('mckinsey')
    text(slide, left, 0.45, cw, 0.5, title,
         font_size=t.h1, color=C.get('text_dark', '#000000'), bold=True,
         font_name=C.get('font_heading'), C=C)
    if subtitle:
        text(slide, left, 0.95, cw, 0.25, subtitle,
             font_size=t.caption, color=C.get('text_muted', '#666666'),
             font_name=C.get('font_body'), C=C)
    rect(slide, left, 1.25, cw, 0.004, C.get('divider', '#CCCCCC'))


def kpi_card(slide, left, top, width, height, number, label,
             trend='', trend_up=True, C=None, typo=None, grouped=True):
    C = C or {}
    t = typo or TYPOGRAPHY.get('mckinsey')
    pad = 0.2

    if grouped:
        group = slide.shapes.add_group_shape()
        gs = group.shapes

        bg = gs.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                          Inches(left), Inches(top),
                          Inches(width), Inches(height))
        bg.fill.solid()
        bg.fill.fore_color.rgb = _rgb(C.get('white', '#FFFFFF'))
        bg.line.color.rgb = _rgb(C.get('light', '#DDDDDD'))
        bg.line.width = Pt(1)

        accent_bar = gs.add_shape(MSO_SHAPE.RECTANGLE,
                                  Inches(left), Inches(top),
                                  Inches(width), Inches(0.06))
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = _rgb(C.get('accent', '#4CAF50'))
        accent_bar.line.fill.background()

        num_box = gs.add_textbox(Inches(left + pad), Inches(top + 0.25),
                                 Inches(width - 2 * pad), Inches(0.5))
        p = num_box.text_frame.paragraphs[0]
        p.text = number
        p.font.size = Pt(t.h1)
        p.font.color.rgb = _rgb(C.get('primary', '#1B5E20'))
        p.font.bold = True
        if C.get('font_heading'):
            p.font.name = C['font_heading']

        lbl_box = gs.add_textbox(Inches(left + pad), Inches(top + 0.75),
                                 Inches(width - 2 * pad), Inches(0.3))
        p2 = lbl_box.text_frame.paragraphs[0]
        p2.text = label
        p2.font.size = Pt(t.caption)
        p2.font.color.rgb = _rgb(C.get('text_muted', '#666666'))
        if C.get('font_body'):
            p2.font.name = C['font_body']

        if trend:
            tc = C.get('primary', '#1B5E20') if trend_up else '#C53030'
            trend_box = gs.add_textbox(Inches(left + pad), Inches(top + 1.05),
                                       Inches(width - 2 * pad), Inches(0.25))
            p3 = trend_box.text_frame.paragraphs[0]
            p3.text = trend
            p3.font.size = Pt(t.micro)
            p3.font.color.rgb = _rgb(tc)
            p3.font.bold = True
            if C.get('font_body'):
                p3.font.name = C['font_body']

        return group
    else:
        rrect(slide, left, top, width, height, C.get('white', '#FFFFFF'),
              line=C.get('light', '#DDDDDD'), C=C)
        rect(slide, left, top, width, 0.06, C.get('accent', '#4CAF50'), C=C)
        text(slide, left + pad, top + 0.25, width - 2 * pad, 0.5, number,
             font_size=t.h1, color=C.get('primary', '#1B5E20'), bold=True,
             font_name=C.get('font_heading'), C=C)
        text(slide, left + pad, top + 0.75, width - 2 * pad, 0.3, label,
             font_size=t.caption, color=C.get('text_muted', '#666666'),
             font_name=C.get('font_body'), C=C)
        if trend:
            tc = C.get('primary', '#1B5E20') if trend_up else '#C53030'
            text(slide, left + pad, top + 1.05, width - 2 * pad, 0.25, trend,
                 font_size=t.micro, color=tc, bold=True,
                 font_name=C.get('font_body'), C=C)


def bar_chart(slide, left, top, data, max_width=5.0, bar_height=0.3, C=None,
              typo=None, spacing=None, grouped=True):
    C = C or {}
    t = typo or TYPOGRAPHY.get('mckinsey')
    sp = spacing or SPACING.get('mckinsey')
    bar_colors = [C.get('primary', '#1B5E20'), C.get('accent', '#4CAF50'),
                  C.get('muted', '#81C784'), C.get('light', '#C8E6C9')]
    gap = sp.bar_gap

    if grouped:
        group = slide.shapes.add_group_shape()
        gs = group.shapes

        for i, (label, pct, val) in enumerate(data):
            y = top + i * (bar_height + gap)

            bg = gs.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(left), Inches(y),
                              Inches(max_width), Inches(bar_height))
            bg.fill.solid()
            bg.fill.fore_color.rgb = _rgb(C.get('bg_tint', '#F5F5F5'))
            bg.line.fill.background()

            bar_w = max_width * pct
            if bar_w > 0:
                bar = gs.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(left), Inches(y),
                                   Inches(bar_w), Inches(bar_height))
                bar.fill.solid()
                bar.fill.fore_color.rgb = _rgb(bar_colors[i % len(bar_colors)])
                bar.line.fill.background()

            lbl_box = gs.add_textbox(Inches(left - 0.9), Inches(y - 0.03),
                                     Inches(0.85), Inches(bar_height))
            p = lbl_box.text_frame.paragraphs[0]
            p.text = label
            p.font.size = Pt(t.caption)
            p.font.color.rgb = _rgb(C.get('text_body', '#333333'))
            p.alignment = PP_ALIGN.RIGHT
            if C.get('font_body'):
                p.font.name = C['font_body']

            val_box = gs.add_textbox(Inches(left + max_width + 0.08), Inches(y - 0.03),
                                     Inches(0.6), Inches(bar_height))
            p2 = val_box.text_frame.paragraphs[0]
            p2.text = val
            p2.font.size = Pt(t.caption)
            p2.font.color.rgb = _rgb(C.get('text_dark', '#000000'))
            p2.font.bold = True
            if C.get('font_body'):
                p2.font.name = C['font_body']

        return group
    else:
        for i, (label, pct, val) in enumerate(data):
            y = top + i * (bar_height + gap)
            rrect(slide, left, y, max_width, bar_height, C.get('bg_tint', '#F5F5F5'), C=C)
            bar_w = max_width * pct
            if bar_w > 0:
                rrect(slide, left, y, bar_w, bar_height,
                      bar_colors[i % len(bar_colors)], C=C)
            text(slide, left - 0.9, y - 0.03, 0.85, bar_height, label,
                 font_size=t.caption, color=C.get('text_body', '#333333'), align='right',
                 font_name=C.get('font_body'), C=C)
            text(slide, left + max_width + 0.08, y - 0.03, 0.6, bar_height, val,
                 font_size=t.caption, color=C.get('text_dark', '#000000'), bold=True,
                 font_name=C.get('font_body'), C=C)


def comparison_bars(slide, left, top, metrics, max_width=4.0, C=None,
                    typo=None, spacing=None, grouped=True):
    C = C or {}
    t = typo or TYPOGRAPHY.get('mckinsey')
    sp = spacing or SPACING.get('mckinsey')
    row_h = 0.55

    if grouped:
        group = slide.shapes.add_group_shape()
        gs = group.shapes

        for idx, (label, v_old, v_new, pct_old, pct_new) in enumerate(metrics):
            y = top + idx * row_h

            lbl_box = gs.add_textbox(Inches(left - 1.1), Inches(y - 0.02),
                                     Inches(1.0), Inches(0.2))
            p = lbl_box.text_frame.paragraphs[0]
            p.text = label
            p.font.size = Pt(t.caption)
            p.font.color.rgb = _rgb(C.get('text_body', '#333333'))
            p.font.bold = True
            p.alignment = PP_ALIGN.RIGHT
            if C.get('font_body'):
                p.font.name = C['font_body']

            bg1 = gs.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(left), Inches(y),
                               Inches(max_width), Inches(0.18))
            bg1.fill.solid()
            bg1.fill.fore_color.rgb = _rgb(C.get('bg_tint', '#F5F5F5'))
            bg1.line.fill.background()

            bar_old = max_width * pct_old
            if bar_old > 0:
                b1 = gs.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(left), Inches(y),
                                  Inches(bar_old), Inches(0.18))
                b1.fill.solid()
                b1.fill.fore_color.rgb = _rgb(C.get('muted', '#81C784'))
                b1.line.fill.background()

            bg2 = gs.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(left), Inches(y + 0.22),
                               Inches(max_width), Inches(0.18))
            bg2.fill.solid()
            bg2.fill.fore_color.rgb = _rgb(C.get('bg_tint', '#F5F5F5'))
            bg2.line.fill.background()

            bar_new = max_width * pct_new
            if bar_new > 0:
                b2 = gs.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(left), Inches(y + 0.22),
                                  Inches(bar_new), Inches(0.18))
                b2.fill.solid()
                b2.fill.fore_color.rgb = _rgb(C.get('primary', '#1B5E20'))
                b2.line.fill.background()

            old_box = gs.add_textbox(Inches(left + max_width + 0.1), Inches(y - 0.03),
                                     Inches(0.8), Inches(0.2))
            p2 = old_box.text_frame.paragraphs[0]
            p2.text = v_old
            p2.font.size = Pt(t.micro)
            p2.font.color.rgb = _rgb(C.get('text_muted', '#666666'))
            if C.get('font_body'):
                p2.font.name = C['font_body']

            new_box = gs.add_textbox(Inches(left + max_width + 0.1), Inches(y + 0.19),
                                     Inches(0.8), Inches(0.2))
            p3 = new_box.text_frame.paragraphs[0]
            p3.text = v_new
            p3.font.size = Pt(t.micro)
            p3.font.color.rgb = _rgb(C.get('text_dark', '#000000'))
            p3.font.bold = True
            if C.get('font_body'):
                p3.font.name = C['font_body']

        return group
    else:
        for label, v_old, v_new, pct_old, pct_new in metrics:
            text(slide, left - 1.1, top - 0.02, 1.0, 0.2, label,
                 font_size=t.caption, color=C.get('text_body', '#333333'), bold=True,
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
                 font_size=t.micro, color=C.get('text_muted', '#666666'),
                 font_name=C.get('font_body'), C=C)
            text(slide, left + max_width + 0.1, top + 0.19, 0.8, 0.2, v_new,
                 font_size=t.micro, color=C.get('text_dark', '#000000'), bold=True,
                 font_name=C.get('font_body'), C=C)
            top += row_h
        return top


def donut_chart(slide, cx, cy, radius, inner_radius, sectors, C=None,
                typo=None, grouped=True):
    C = C or {}
    t = typo or TYPOGRAPHY.get('mckinsey')

    if grouped:
        group = slide.shapes.add_group_shape()
        gs = group.shapes

        for name, pct_str, clr in sectors:
            outer = gs.add_shape(MSO_SHAPE.OVAL,
                                 Inches(cx - radius), Inches(cy - radius),
                                 Inches(radius * 2), Inches(radius * 2))
            outer.fill.solid()
            outer.fill.fore_color.rgb = _rgb(clr)
            outer.line.fill.background()

        inner = gs.add_shape(MSO_SHAPE.OVAL,
                             Inches(cx - inner_radius), Inches(cy - inner_radius),
                             Inches(inner_radius * 2), Inches(inner_radius * 2))
        inner.fill.solid()
        inner.fill.fore_color.rgb = _rgb(C.get('background', '#FFFFFF'))
        inner.line.fill.background()

        center_box = gs.add_textbox(
            Inches(cx - 0.5), Inches(cy - 0.2),
            Inches(1.0), Inches(0.4))
        p = center_box.text_frame.paragraphs[0]
        p.text = '100%'
        p.font.size = Pt(t.h2)
        p.font.color.rgb = _rgb(C.get('primary', '#1B5E20'))
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        if C.get('font_heading'):
            p.font.name = C['font_heading']

        ly = cy - radius
        lx = cx + radius + 0.5
        for name, pct_str, clr in sectors:
            dot = gs.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(lx), Inches(ly),
                               Inches(0.2), Inches(0.2))
            dot.fill.solid()
            dot.fill.fore_color.rgb = _rgb(clr)
            dot.line.fill.background()

            lbl = gs.add_textbox(Inches(lx + 0.3), Inches(ly - 0.02),
                                 Inches(1.5), Inches(0.25))
            p2 = lbl.text_frame.paragraphs[0]
            p2.text = f'{name}  {pct_str}'
            p2.font.size = Pt(t.caption)
            p2.font.color.rgb = _rgb(C.get('text_body', '#333333'))
            if C.get('font_body'):
                p2.font.name = C['font_body']
            ly += 0.35

        return group
    else:
        for name, pct_str, clr in sectors:
            oval(slide, cx - radius, cy - radius, radius * 2, radius * 2, clr, C=C)
        oval(slide, cx - inner_radius, cy - inner_radius,
             inner_radius * 2, inner_radius * 2,
             C.get('background', '#FFFFFF'), C=C)
        text(slide, cx - 0.5, cy - 0.2, 1.0, 0.4, '100%',
             font_size=t.h2, color=C.get('primary', '#1B5E20'), bold=True,
             align='center', font_name=C.get('font_heading'), C=C)
        ly = cy - radius
        lx = cx + radius + 0.5
        for name, pct_str, clr in sectors:
            rrect(slide, lx, ly, 0.2, 0.2, clr, C=C)
            text(slide, lx + 0.3, ly - 0.02, 1.5, 0.25, f'{name}  {pct_str}',
                 font_size=t.caption, color=C.get('text_body', '#333333'),
                 font_name=C.get('font_body'), C=C)
            ly += 0.35


def highlight_cards(slide, left, top, cards, total_width=12.0, C=None,
                    typo=None, spacing=None, grouped=True):
    C = C or {}
    t = typo or TYPOGRAPHY.get('mckinsey')
    sp = spacing or SPACING.get('mckinsey')
    n = len(cards)
    gap = sp.card_gap
    card_w = (total_width - gap * (n - 1)) / n
    card_h = 1.4
    pad = sp.card_padding

    if grouped:
        group = slide.shapes.add_group_shape()
        gs = group.shapes

        for i, (title, desc, accent) in enumerate(cards):
            x = left + i * (card_w + gap)

            bg = gs.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(x), Inches(top),
                              Inches(card_w), Inches(card_h))
            bg.fill.solid()
            bg.fill.fore_color.rgb = _rgb(C.get('card_bg', '#F9F9F9'))
            bg.line.color.rgb = _rgb(C.get('light', '#DDDDDD'))
            bg.line.width = Pt(1)

            accent_bar = gs.add_shape(MSO_SHAPE.RECTANGLE,
                                      Inches(x), Inches(top),
                                      Inches(card_w), Inches(0.06))
            accent_bar.fill.solid()
            accent_bar.fill.fore_color.rgb = _rgb(accent)
            accent_bar.line.fill.background()

            title_box = gs.add_textbox(Inches(x + pad), Inches(top + 0.18),
                                       Inches(card_w - 2 * pad), Inches(0.3))
            p = title_box.text_frame.paragraphs[0]
            p.text = title
            p.font.size = Pt(t.h3)
            p.font.color.rgb = _rgb(C.get('text_dark', '#000000'))
            p.font.bold = True
            if C.get('font_heading'):
                p.font.name = C['font_heading']

            desc_box = gs.add_textbox(Inches(x + pad), Inches(top + 0.52),
                                      Inches(card_w - 2 * pad), Inches(0.7))
            p2 = desc_box.text_frame.paragraphs[0]
            p2.text = desc
            p2.font.size = Pt(t.caption)
            p2.font.color.rgb = _rgb(C.get('text_muted', '#666666'))
            if C.get('font_body'):
                p2.font.name = C['font_body']

        return group
    else:
        for i, (title, desc, accent) in enumerate(cards):
            x = left + i * (card_w + gap)
            rrect(slide, x, top, card_w, 1.2, C.get('card_bg', '#F9F9F9'),
                  line=C.get('light', '#DDDDDD'), C=C)
            rect(slide, x, top, card_w, 0.06, accent, C=C)
            text(slide, x + pad, top + 0.18, card_w - 2 * pad, 0.3, title,
                 font_size=t.h3, color=C.get('text_dark', '#000000'), bold=True,
                 font_name=C.get('font_heading'), C=C)
            text(slide, x + pad, top + 0.52, card_w - 2 * pad, 0.5, desc,
                 font_size=t.caption, color=C.get('text_muted', '#666666'),
                 font_name=C.get('font_body'), C=C)


def code_block(slide, left, top, width, height, lines, language='python',
               C=None, typo=None, grouped=True):
    C = C or {}
    t = typo or TYPOGRAPHY.get('mckinsey')

    if grouped:
        group = slide.shapes.add_group_shape()
        gs = group.shapes

        bg = gs.add_shape(MSO_SHAPE.RECTANGLE,
                          Inches(left), Inches(top),
                          Inches(width), Inches(height))
        bg.fill.solid()
        bg.fill.fore_color.rgb = _rgb('#1E1E1E')
        bg.line.fill.background()

        badge_w = len(language) * 0.12 + 0.3
        badge = gs.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                             Inches(left + 0.1), Inches(top - 0.28),
                             Inches(badge_w), Inches(0.25))
        badge.fill.solid()
        badge.fill.fore_color.rgb = _rgb(C.get('accent', '#4CAF50'))
        badge.line.fill.background()

        badge_txt = gs.add_textbox(Inches(left + 0.15), Inches(top - 0.28),
                                   Inches(badge_w - 0.1), Inches(0.25))
        p = badge_txt.text_frame.paragraphs[0]
        p.text = language
        p.font.size = Pt(t.micro)
        p.font.color.rgb = _rgb(C.get('white', '#FFFFFF'))
        p.font.bold = True
        if C.get('font_body'):
            p.font.name = C['font_body']

        code_box = gs.add_textbox(Inches(left + 0.2), Inches(top + 0.15),
                                  Inches(width - 0.4), Inches(height - 0.3))
        tf = code_box.text_frame
        tf.word_wrap = True
        for i, line in enumerate(lines):
            p2 = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p2.text = line
            p2.font.size = Pt(t.body)
            p2.font.color.rgb = _rgb('#D4D4D4')
            p2.font.name = 'Consolas'
            p2.space_before = Pt(3)
            p2.space_after = Pt(3)

        return group
    else:
        rect(slide, left, top, width, height, '#1E1E1E')
        rrect(slide, left, top - 0.28, len(language) * 0.12 + 0.3, 0.25,
              C.get('accent', '#4CAF50'), C=C)
        text(slide, left + 0.05, top - 0.28, len(language) * 0.12 + 0.2, 0.25,
             language, font_size=t.micro, color=C.get('white', '#FFFFFF'),
             bold=True, font_name='Consolas', C=C)
        multiline(slide, left + 0.2, top + 0.15, width - 0.4, height - 0.3,
                  lines, font_size=t.body, color='#D4D4D4',
                  font_name='Consolas', line_spacing=3)


def section_divider(slide, number, title, C=None, typo=None, grouped=True):
    C = C or {}
    t = typo or TYPOGRAPHY.get('mckinsey')

    if grouped:
        group = slide.shapes.add_group_shape()
        gs = group.shapes

        bg = gs.add_shape(MSO_SHAPE.RECTANGLE,
                          Inches(0), Inches(0),
                          Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = _rgb(C.get('primary', '#1B5E20'))
        bg.line.fill.background()

        num_box = gs.add_textbox(Inches(1.2), Inches(1.5),
                                 Inches(3.0), Inches(2.0))
        p = num_box.text_frame.paragraphs[0]
        p.text = str(number).zfill(2)
        p.font.size = Pt(72)
        p.font.color.rgb = _rgb(C.get('light', '#C8E6C9'))
        p.font.bold = True
        if C.get('font_heading'):
            p.font.name = C['font_heading']

        line = gs.add_shape(MSO_SHAPE.RECTANGLE,
                            Inches(1.2), Inches(3.6),
                            Inches(2.0), Inches(0.04))
        line.fill.solid()
        line.fill.fore_color.rgb = _rgb(C.get('accent', '#4CAF50'))
        line.line.fill.background()

        title_box = gs.add_textbox(Inches(1.2), Inches(3.9),
                                   Inches(10.0), Inches(1.5))
        p2 = title_box.text_frame.paragraphs[0]
        p2.text = title
        p2.font.size = Pt(t.hero)
        p2.font.color.rgb = _rgb(C.get('white', '#FFFFFF'))
        p2.font.bold = True
        if C.get('font_heading'):
            p2.font.name = C['font_heading']

        return group
    else:
        rect(slide, 0, 0, 13.333, 7.5, C.get('primary', '#1B5E20'))
        text(slide, 1.2, 1.5, 3.0, 2.0, str(number).zfill(2),
             font_size=72, color=C.get('light', '#C8E6C9'), bold=True,
             font_name=C.get('font_heading'), C=C)
        rect(slide, 1.2, 3.6, 2.0, 0.04, C.get('accent', '#4CAF50'))
        text(slide, 1.2, 3.9, 10.0, 1.5, title,
             font_size=t.hero, color=C.get('white', '#FFFFFF'), bold=True,
             font_name=C.get('font_heading'), C=C)


def hero_slide(slide, title, subtitle='', C=None, typo=None, grouped=True):
    C = C or {}
    t = typo or TYPOGRAPHY.get('mckinsey')

    if grouped:
        group = slide.shapes.add_group_shape()
        gs = group.shapes

        bg = gs.add_shape(MSO_SHAPE.RECTANGLE,
                          Inches(0), Inches(0),
                          Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = _rgb(C.get('primary', '#1B5E20'))
        bg.line.fill.background()

        title_box = gs.add_textbox(Inches(1.2), Inches(2.0),
                                   Inches(10.0), Inches(1.5))
        p = title_box.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(t.hero)
        p.font.color.rgb = _rgb(C.get('white', '#FFFFFF'))
        p.font.bold = True
        if C.get('font_heading'):
            p.font.name = C['font_heading']

        if subtitle:
            sub_box = gs.add_textbox(Inches(1.2), Inches(3.6),
                                     Inches(10.0), Inches(0.5))
            p2 = sub_box.text_frame.paragraphs[0]
            p2.text = subtitle
            p2.font.size = Pt(t.h2)
            p2.font.color.rgb = _rgb(C.get('light', '#C8E6C9'))
            if C.get('font_body'):
                p2.font.name = C['font_body']

        return group
    else:
        rect(slide, 0, 0, 13.333, 7.5, C.get('primary'))
        text(slide, 1.2, 2.0, 10.0, 1.5, title,
             font_size=t.hero, color=C.get('white', '#FFFFFF'), bold=True,
             font_name=C.get('font_heading'), C=C)
        if subtitle:
            text(slide, 1.2, 3.6, 10.0, 0.5, subtitle,
                 font_size=t.h2, color=C.get('light', '#C8E6C9'),
                 font_name=C.get('font_body'), C=C)


def cta_slide(slide, title, subtitle='', C=None, typo=None, grouped=True):
    C = C or {}
    t = typo or TYPOGRAPHY.get('mckinsey')

    if grouped:
        group = slide.shapes.add_group_shape()
        gs = group.shapes

        bg = gs.add_shape(MSO_SHAPE.RECTANGLE,
                          Inches(0), Inches(0),
                          Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = _rgb(C.get('primary', '#1B5E20'))
        bg.line.fill.background()

        title_box = gs.add_textbox(Inches(1.2), Inches(2.5),
                                   Inches(10.0), Inches(1.5))
        p = title_box.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(t.h1 + 12)
        p.font.color.rgb = _rgb(C.get('white', '#FFFFFF'))
        p.font.bold = True
        if C.get('font_heading'):
            p.font.name = C['font_heading']

        if subtitle:
            sub_box = gs.add_textbox(Inches(1.2), Inches(4.0),
                                     Inches(10.0), Inches(0.5))
            p2 = sub_box.text_frame.paragraphs[0]
            p2.text = subtitle
            p2.font.size = Pt(t.h3)
            p2.font.color.rgb = _rgb(C.get('light', '#C8E6C9'))
            if C.get('font_body'):
                p2.font.name = C['font_body']

        return group
    else:
        rect(slide, 0, 0, 13.333, 7.5, C.get('primary'))
        text(slide, 1.2, 2.5, 10.0, 1.5, title,
             font_size=t.h1 + 12, color=C.get('white', '#FFFFFF'), bold=True,
             font_name=C.get('font_heading'), C=C)
        if subtitle:
            text(slide, 1.2, 4.0, 10.0, 0.5, subtitle,
                 font_size=t.h3, color=C.get('light', '#C8E6C9'),
                 font_name=C.get('font_body'), C=C)
