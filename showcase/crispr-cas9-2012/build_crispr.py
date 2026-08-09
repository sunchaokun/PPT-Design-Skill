"""Scientific Research PPT — CRISPR-Cas9 (Jinek et al., Science 2012).

Layout modeled on professional research decks: MULTI-PANEL (a/b/c/d) pages
built from information graphics (sequence color blocks, flow diagrams,
statistics, log-scale data charts, timelines) — NOT large images.

Design rules:
  - Each result page = Figure N + title + 3-4 lettered panels
  - Sequence alignment = per-base color blocks (AUTO_SHAPE)
  - Flow/mechanism = process boxes + arrows
  - Statistics = large prominent numbers
  - Bottom: Note + citation
"""

from ppt_pro_max.build_helpers import *
import os

# ── Semantic research palette ──────────────────────────────────────
C = {
    'up_color': '#C0392B',       # 切割/实验组 (red)
    'down_color': '#2C3E50',     # 对照/阴性 (dark)
    'control_color': '#27AE60',  # 引导/阳性对照 (green)
    'mutant_color': '#8E44AD',   # 突变体 (purple)
    'target_color': '#E74C3C',   # 靶标链 (red)
    'nontarget_color': '#3498DB',# 非靶标链 (blue)
    'guide_color': '#2ECC71',    # sgRNA 引导链 (green)
    'pam_color': '#F39C12',      # PAM (amber)
    'cas9_color': '#9B59B6',     # Cas9 蛋白 (purple)
    'grid_color': '#BDC3C7',
    'text_dark': '#2C3E50',
    'text_body': '#34495E',
    'text_muted': '#7F8C8D',
    'background': '#FFFFFF',
    'bg_tint': '#F4F6F8',
    'divider': '#D5DBDB',
    'white': '#FFFFFF',
    'font_heading': 'Arial',
    'font_body': 'Arial',
    'font_cjk': '微软雅黑',
    'font_mono': 'Consolas',
}

t = TYPOGRAPHY['cjk_professional']
sp = SPACING['professional']

PRS = Presentation()
set_widescreen(PRS)

# ── Layout helpers ─────────────────────────────────────────────────
def figure_header(slide, num, title):
    """Top: Figure number + title."""
    text(slide, 0.7, 0.18, 1.6, 0.3, f'Figure {num}', font_size=12,
         color='text_dark', bold=True, C=C)
    text(slide, 2.0, 0.18, 10.8, 0.35, title, font_size=17,
         color='text_dark', bold=True, C=C)
    rect(s, 0.7, 0.55, 11.9, 0.015, C['divider'], C=C)

def panel_label(slide, x, y, label, title):
    """Panel label: 'a' + title (lowercase letter + subtitle)."""
    text(slide, x, y, 0.3, 0.3, label, font_size=11, bold=True,
         color='text_dark', C=C)
    tw = min(8.0, 12.7 - x - 0.35)
    text(slide, x + 0.35, y, tw, 0.3, title, font_size=12, bold=True,
         color='text_dark', C=C)

def panel_note(slide, x, y, note, cite=''):
    """Bottom note + citation."""
    text(slide, x, y, 0.55, 0.25, 'Note', font_size=8, bold=True,
         color='text_muted', C=C)
    text(slide, x + 0.55, y, 11.0, 0.4, note, font_size=9, color='text_muted', C=C)
    if cite:
        text(slide, x + 0.55, y + 0.22, 11.0, 0.25, cite, font_size=8,
             color='text_muted', C=C)

def base_row(slide, x, y, w, h, bases, colors, base_w=0.3):
    """Per-base sequence row with colored blocks."""
    for i, (b, color) in enumerate(zip(bases, colors)):
        bx = x + i * base_w
        rect(s, bx, y, base_w - 0.02, h, color, C=C)
        text(s, bx, y + h*0.12, base_w - 0.02, h*0.7, b, font_size=9,
             color='white', align='center', C=C)

def legend_item(slide, x, y, color, label):
    rect(s, x, y + 0.05, 0.18, 0.18, color, C=C)
    text(s, x + 0.24, y, 2.5, 0.28, label, font_size=9, color='text_body', C=C)

# ══════════════════════════════════════════════════════════════════
# Page 1 — 封面
# ══════════════════════════════════════════════════════════════════
s = add_slide(PRS)
rect(s, 0, 0, 13.333, 7.5, C['background'], C=C)
rect(s, 0, 0, 13.333, 0.07, C['up_color'], C=C)

text(s, 1.0, 0.8, 8.0, 0.3, 'FRONTIERS IN BIOLOGY', font_size=11,
     color='text_muted', C=C)
text(s, 1.0, 1.2, 11.0, 0.7, 'CRISPR-Cas9 基因编辑', font_size=34,
     color='text_dark', bold=True, C=C)
text(s, 1.0, 2.0, 11.0, 0.5, '从分子机制到技术应用', font_size=16,
     color='text_body', C=C)

# 封面：分子机制示意 (非图片)
# sgRNA 序列比对 + Cas9 + PAM
cover_y = 3.1
text(s, 1.0, cover_y, 2.0, 0.3, 'Target DNA:', font_size=10, bold=True, color='text_dark', C=C)
base_row(s, 3.0, cover_y, 8.5, 0.42,
         list('ATCGTACGTAAAGGGCG'), ['white']*18, 0.3)
# 用彩色区分: 靶标链
base_row(s, 3.0, cover_y, 8.5, 0.42,
         list('ATCGTACGTAAAGGGCG'),
         [C['target_color']]*18, 0.3)

# 引用
text(s, 1.0, 6.6, 11.0, 0.3,
     'Jinek M, et al. Science 2012; 337:816-821  |  doi:10.1126/science.1225829',
     font_size=9, color='text_muted', C=C)

# ══════════════════════════════════════════════════════════════════
# Page 2 — 结果 Figure 1: 双 RNA 依赖切割 (多面板)
# ══════════════════════════════════════════════════════════════════
s = add_slide(PRS)
figure_header(s, 1, 'Cas9 切割依赖双 RNA 引导')

# ── 面板 a: sgRNA 序列比对 + 切割 ──
panel_label(s, 0.7, 0.75, 'a', 'sgRNA-DNA 序列比对与 Cas9 切割')
seq_t = 'ATCGTACGTAAAGGGCG'
seq_n = 'TAGCATGCATTTCCCG'
seq_r = 'AUCGUACG' + 'U' + 'AAAGGGCG'
pam_start = 14
# 靶标链 (5'→3')
text(s, 0.7, 1.15, 1.4, 0.3, '5\'→3\' Target', font_size=8, color='text_muted', C=C)
base_row(s, 2.1, 1.1, 7.0, 0.36, list(seq_t),
         [C['target_color'] if i < 14 else C['pam_color'] for i in range(18)], 0.34)
text(s, 9.5, 1.1, 1.0, 0.3, 'PAM', font_size=8, color=C['pam_color'], bold=True, C=C)
# 非靶标链 (3'→5')
text(s, 0.7, 1.62, 1.4, 0.3, '3\'→5\' Non-target', font_size=8, color='text_muted', C=C)
base_row(s, 2.1, 1.58, 7.0, 0.36, list(seq_n),
         [C['nontarget_color'] if i < 14 else C['pam_color'] for i in range(18)], 0.34)
# sgRNA
text(s, 0.7, 2.1, 1.4, 0.3, '5\'→3\' sgRNA', font_size=8, color='text_muted', C=C)
base_row(s, 2.1, 2.06, 7.0, 0.36, list(seq_r),
         [C['guide_color'] if i < 14 else C['grid_color'] for i in range(18)], 0.34)
# 切割位点标注
text(s, 6.3, 2.5, 2.0, 0.3, '✂', font_size=12, color=C['up_color'], C=C)
text(s, 6.6, 2.8, 3.0, 0.3, 'DSB 切割位点', font_size=8, color=C['up_color'], C=C)
text(s, 3.0, 2.5, 3.0, 0.3, 'PAM 上游 3 bp', font_size=8, color='text_muted', C=C)

# ── 面板 b: 切割效率对比 (数据) ──
panel_label(s, 9.2, 0.75, 'b', '切割效率定量')
bar_chart(s, 9.6, 1.2, [
    ('双 RNA', 0.95, '95%'),
    ('仅 Cas9', 0.03, '3%'),
    ('单 RNA', 0.15, '15%'),
], max_width=3.0, bar_height=0.4, C=C, typo=t, spacing=sp)

# ── 面板 c: 双 RNA 机制图 (RNP 组装流程) ──
panel_label(s, 0.7, 3.35, 'c', '双 RNA 引导组装')
# crRNA + tracrRNA → RNP → 切割
flow_steps = [
    ('crRNA', '靶向信息', C['guide_color']),
    ('tracrRNA', 'Cas9 结合骨架', C['nontarget_color']),
]
x = 0.9
for i, (name, desc, color) in enumerate(flow_steps):
    rect(s, x, 3.75, 2.2, 0.8, color, C=C)
    text(s, x, 3.85, 2.2, 0.3, name, font_size=11, bold=True, color='white', align='center', C=C)
    text(s, x, 4.15, 2.2, 0.3, desc, font_size=8, color='white', align='center', C=C)
    if i == 0:
        text(s, x + 2.2, 4.0, 0.4, 0.3, '+', font_size=16, color='text_dark', align='center', C=C)
    x += 2.6
text(s, x, 4.0, 0.4, 0.3, '→', font_size=16, color='text_dark', align='center', C=C)
rect(s, x + 0.4, 3.75, 2.2, 0.8, C['cas9_color'], C=C)
text(s, x + 0.4, 3.85, 2.2, 0.3, 'RNP 复合体', font_size=11, bold=True, color='white', align='center', C=C)
text(s, x + 0.4, 4.15, 2.2, 0.3, 'Cas9 + 双 RNA', font_size=8, color='white', align='center', C=C)
text(s, x + 2.6, 4.0, 0.4, 0.3, '→', font_size=16, color='text_dark', align='center', C=C)
rect(s, x + 3.0, 3.75, 2.2, 0.8, C['up_color'], C=C)
text(s, x + 3.0, 3.85, 2.2, 0.3, '靶 DNA 切割', font_size=11, bold=True, color='white', align='center', C=C)

# ── 面板 d: 关键参数 (统计) ──
panel_label(s, 9.2, 3.35, 'd', '关键参数')
params = [
    ('95.3±2.1%', '双 RNA 切割效率', C['up_color']),
    ('15.1±3.4%', '单 RNA 切割效率', C['mutant_color']),
    ('20 nt', 'sgRNA 引导序列长度', C['guide_color']),
    ('3 bp', 'PAM 上游切割位点', C['pam_color']),
]
for i, (num, label, color) in enumerate(params):
    y = 3.75 + i * 0.75
    text(s, 9.6, y, 2.0, 0.35, num, font_size=16, bold=True, color=color, C=C)
    text(s, 11.6, y + 0.08, 1.5, 0.35, label, font_size=8, color='text_muted', C=C)

# 图例
legend_item(s, 0.7, 6.3, C['target_color'], '靶 DNA 链')
legend_item(s, 3.3, 6.3, C['nontarget_color'], '非靶 DNA 链')
legend_item(s, 5.9, 6.3, C['guide_color'], 'sgRNA 引导链')
legend_item(s, 8.5, 6.3, C['pam_color'], 'PAM (5\'-NGG-3\')')

panel_note(s, 0.7, 6.65, 'Cas9 是可编程核酸内切酶：sgRNA 20nt 决定靶点，PAM 决定可编辑范围。双 RNA 引导下切割效率显著高于单 RNA。',
           'Jinek M, et al. Science 2012; 337:816-821')

# ══════════════════════════════════════════════════════════════════
# Page 3 — 结果 Figure 2: PAM 要求与切割位点
# ══════════════════════════════════════════════════════════════════
s = add_slide(PRS)
figure_header(s, 2, 'PAM (5\'-NGG-3\') 与切割位点')

# ── 面板 a: PAM 变异扫描 (表格 + 色块) ──
panel_label(s, 0.7, 0.75, 'a', 'PAM 变异对切割的影响')
pam_rows = [
    ('NGG (野生型)', '5\'-NGG-3\'', '切割 ✓', '100%', C['up_color']),
    ('NAG (错配)', '5\'-NAG-3\'', '弱切割', '38%', C['mutant_color']),
    ('NGA (错配)', '5\'-NGA-3\'', '无切割 ✗', '0%', C['down_color']),
    ('NGG (截短)', '截短 PAM', '无切割 ✗', '0%', C['down_color']),
]
x0, y0 = 0.7, 1.2
colw = [2.2, 1.7, 1.3, 1.1]
text(s, x0, y0, colw[0], 0.3, 'PAM 变体', font_size=10, bold=True, color='text_dark', C=C)
text(s, x0+colw[0], y0, colw[1], 0.3, '序列', font_size=10, bold=True, color='text_dark', C=C)
text(s, x0+colw[0]+colw[1], y0, colw[2], 0.3, '切割', font_size=10, bold=True, color='text_dark', C=C)
text(s, x0+colw[0]+colw[1]+colw[2], y0, colw[3], 0.3, '效率', font_size=10, bold=True, color='text_dark', C=C)
for i, (name, pam, res, eff, color) in enumerate(pam_rows):
    y = y0 + 0.38 + i * 0.5
    rect(s, x0, y, 6.3, 0.4, C['bg_tint'] if i%2==0 else C['white'], C=C)
    text(s, x0+0.1, y+0.06, colw[0]-0.1, 0.3, name, font_size=9, color='text_body', C=C)
    text(s, x0+colw[0], y+0.06, colw[1], 0.3, pam, font_size=9, color=color, bold=True, C=C)
    text(s, x0+colw[0]+colw[1], y+0.06, colw[2], 0.3, res, font_size=9, color=color, C=C)
    text(s, x0+colw[0]+colw[1]+colw[2], y+0.06, colw[3], 0.3, eff, font_size=9, color=color, bold=True, C=C)
# 相对效率柱状图
text(s, 0.9, 3.6, 3.0, 0.3, '相对切割效率', font_size=10, bold=True, color='text_dark', C=C)
bar_chart(s, 0.9, 4.0, [
    ('NGG', 1.0, '100%'),
    ('NAG', 0.38, '38%'),
    ('NGA', 0.02, '2%'),
    ('截短', 0.01, '1%'),
], max_width=3.0, bar_height=0.32, C=C, typo=t, spacing=sp)
text(s, 0.9, 5.9, 6.0, 0.3, 'NGG→NAG 效率降至 38%；NGA 完全失活', font_size=10,
     color=C['up_color'], C=C)

# ── 面板 b: 切割位点定位 (逐碱基色块) ──
panel_label(s, 7.2, 0.75, 'b', '切割位点定位于 PAM 上游 3 bp')
# 序列色块: 靶标链 (前 20nt 靶序列, 后 3nt PAM)
seq_blocks = [
    ('G', C['target_color']), ('G', C['target_color']), ('C', C['target_color']),
    ('T', C['target_color']), ('A', C['target_color']), ('G', C['target_color']),
    ('C', C['target_color']), ('G', C['target_color']), ('A', C['target_color']),
    ('G', C['target_color']), ('A', C['target_color']), ('G', C['target_color']),
    ('T', C['target_color']), ('G', C['target_color']), ('G', C['pam_color']),
    ('G', C['pam_color']), ('C', C['pam_color']), ('T', C['grid_color']),
    ('G', C['grid_color']), ('C', C['grid_color']), ('G', C['grid_color']),
]
bx, by = 7.2, 1.4
bw = 0.26
for i, (b, color) in enumerate(seq_blocks):
    rect(s, bx + i*bw, by, bw - 0.02, 0.38, color, C=C)
    text(s, bx + i*bw, by + 0.06, bw - 0.02, 0.26, b, font_size=8,
         color='white', align='center', C=C)
# PAM 标注
text(s, bx + 14*bw, by + 0.45, 1.5, 0.25, 'PAM (NGG)', font_size=8, color=C['pam_color'], C=C)
# 切割位点
text(s, bx + 11*bw, by - 0.3, 2.0, 0.25, '▲ 切割 (PAM 上游 3bp)', font_size=8, color=C['up_color'], C=C)
rect(s, bx + 11*bw + 0.1, by - 0.05, 0.015, 0.5, C['up_color'], C=C)

# 切割产物
text(s, 7.2, 2.5, 5.5, 0.3, '切割产物：平末端 (blunt ends)', font_size=11, bold=True, color='text_dark', C=C)
multiline(s, 7.2, 2.9, 5.5, 1.3, [
    '靶标链与非靶标链在相同位置切断',
    '变性 PAGE 显示两条等长产物带',
    '与 ZFN/TALEN 的粘性末端不同',
], font_size=10, color='text_body', C=C, line_spacing=1.5)

# ── 面板 c: Cas9 结构域 ──
panel_label(s, 0.7, 4.8, 'c', 'SpCas9 功能域')
# 蛋白条带图
doms = [
    ('REC', '识别螺旋', C['guide_color']),
    ('Bridge', '桥接螺旋', C['nontarget_color']),
    ('RuvC', '非靶链切割', C['up_color']),
    ('HNH', '靶链切割', C['mutant_color']),
    ('PI', 'PAM 识别', C['pam_color']),
]
x = 0.9
for i, (name, func, color) in enumerate(doms):
    w = 1.6
    rect(s, x, 5.3, w, 0.55, color, C=C)
    text(s, x, 5.38, w, 0.25, name, font_size=8, bold=True, color='white', align='center', C=C)
    text(s, x, 5.95, w, 0.3, func, font_size=8, color='text_muted', align='center', C=C)
    x += w + 0.12
text(s, 0.9, 6.3, 5.5, 0.3, 'SpCas9 全长 1,368 aa，RuvC/HNH 双结构域切割', font_size=9, color='text_muted', C=C)

# ── 面板 d: 统计 ──
panel_label(s, 7.2, 4.8, 'd', '关键统计')
text(s, 7.2, 5.3, 5.5, 0.35, '约每 8 bp 出现一次 NGG (统计)', font_size=11, bold=True, color=C['up_color'], C=C)
text(s, 7.2, 5.75, 5.5, 0.35, '60-90% 体外编辑效率 (细胞系典型范围)', font_size=11, color='text_body', C=C)
text(s, 7.2, 6.15, 5.5, 0.3, 'PAM 频率高 → 全基因组可编辑位点丰富', font_size=10, color='text_muted', C=C)

panel_note(s, 0.7, 6.65, 'PAM 第三位必须为 G (NGG)；NGA 变异完全丧失活性。PAM 使 Cas9 区分自我 DNA 与非自我 DNA。',
           'Jinek M, et al. Science 2012; 337:816-821')

# ══════════════════════════════════════════════════════════════════
# Page 4 — 结果 Figure 3: 双结构域机制 + DNA 修复路径
# ══════════════════════════════════════════════════════════════════
s = add_slide(PRS)
figure_header(s, 3, 'Cas9 切割机制与 DNA 修复路径')

# ── 面板 a: 机制图 ──
panel_label(s, 0.7, 0.75, 'a', 'Cas9 双结构域切割机制')
# Cas9 蛋白 + 双链
rect(s, 1.0, 2.0, 3.0, 1.8, C['bg_tint'], C=C)
text(s, 1.0, 2.1, 3.0, 0.3, 'Cas9', font_size=12, bold=True, color='text_dark', align='center', C=C)
rect(s, 1.2, 2.6, 1.2, 0.6, C['up_color'], C=C)
text(s, 1.2, 2.7, 1.2, 0.25, 'RuvC', font_size=10, color='white', align='center', C=C)
rect(s, 2.6, 2.6, 1.2, 0.6, C['mutant_color'], C=C)
text(s, 2.6, 2.7, 1.2, 0.25, 'HNH', font_size=10, color='white', align='center', C=C)
text(s, 1.0, 3.4, 3.0, 0.3, 'RuvC→非靶链  HNH→靶链', font_size=9, color='text_body', align='center', C=C)
# 双链 DNA
text(s, 4.5, 1.7, 2.0, 0.3, '靶标链', font_size=9, color=C['target_color'], C=C)
rect(s, 4.5, 2.0, 5.5, 0.12, C['target_color'], C=C)
text(s, 4.5, 2.45, 2.0, 0.3, '非靶标链', font_size=9, color=C['nontarget_color'], C=C)
rect(s, 4.5, 2.75, 5.5, 0.12, C['nontarget_color'], C=C)
text(s, 8.2, 1.65, 2.0, 0.3, '✂', font_size=12, color=C['up_color'], C=C)
rect(s, 8.5, 1.95, 0.015, 1.0, C['up_color'], C=C)
text(s, 8.7, 1.65, 2.5, 0.3, '切割位点', font_size=8, color=C['up_color'], C=C)

# ── 面板 b: DNA 修复路径分叉 (NHEJ vs HDR) ──
panel_label(s, 7.2, 0.75, 'b', 'DNA 修复路径分叉: NHEJ vs HDR')
text(s, 8.5, 1.6, 1.0, 0.35, 'DSB', font_size=14, bold=True, color=C['up_color'], align='center', C=C)
# 左: NHEJ
rect(s, 7.2, 2.1, 2.8, 0.5, C['down_color'], C=C)
text(s, 7.2, 2.2, 2.8, 0.3, 'NHEJ — 非同源末端连接', font_size=10, bold=True, color='white', align='center', C=C)
multiline(s, 7.2, 2.7, 2.8, 1.5, [
    'Ku70/80 结合断端',
    'Lig4/XRCC4 连接',
    '结果: indel 突变',
], font_size=9, color='text_body', C=C, line_spacing=1.4)
# 右: HDR
rect(s, 10.2, 2.1, 2.6, 0.5, C['control_color'], C=C)
text(s, 10.2, 2.2, 2.6, 0.3, 'HDR — 同源定向修复', font_size=10, bold=True, color='white', align='center', C=C)
multiline(s, 10.2, 2.7, 2.6, 1.5, [
    'Rad51 链侵入',
    '同源模板合成',
    '结果: 精准修复',
], font_size=9, color='text_body', C=C, line_spacing=1.4)
text(s, 7.2, 4.4, 5.6, 0.3, '应用：NHEJ→基因敲除 (KO)；HDR→基因敲入 (KI)', font_size=9, color='text_muted', C=C)

# ── 面板 c: 统计大数字 ──
panel_label(s, 0.7, 4.6, 'c', '编辑效率')
stats = [('95%', '双 RNA 切割'), ('60-90%', '细胞系编辑'), ('38%', 'NAG 弱切割')]
for i, (num, label) in enumerate(stats):
    x = 0.9 + i * 2.3
    text(s, x, 5.2, 2.0, 0.4, num, font_size=18, bold=True, color=C['up_color'], C=C)
    text(s, x, 5.65, 2.0, 0.3, label, font_size=9, color='text_muted', C=C)

# ── 面板 d: 修复应用 ──
panel_label(s, 7.2, 4.6, 'd', '基因编辑应用')
multiline(s, 7.2, 5.1, 5.6, 1.4, [
    '敲除 (KO): NHEJ 引入 indel → 移码失活',
    '敲入 (KI): HDR 供体模板 → 精确替换',
    '碱基编辑 (BE): 无需 DSB，C→T 单碱基',
], font_size=10, color='text_body', C=C, line_spacing=1.5)

panel_note(s, 0.7, 6.65, 'Cas9 通过 RuvC (非靶链) 与 HNH (靶链) 双结构域切割双链，产生平末端 DSB。细胞通过 NHEJ 或 HDR 修复，决定编辑类型。',
           'Jinek M, et al. Science 2012; 337:816-821')

# ══════════════════════════════════════════════════════════════════
# Page 5 — 结果 Figure 4: 靶向特异性 (错配)
# ══════════════════════════════════════════════════════════════════
s = add_slide(PRS)
figure_header(s, 4, '靶向特异性：种子区错配最敏感')

# ── 面板 a: 错配敏感性 (折线图) ──
panel_label(s, 0.7, 0.75, 'a', '错配位置 vs 相对切割效率')
native_chart(s, 0.9, 1.1, 7.4, 4.4, 'line',
             categories=['PAM端1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14远端'],
             series=[{'name': '相对切割效率', 'values': [0.95, 0.88, 0.75, 0.6, 0.4, 0.25, 0.1, 0.05, 0.08, 0.12, 0.2, 0.35, 0.5, 0.7]}],
             style={'show_legend': False, 'gridlines': 'major_y',
                    'value_axis_title': '相对切割效率', 'color_scheme': [C['up_color']]},
             C=C)
text(s, 1.2, 5.7, 7.0, 0.3, '种子区 (PAM 邻近) → 错配最敏感', font_size=9, color=C['up_color'], C=C)

# ── 面板 b: 关键数据 ──
panel_label(s, 8.7, 0.75, 'b', '错配抑制效应')
text(s, 8.7, 1.3, 4.2, 0.4, '种子区单碱基错配：', font_size=11, bold=True, color='text_dark', C=C)
text(s, 8.7, 1.7, 4.2, 0.5, '切割效率降低 >90%', font_size=16, bold=True, color=C['up_color'], C=C)
text(s, 8.7, 2.3, 4.2, 0.4, '远端错配：容忍度较高', font_size=11, color='text_body', C=C)
rect(s, 8.7, 2.8, 4.2, 0.015, C['divider'], C=C)
multiline(s, 8.7, 3.0, 4.2, 1.5, [
    'Cas9 先通过 PAM 识别锚定',
    'crRNA:DNA 配对沿 5\'→3\' 检查',
    '种子区决定 R-loop 稳定性',
], font_size=10, color='text_body', C=C, line_spacing=1.5)

# ── 面板 c: 特异性统计 ──
panel_label(s, 0.7, 5.2, 'c', '特异性指标')
bar_chart(s, 1.0, 5.6, [
    ('种子区', 0.08, '8%'),
    ('中间区', 0.45, '45%'),
    ('远端', 0.72, '72%'),
    ('完全匹配', 1.0, '100%'),
], max_width=4.0, bar_height=0.32, C=C, typo=t, spacing=sp)

# ── 面板 d: 机制结论 ──
panel_label(s, 8.7, 5.2, 'd', '结论')
multiline(s, 8.7, 5.6, 4.2, 1.0, [
    '种子区 (PAM 邻近 10-12 nt) 是特异性核心',
    '该区错配阻止 R-loop 形成，切割被抑制',
], font_size=10, color='text_body', C=C, line_spacing=1.5)

panel_note(s, 0.7, 6.65, '错配敏感性随位置变化：PAM 邻近的种子区对错配最敏感，决定 Cas9 靶向特异性。',
           'Jinek M, et al. Science 2012; 337:816-821')

clean_save(PRS, 'showcase/crispr-cas9-2012/CRISPR-Cas9_Science_2012.pptx')
print('DONE')
