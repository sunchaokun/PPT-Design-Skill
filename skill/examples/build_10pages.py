"""Build Mode — AI 企业服务年度战略汇报 (10 pages)

Page structure (4 distinct layouts, no 3 consecutive same-structure):
  1. hero_slide          — 封面 (深蓝商务 + 金色点缀)
  2. section_divider     — 01 市场洞察
  3. page_header + kpi_card grid (2x2) — 市场规模与增长
  4. page_header + native_chart (bar)  — 行业收入对比
  5. section_divider     — 02 产品与技术
  6. page_header + code_block          — 核心引擎架构
  7. page_header + highlight_cards     — 三大产品能力
  8. section_divider     — 03 商业落地
  9. page_header + comparison_bars     — 客户价值对比
  10. cta_slide          — 结尾行动号召
"""

from ppt_pro_max.build_helpers import *

# ── Design Token (统一深蓝商务 + 金色点缀) ─────────────────────────
C = {
    'primary': '#1E3A5F',    # 深蓝
    'accent': '#C9A96E',     # 金色
    'muted': '#5B7BA6',      # 蓝灰
    'light': '#D6E4F0',      # 浅蓝
    'white': '#FFFFFF',
    'background': '#F8FAFC', # 浅灰蓝底
    'card_bg': '#FFFFFF',
    'bg_tint': '#F1F5F9',
    'text_dark': '#1A2B3C',
    'text_body': '#37474F',
    'text_muted': '#78909C',
    'divider': '#E0E8F0',
    'font_heading': 'Georgia',
    'font_body': 'Calibri',
    'font_cjk': '微软雅黑',  # 中文必填
}

t = TYPOGRAPHY['cjk_professional']   # 中文 body=14pt
sp = SPACING['professional']

prs = Presentation()
set_widescreen(prs)
set_theme_colors(prs, C)   # 将 C 调色板写入 PowerPoint 主题色

# ══════════════════════════════════════════════════════════════════
# Page 1 — 封面 (视觉锚点: 金色圆环装饰 + 数据节点)
# ══════════════════════════════════════════════════════════════════
s = add_slide(prs)
hero_slide(s, '让企业智能触手可及', 'AI 企业服务 2026 年度战略汇报', C=C, typo=t)

# 视觉锚点: 右上金色圆环 + 蓝灰小圆点缀 (同 palette, 完全在画布内)
donut(s, 10.9, 1.2, 1.6, C['accent'], C=C, line=None)
oval(s, 11.3, 2.2, 0.3, 0.3, C['light'], C=C)
oval(s, 11.6, 0.9, 0.2, 0.2, C['light'], C=C)
# 左下角金色小圆装饰 (画布内)
oval(s, 0.3, 6.1, 0.5, 0.5, C['accent'], C=C)

# 底部金色点缀条
rect(s, 0, 6.9, 13.333, 0.08, C['accent'], C=C)
text(s, 1.2, 6.3, 10.0, 0.4, '战略发展部  |  2026 年 8 月', font_size=12,
     color='light', C=C)

# ══════════════════════════════════════════════════════════════════
# Page 2 — 章节分隔 01 市场洞察
# ══════════════════════════════════════════════════════════════════
s = add_slide(prs)
section_divider(s, 1, '市场洞察', C=C, typo=t)
text(s, 1.2, 4.8, 10.0, 0.5, '企业级 AI 市场进入高速增长期', font_size=16,
     color='light', C=C)

# ══════════════════════════════════════════════════════════════════
# Page 3 — 市场规模与增长 (KPI 2x2)
# ══════════════════════════════════════════════════════════════════
s = add_slide(prs)
page_header(s, '市场规模与增长', '中国企业级 AI 支出持续扩大', C=C, typo=t, spacing=sp)

kpi_card(s, 0.65, 1.9, 3.9, 1.4, '1,280亿', '企业级 AI 市场规模', '年增 38.5%', C=C, typo=t)
kpi_card(s, 4.85, 1.9, 3.9, 1.4, '87.6%', '企业 AI 采纳率', '同比 +12.3%', C=C, typo=t)
kpi_card(s, 0.65, 3.55, 3.9, 1.4, '46亿', 'AI 相关岗位缺口', '供需失衡', C=C, typo=t)
kpi_card(s, 4.85, 3.55, 3.9, 1.4, '3.2倍', 'AI 应用投资回报', '平均 ROI', C=C, typo=t)

text(s, 8.9, 1.9, 3.8, 1.4, '数据来源：IDC / Gartner 2025',
     font_size=11, color='text_muted', C=C)

# ══════════════════════════════════════════════════════════════════
# Page 4 — 行业收入对比 (native bar chart)
# ══════════════════════════════════════════════════════════════════
s = add_slide(prs)
page_header(s, '行业 AI 收入对比', '2025 年各行业 AI 解决方案支出（亿元）', C=C, typo=t, spacing=sp)

native_chart(s, 0.65, 2.0, 12.0, 4.6, 'bar',
             categories=['金融', '制造', '医疗', '零售', '能源', '教育'],
             series=[{'name': 'AI 支出', 'values': [520, 410, 320, 280, 190, 150]}],
             style={
                 'show_legend': False,
                 'show_labels': True,
                 'gridlines': 'major_y',
                 'color_scheme': [C['primary'], C['accent'], C['muted']],
                 'number_format': '#,##0',
             },
             C=C)

# ══════════════════════════════════════════════════════════════════
# Page 5 — 章节分隔 02 产品与技术
# ══════════════════════════════════════════════════════════════════
s = add_slide(prs)
section_divider(s, 2, '产品与技术', C=C, typo=t)
text(s, 1.2, 4.8, 10.0, 0.5, '从模型能力到企业级落地的全栈方案', font_size=16,
     color='light', C=C)

# ══════════════════════════════════════════════════════════════════
# Page 6 — 核心引擎架构 (code block)
# ══════════════════════════════════════════════════════════════════
s = add_slide(prs)
page_header(s, '核心引擎架构', '企业智能引擎一次调用完成推理全流程', C=C, typo=t, spacing=sp)

code_block(s, 0.65, 1.9, 7.5, 4.6, [
    'from enterprise_ai import Engine',
    '',
    '# 初始化企业智能引擎',
    'engine = Engine(region="cn-east", tier="enterprise")',
    '',
    'result = engine.infer(',
    '    task="report_generation",',
    '    context=knowledge_base.query("Q2 财报"),',
    '    model="deepseek-r1",',
    '    guardrails=True,   # 企业安全护栏',
    ')',
    '',
    'print(result.summary)  # → 结构化业务洞察',
], language='python', C=C, typo=t)

text(s, 8.35, 1.9, 4.3, 0.5, '架构亮点', font_size=t.h3, bold=True,
     color='primary', C=C)
multiline(s, 8.35, 2.5, 4.3, 3.6, [
    '多模型路由，按场景智能选型',
    '内置 RAG 知识库接入',
    '企业级安全护栏默认开启',
    '全链路可观测与审计',
], font_size=14, color='text_body', C=C, line_spacing=1.6)

# ══════════════════════════════════════════════════════════════════
# Page 7 — 三大产品能力 (highlight_cards)
# ══════════════════════════════════════════════════════════════════
s = add_slide(prs)
page_header(s, '三大产品能力', '覆盖企业智能化关键场景', C=C, typo=t, spacing=sp)

highlight_cards(s, 0.65, 2.2, [
    ('智能报告', '财报、行业报告、尽调材料自动生成，平均节省 12 小时/份', C['primary']),
    ('智能客服', '7×24 多轮对话，一次解决率提升至 92%，人力成本降低 40%', C['accent']),
    ('数据洞察', '自然语言查询 BI 数据，决策链路从 3 天缩短到 3 分钟', C['muted']),
], total_width=12.0, C=C, typo=t, spacing=sp)

text(s, 0.65, 4.2, 12.0, 0.5, '三大能力共享同一引擎底座，一次接入、全场景复用',
     font_size=14, color='text_muted', C=C)

# ══════════════════════════════════════════════════════════════════
# Page 8 — 章节分隔 03 商业落地
# ══════════════════════════════════════════════════════════════════
s = add_slide(prs)
section_divider(s, 3, '商业落地', C=C, typo=t)
text(s, 1.2, 4.8, 10.0, 0.5, '已服务 200+ 行业头部客户', font_size=16,
     color='light', C=C)

# ══════════════════════════════════════════════════════════════════
# Page 9 — 客户价值对比 (comparison_bars)
# ══════════════════════════════════════════════════════════════════
s = add_slide(prs)
page_header(s, '客户价值对比', 'AI 化改造前后关键指标', C=C, typo=t, spacing=sp)

comparison_bars(s, 2.0, 2.2, [
    ('处理效率', '25%', '92%', 0.25, 0.92),
    ('人工成本', '80%', '48%', 0.80, 0.48),
    ('响应时间', '95%', '30%', 0.95, 0.30),
    ('错误率', '15%', '5%', 0.15, 0.05),
], max_width=5.0, C=C, typo=t, spacing=sp)

# 图例
text(s, 3.2, 5.2, 3.0, 0.3, '■ 改造前', font_size=12, color='primary', C=C)
text(s, 5.2, 5.2, 3.0, 0.3, '■ 改造后', font_size=12, color='accent', C=C)

text(s, 8.9, 2.2, 3.8, 0.5, '客户案例', font_size=t.h3, bold=True,
     color='primary', C=C)
multiline(s, 8.9, 2.8, 3.8, 3.5, [
    '某股份制银行：月均节省 2,000 人时',
    '某头部制造企业：质检通过率 +18%',
    '某连锁零售：客服成本下降 40%',
    '3 个月内实现 ROI 转正',
], font_size=14, color='text_body', C=C, line_spacing=1.6)

# ══════════════════════════════════════════════════════════════════
# Page 10 — 结尾 CTA (视觉锚点: 金色圆环呼应封面)
# ══════════════════════════════════════════════════════════════════
s = add_slide(prs)
cta_slide(s, '与我们一起，开启企业智能之旅',
          '现在预约演示，获取专属企业 AI 落地方案与行业白皮书', C=C, typo=t)

# 视觉锚点: 呼应封面金色圆环 (画布内)
donut(s, 2.4, 1.2, 1.6, C['accent'], C=C, line=None)
oval(s, 2.8, 2.2, 0.3, 0.3, C['light'], C=C)
oval(s, 3.1, 0.9, 0.2, 0.2, C['light'], C=C)
oval(s, 12.6, 6.1, 0.5, 0.5, C['accent'], C=C)
# 底部金色点缀条
rect(s, 0, 6.9, 13.333, 0.08, C['accent'], C=C)
text(s, 1.2, 6.2, 10.0, 0.4, '联系：ai-strategy@company.com  |  400-888-0000',
     font_size=12, color='light', C=C)

clean_save(prs, 'output/build_10pages.pptx')
print('DONE: output/build_10pages.pptx')
