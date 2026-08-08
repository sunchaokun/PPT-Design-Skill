#!/usr/bin/env python3
"""DeepSeek B 轮融资商业计划书 — Build Mode（逐像素控制，暗色 AI 科技风）.

Run: python build_deepseek_bp.py
Output: deepseek_build.pptx
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "src"))
from ppt_pro_max.build_helpers import (
    Presentation, Inches, Typography, add_slide, add_glow, clean_save, code_block,
    donut_chart, gradient_text,
    grid_background, kpi_card, native_chart, neon_border, oval, page_header, rect, rrect, section_divider,
    set_dark_theme, set_widescreen, text,
)

# CJK 调整的排版尺度（body/caption/micro 抬升，避免 CJK 偏小）
T = Typography(hero=46, h1=30, h2=20, h3=16, body=14, caption=12, micro=11)

# ── 设计令牌：暗色 AI 科技主题 ────────────────────────────────
C = {
    "primary": "#4F46E5",
    "accent": "#22D3EE",
    "muted": "#1E293B",
    "light": "#818CF8",
    "white": "#FFFFFF",
    "background": "#0B1020",
    "card_bg": "#141B33",
    "card_line": "#2A3650",
    "text_dark": "#E2E8F0",
    "text_body": "#B6C2D6",
    "text_muted": "#64748B",
    "divider": "#2A3650",
    "success": "#34D399",
    "danger": "#F472B6",
    "font_heading": "Orbitron",
    "font_body": "JetBrains Mono",
    "font_cjk": "微软雅黑",
}

def tech_bg(slide):
    """统一技术底：深色背景 + 网格 + 顶部霓虹线 + 底部进度条。"""
    rect(slide, 0, 0, 13.333, 7.5, C["background"])
    grid_background(slide, spacing=0.85, color=C["divider"], alpha=7)
    rect(slide, 0, 0, 13.333, 0.06, C["accent"])
    rect(slide, 0, 7.45, 13.333, 0.05, C["divider"])

def glow_orb(slide, x, y, d, color, alpha=14):
    sh = oval(slide, x, y, d, d, color)
    add_glow(sh, color=color, size_pt=18, alpha_pct=30)

def tech_card(slide, left, top, width, height, fill="#151B30"):
    return rrect(slide, left, top, width, height, fill)

def main():
    prs = Presentation()
    set_widescreen(prs)
    set_dark_theme(prs, C)

    # ═══ 1. 封面 ═══
    s = add_slide(prs)
    rect(s, 0, 0, 13.333, 7.5, "#0A0E1E")
    grid_background(s, spacing=0.9, color="#2A3650", alpha=8)
    rect(s, 0, 0, 13.333, 0.07, C["accent"])
    # 两个霓虹光晕
    glow_orb(s, 10.6, -1.2, 3.6, C["accent"])
    glow_orb(s, -1.5, 5.4, 4.2, C["primary"])
    # 左上角 mono 标签
    text(s, 0.9, 0.55, 6, 0.4, "// DEEPSEEK · B ROUND · ¥50,000,000,000",
         font_size=13, color="accent", font_name="JetBrains Mono", C=C)
    # 主标题（Orbitron 拉丁 + CJK 回落）
    gradient_text(s, 1.0, 2.15, 11.3, 1.5, "让智能的成本，下降一个数量级",
                  stops=[("#22D3EE", 0), ("#818CF8", 60), ("#4F46E5", 100)],
                  font_size=46, bold=True, font_name="Orbitron", cjk_font="微软雅黑", align="left")
    rect(s, 1.05, 3.85, 1.6, 0.06, C["accent"])
    text(s, 1.0, 4.1, 11.3, 0.6,
         "DeepSeek · 以极致效率重塑大模型竞争格局", font_size=20,
         color="text_body", font_name="微软雅黑", C=C)
    text(s, 1.0, 4.75, 11.3, 0.5,
         "B 轮融资商业计划书  |  融资规模 500 亿元人民币", font_size=15,
         color="text_muted", font_name="微软雅黑", C=C)
    # 底部技术指标条
    rect(s, 1.0, 5.6, 0.12, 0.7, C["accent"])
    text(s, 1.3, 5.6, 4, 0.35, "综合能力全球第一梯队", font_size=14, color="text_dark", font_name="微软雅黑", C=C)
    text(s, 1.3, 5.95, 4, 0.3, "Capability · Top 3 Globally", font_size=11, color="text_muted", font_name="JetBrains Mono", C=C)
    rect(s, 5.2, 5.6, 0.12, 0.7, C["success"])
    text(s, 5.5, 5.6, 4, 0.35, "推理成本低一个数量级", font_size=14, color="text_dark", font_name="微软雅黑", C=C)
    text(s, 5.5, 5.95, 4, 0.3, "1/20 Cost vs. Closed Models", font_size=11, color="text_muted", font_name="JetBrains Mono", C=C)
    rect(s, 9.4, 5.6, 0.12, 0.7, C["light"])
    text(s, 9.7, 5.6, 3.6, 0.35, "开源生态全球第一梯队", font_size=14, color="text_dark", font_name="微软雅黑", C=C)
    text(s, 9.7, 5.95, 3.6, 0.3, "MIT · 1,500+ Derivatives", font_size=11, color="text_muted", font_name="JetBrains Mono", C=C)

    # ═══ 2. 愿景 ═══
    s = add_slide(prs)
    tech_bg(s)
    page_header(s, "愿景：金钱以外的事情", "VISION — 以善意与技术普惠智能",
                C=C, left=0.9, typo=T)
    rrect(s, 0.9, 1.6, 11.5, 1.3, "#0F172A")
    text(s, 1.2, 1.75, 10.9, 0.8,
         "「我们是怀着一个对这个世界非常大的善意来做这个事情，这是一个金钱以外的事情。」",
         font_size=18, color="text_dark", bold=True, font_name="微软雅黑", C=C)
    text(s, 1.2, 2.5, 10.9, 0.35, "— 梁文锋 · DeepSeek 创始人", font_size=13, color="text_muted", font_name="微软雅黑", C=C)
    # 三大理念卡片
    cards = [
        ("愿景驱动", "「一个公司最重要的是愿景」—— 靠愿景而非规章制度", "accent"),
        ("克制即开源", "「开源属于克制的一部分，越克制越容易做成」", "success"),
        ("效率即普惠", "「让大家都用得起」—— 低成本是愿景，不是商业策略", "light"),
    ]
    card_w = 3.6
    for i, (t, desc, color) in enumerate(cards):
        x = 0.9 + i * (card_w + 0.35)
        rrect(s, x, 3.2, card_w, 2.5, C["card_bg"])
        rect(s, x, 3.2, card_w, 0.09, color)
        text(s, x + 0.3, 3.5, card_w - 0.6, 0.5, t, font_size=19, color="text_dark", bold=True, font_name="微软雅黑", C=C)
        text(s, x + 0.3, 4.1, card_w - 0.6, 1.4, desc, font_size=13, color="text_body", font_name="微软雅黑", C=C)

    # ═══ 3. Section 01 ═══
    s = add_slide(prs)
    section_divider(s, "01", "竞争格局剖析", C=C, grouped=False)

    # ═══ 4. 全球版图 ═══
    s = add_slide(prs)
    tech_bg(s)
    page_header(s, "全球基座模型竞争版图", "COMPETITIVE LANDSCAPE — 头部模型综合能力",
                C=C, left=0.9, typo=T)
    native_chart(s, 0.9, 1.7, 7.6, 4.6, "bar",
                 categories=["DeepSeek", "GPT-5", "Claude 4", "Gemini", "Qwen", "Doubao"],
                 series=[{"name": "综合能力", "values": [92, 95, 93, 91, 84, 82]}],
                 style={"show_legend": False, "color_scheme": "brand", "show_labels": True},
                 C=C)
    # 右侧 KPI
    kpi_card(s, 8.9, 1.7, 3.5, 1.1, "92", "综合能力指数", trend="全球前三", trend_up=True, C=C, typo=T, grouped=False)
    kpi_card(s, 8.9, 2.95, 3.5, 1.1, "1/20", "推理成本倍差", trend="成本领先", trend_up=True, C=C, typo=T, grouped=False)
    kpi_card(s, 8.9, 4.2, 3.5, 1.1, "Top 3", "开发者心智份额", trend="开源生态", trend_up=True, C=C, typo=T, grouped=False)

    # ═══ 5. 竞争维度 ═══
    s = add_slide(prs)
    tech_bg(s)
    page_header(s, "五大竞争维度，DeepSeek 已占其三", "FIVE DIMENSIONS — Where We Lead",
                C=C, left=0.9, typo=T)
    dims = [
        ("能力", "92/100", "第一梯队，编码与推理持平最强闭源", "accent"),
        ("成本", "1/20", "推理成本仅为头部闭源竞品二十分之一", "success"),
        ("开源", "MIT", "全量开源，全球衍生模型 1,500+", "light"),
        ("生态", "Top 3", "Hugging Face 影响力、框架原生集成领先", "accent"),
        ("商业化", "领跑", "API 价格战主动方，以价换量", "light"),
    ]
    for i, (name, val, desc, color) in enumerate(dims):
        x = 0.9
        y = 1.7 + i * 1.08
        rect(s, x, y, 0.14, 0.8, color)
        text(s, x + 0.3, y, 1.2, 0.8, name, font_size=18, color="text_dark", bold=True, font_name="微软雅黑", C=C)
        text(s, x + 1.6, y, 2.2, 0.8, val, font_size=20, color=color, bold=True, font_name="Orbitron", C=C)
        text(s, x + 4.0, y + 0.12, 8.4, 0.6, desc, font_size=14, color="text_body", font_name="微软雅黑", C=C)

    # ═══ 6. Section 02 ═══
    s = add_slide(prs)
    section_divider(s, "02", "DeepSeek 的技术内核", C=C, grouped=False)

    # ═══ 7. 三大技术突破 ═══
    s = add_slide(prs)
    tech_bg(s)
    page_header(s, "三大技术护城河", "TECH MOAT — 架构 · 范式 · 工程",
                C=C, left=0.9, typo=T)
    techs = [
        ("01", "MoE 架构工程化", "671B 参数、激活 37B 的稀疏专家架构，训练成本仅为同规模密集模型的零头。", "accent"),
        ("02", "强化学习新范式", "R1 以纯强化学习激发推理涌现，论文登上 Nature 封面，方法论被全球跟进。", "success"),
        ("03", "极致成本工程", "算子级优化 + 训练框架自研，万卡集群 MFU 达行业顶尖水平。", "light"),
    ]
    for i, (num, t, desc, color) in enumerate(techs):
        x = 0.9 + i * (3.9 + 0.3)
        rrect(s, x, 1.7, 3.9, 4.6, C["card_bg"])
        neon_border(s, x, 1.7, 3.9, 4.6, color=color, radius=0.12)
        text(s, x + 0.35, 2.1, 3.2, 1.0, num, font_size=44, color=color, bold=True, font_name="Orbitron", C=C)
        rect(s, x + 0.35, 3.15, 0.7, 0.05, color)
        text(s, x + 0.35, 3.35, 3.2, 0.8, t, font_size=20, color="text_dark", bold=True, font_name="微软雅黑", C=C)
        text(s, x + 0.35, 4.25, 3.2, 1.8, desc, font_size=14, color="text_body", font_name="微软雅黑", C=C)

    # ═══ 8. 推理成本对比 ═══
    s = add_slide(prs)
    tech_bg(s)
    page_header(s, "推理成本：一个数量级的差距", "COST — Per 1M Output Tokens (USD)",
                C=C, left=0.9, typo=T)
    native_chart(s, 0.9, 1.7, 7.6, 4.6, "bar",
                 categories=["DeepSeek", "Qwen-Max", "Gemini", "GPT-5", "Claude 4"],
                 series=[{"name": "价格", "values": [0.14, 0.5, 1.25, 2.5, 3.0]}],
                 style={"show_legend": False, "color_scheme": "brand", "show_labels": True},
                 C=C)
    code_block(s, 8.8, 1.7, 3.6, 4.6, [
        "deepseek API",
        "",
        "POST /chat/completions",
        "model: deepseek-chat",
        "input : $0.27 / 1M tokens",
        "output: $0.14 / 1M tokens",
        "",
        "// 一个数量级的价差",
        "// = 开发者心智的起点",
    ], language="shell", C=C, typo=T, grouped=False)

    # ═══ 9. 开源生态 ═══
    s = add_slide(prs)
    tech_bg(s)
    page_header(s, "开源战略：克制，是最好的进攻", "OPEN SOURCE — MIT · 克制 · 普惠",
                C=C, left=0.9, typo=T)
    code_block(s, 0.9, 1.7, 5.4, 4.4, [
        "MIT License",
        "",
        "Copyright (c) 2025 DeepSeek",
        "",
        "Permission is hereby granted,",
        "free of charge, to any person",
        "obtaining a copy of this software",
        "and associated documentation",
        "files (the \"Software\"), to deal in",
        "the Software without restriction.",
    ], language="license", C=C, typo=T, grouped=False)
    # 生态数据
    stats = [
        ("1,500+", "开源衍生模型", "事实上的技术标准", "accent"),
        ("#1", "开发者心智增速", "Hugging Face 全球前列", "success"),
        ("100%", "三方框架支持", "vLLM / SGLang / Ollama 原生集成", "light"),
    ]
    for i, (num, t, desc, color) in enumerate(stats):
        x = 6.7
        y = 1.7 + i * 1.62
        rect(s, x, y, 0.14, 1.3, color)
        text(s, x + 0.35, y, 2.2, 0.7, num, font_size=34, color=color, bold=True, font_name="Orbitron", C=C)
        text(s, x + 0.35, y + 0.75, 4.5, 0.4, t, font_size=16, color="text_dark", bold=True, font_name="微软雅黑", C=C)
        text(s, x + 0.35, y + 1.12, 5.6, 0.4, desc, font_size=12, color="text_muted", font_name="微软雅黑", C=C)

    # ═══ 10. Section 03 ═══
    s = add_slide(prs)
    section_divider(s, "03", "B 轮融资计划", C=C, grouped=False)

    # ═══ 11. 融资 500 亿 ═══
    s = add_slide(prs)
    tech_bg(s)
    page_header(s, "融资规模：500 亿元人民币", "B ROUND — CAPITAL RAISING ¥50B",
                C=C, left=0.9, typo=T)
    gradient_text(s, 0.9, 1.7, 4.0, 1.4, "500 亿", stops=[("#22D3EE", 0), ("#818CF8", 100)],
                  font_size=72, bold=True, font_name="Orbitron", align="left")
    text(s, 0.9, 3.2, 4.0, 0.5, "RMB · B 轮融资规模", font_size=16, color="text_body", font_name="微软雅黑", C=C)
    # 四个用途 KPI
    uses = [
        ("45%", "算力基础设施", "225 亿元 · 10 万卡集群", "accent"),
        ("25%", "模型研发", "125 亿元 · V4/R2 下一代", "success"),
        ("12%", "生态与开源", "60 亿元 · 社区与开发者", "light"),
        ("13%", "商业化与人才", "65 亿元 · 团队翻倍至 2,000 人", "accent"),
    ]
    for i, (pct, t, desc, color) in enumerate(uses):
        x = 5.3 + i * 1.95
        rrect(s, x, 1.8, 1.75, 3.2, C["card_bg"])
        rect(s, x, 1.8, 1.75, 0.09, color)
        text(s, x + 0.2, 2.2, 1.35, 0.8, pct, font_size=30, color=color, bold=True, font_name="Orbitron", C=C)
        text(s, x + 0.2, 3.0, 1.35, 0.8, t, font_size=13, color="text_dark", bold=True, font_name="微软雅黑", C=C)
        text(s, x + 0.2, 3.7, 1.35, 1.2, desc, font_size=11, color="text_muted", font_name="微软雅黑", C=C)
    # 底部要点
    text(s, 0.9, 5.6, 11.5, 0.5,
         "估值对标全球顶级基座模型 B 轮区间 · 拟引入战略产业资本与长线财务投资人 · 聚焦算力、研发、生态三大战场",
         font_size=13, color="text_body", font_name="微软雅黑", C=C)

    # ═══ 12. 资金用途环形图 ═══
    s = add_slide(prs)
    tech_bg(s)
    page_header(s, "资金用途分配", "ALLOCATION — 500 亿元用途占比",
                C=C, left=0.9, typo=T)
    donut_chart(s, 4.3, 3.9, 2.2, 1.1, [
        ("算力基础设施", "45%", "#22D3EE"),
        ("模型研发", "25%", "#4F46E5"),
        ("商业化与人才", "13%", "#818CF8"),
        ("生态与开源", "12%", "#34D399"),
        ("安全与治理", "5%", "#F472B6"),
    ], C=C, native=True)
    # 右侧明细
    details = [
        ("算力基建 · 225 亿", "自建 3 个智算中心，新增 10 万卡集群；国产芯片多路线适配", "accent"),
        ("研发与生态 · 185 亿", "V4/R2 下一代基模、多模态与持续学习；开源社区与人才引进", "success"),
        ("商业化 · 90 亿", "全球化部署、头部客户拓展、开发者增长", "light"),
    ]
    for i, (t, desc, color) in enumerate(details):
        x = 8.2
        y = 1.7 + i * 1.55
        rect(s, x, y, 0.14, 1.2, color)
        text(s, x + 0.3, y, 4.3, 0.45, t, font_size=16, color="text_dark", bold=True, font_name="微软雅黑", C=C)
        text(s, x + 0.3, y + 0.5, 4.2, 0.7, desc, font_size=12, color="text_body", font_name="微软雅黑", C=C)

    # ═══ 13. CTA ═══
    s = add_slide(prs)
    rect(s, 0, 0, 13.333, 7.5, "#0A0E1E")
    grid_background(s, spacing=0.9, color="#2A3650", alpha=8)
    glow_orb(s, 9.8, -1.5, 4.5, C["accent"])
    glow_orb(s, -2.0, 4.8, 5.0, C["primary"])
    text(s, 0.9, 0.55, 6, 0.4, "// 用最低的成本，赢得最宽的生态", font_size=13, color="accent", font_name="JetBrains Mono", C=C)
    gradient_text(s, 1.0, 2.3, 11.3, 1.4, "让智能普惠到每一个人", stops=[("#22D3EE", 0), ("#818CF8", 100)],
                  font_size=46, bold=True, font_name="Orbitron", cjk_font="微软雅黑", align="left")
    rect(s, 1.05, 3.9, 1.6, 0.06, C["accent"])
    text(s, 1.0, 4.15, 11.3, 0.6,
         "500 亿元融资 · 联合主承销 / 战略投资人开放洽谈中", font_size=18,
         color="text_body", font_name="微软雅黑", C=C)
    text(s, 1.0, 4.85, 11.3, 0.5,
         "我们怀着对这个世界非常大的善意做这件事——这是金钱以外的事情。",
         font_size=14, color="text_muted", font_name="微软雅黑", C=C)
    # CTA 按钮
    rrect(s, 1.0, 5.7, 3.4, 0.7, C["accent"])
    text(s, 1.0, 5.9, 3.4, 0.4, "进入尽调环节", font_size=16, color="#0A0E1E", bold=True, align="center", font_name="微软雅黑", C=C)

    out = "test_output/deepseek_build/deepseek_build.pptx"
    os.makedirs(os.path.dirname(out), exist_ok=True)
    clean_save(prs, out)
    print("saved", out, "slides:", len(prs.slides._sldIdLst))

if __name__ == "__main__":
    main()
