"""Evaluation 001: Chinese enterprise strategy deck, Build Mode."""

from pathlib import Path

from pptx_designer import Presentation
from pptx_designer.tools.charts import bar_chart
from pptx_designer.tools.layout import page_header, page_number, top_bar
from pptx_designer.tools.shapes import rect, rrect
from pptx_designer.tools.text import multiline, text


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output" / "enterprise_strategy_eval.pptx"

C = {
    "primary": "#14B8A6",
    "accent": "#14B8A6",
    "risk": "#F97316",
    "background": "#0F172A",
    "card_bg": "#172033",
    "card": "#172033",
    "border": "#334155",
    "text_dark": "#F8FAFC",
    "text_body": "#CBD5E1",
    "text_muted": "#94A3B8",
    "divider": "#334155",
}


def base(prs, title, subtitle, number):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["background"], C=C)
    top_bar(slide, C["primary"], width=13.333, height=0.07, C=C)
    if title:
        page_header(slide, title, subtitle, C=C, left=0.7, width=11.9)
    page_number(slide, number, 6, C=C)
    return slide


def main():
    prs = Presentation()

    # 1. Hook: strategic tension, intentionally sparse.
    slide = base(prs, "2027 增长战略", "从增长速度，转向增长方式", 1)
    text(slide, 0.75, 2.1, 7.8, 1.25, "增长仍在发生，\n但速度已经不够。", font_size=34, bold=True, color="text_dark", C=C)
    text(slide, 0.78, 4.35, 5.6, 0.55, "管理层战略讨论 / 10 分钟决策版", font_size=18, color="text_body", C=C)
    rect(slide, 0.78, 5.35, 2.4, 0.08, fill=C["primary"], C=C)
    text(slide, 10.1, 5.1, 2.4, 0.9, "2026 → 2027", font_size=22, bold=True, color="primary", C=C, align="right")

    # 2. Diagnosis: make the supplied comparison the page's main visual.
    slide = base(prs, "增长压力已经明确", "2026 增长 12%，低于行业平均 18%", 2)
    text(slide, 0.9, 1.85, 2.2, 0.35, "公司 2026", font_size=16, color="text_muted", C=C)
    text(slide, 0.9, 2.25, 2.6, 0.9, "12%", font_size=42, bold=True, color="primary", C=C)
    rect(slide, 0.95, 3.35, 4.8, 0.28, fill=C["card_bg"], C=C)
    rect(slide, 0.95, 3.35, 3.2, 0.28, fill=C["primary"], C=C)
    text(slide, 0.9, 4.05, 2.2, 0.35, "行业平均", font_size=16, color="text_muted", C=C)
    text(slide, 0.9, 4.45, 2.6, 0.9, "18%", font_size=42, bold=True, color="text_dark", C=C)
    rect(slide, 0.95, 5.55, 4.8, 0.28, fill=C["card_bg"], C=C)
    rect(slide, 0.95, 5.55, 4.8, 0.28, fill=C["text_body"], C=C)
    rrect(slide, 7.0, 1.85, 5.35, 3.25, fill=C["card_bg"], line=C["divider"], C=C)
    text(slide, 7.45, 2.25, 3.4, 0.45, "关键判断", font_size=20, bold=True, color="primary", C=C)
    multiline(slide, 7.45, 2.95, 4.35, 1.0, ["问题不是市场没有增长", "而是现有增长方式难以复制"], font_size=22, color="text_dark", C=C, line_spacing=1.15)
    rect(slide, 7.45, 4.35, 1.3, 0.06, fill=C["risk"], C=C)
    text(slide, 7.45, 4.62, 4.2, 0.4, "风险：继续依赖单点项目", font_size=16, color="risk", C=C)

    # 3. Choice: show a connected growth system rather than a card wall.
    slide = base(prs, "三项战略选择", "从客户、交付和渠道同时改变增长结构", 3)
    rect(slide, 1.15, 3.45, 10.9, 0.08, fill=C["divider"], C=C)
    items = [
        (1.0, "01", "核心客户深耕", "一次性交付 → 持续经营", C["primary"]),
        (4.85, "02", "产品化交付", "项目经验 → 标准能力", "#2DD4BF"),
        (8.7, "03", "生态伙伴渠道", "直销能力 → 伙伴网络", "#5EEAD4"),
    ]
    for i, (left, num, title, body, color) in enumerate(items):
        rect(slide, left + 0.35, 3.15, 0.65, 0.65, fill=color, C=C)
        text(slide, left + 0.47, 3.28, 0.4, 0.3, num, font_size=14, bold=True, color="background", C=C, align="center")
        text(slide, left, 2.25, 3.0, 0.45, title, font_size=20, bold=True, color="text_dark", C=C)
        text(slide, left, 4.25, 3.2, 0.42, body, font_size=16, color="text_body", C=C)
        if i < 2:
            rect(slide, left + 3.05, 3.65, 0.55, 0.06, fill=color, C=C)
            text(slide, left + 3.03, 3.27, 0.6, 0.3, "→", font_size=22, bold=True, color=color, C=C, align="center")
    text(slide, 1.0, 5.55, 8.8, 0.45, "三项选择不是并列项目，而是一条从客户价值到规模复制的增长链路。", font_size=18, color="text_body", C=C)

    # 4. Target: turn the target into a three-lever system.
    slide = base(prs, "2027 目标：10.8 亿元", "目标来自三项增长方式的叠加，而不是单点冲刺", 4)
    rrect(slide, 0.9, 2.05, 3.7, 2.8, fill=C["card_bg"], line=C["divider"], C=C)
    rect(slide, 0.9, 2.05, 3.7, 0.08, fill=C["primary"], C=C)
    text(slide, 1.25, 2.55, 3.0, 1.0, "10.8 亿", font_size=44, bold=True, color="text_dark", C=C)
    text(slide, 1.28, 3.7, 2.8, 0.35, "2027 收入目标", font_size=17, color="text_muted", C=C)
    text(slide, 1.28, 4.25, 2.9, 0.35, "当前基线：8.6 亿", font_size=16, color="text_body", C=C)
    text(slide, 5.25, 2.05, 5.6, 0.45, "目标需要三种能力同时发生", font_size=21, bold=True, color="text_dark", C=C)
    rect(slide, 5.25, 2.85, 0.08, 2.6, fill=C["divider"], C=C)
    for i, (label, body) in enumerate([
        ("客户深耕", "提高存量客户的持续贡献"),
        ("产品化交付", "降低复制和交付成本"),
        ("伙伴渠道", "扩大有效触达范围"),
    ]):
        y = 2.78 + i * 0.92
        rect(slide, 5.8, y + 0.08, 0.34, 0.34, fill=C["primary"] if i == 0 else "#2DD4BF" if i == 1 else "#5EEAD4", C=C)
        text(slide, 6.45, y, 2.2, 0.35, label, font_size=18, bold=True, color="primary", C=C)
        text(slide, 8.75, y, 3.2, 0.4, body, font_size=17, color="text_body", C=C)
        rect(slide, 5.8, y + 0.62, 6.05, 0.03, fill=C["divider"], C=C)
    text(slide, 0.95, 5.65, 10.8, 0.42, "不是把一个数字做大，而是让三种增长能力形成闭环。", font_size=18, color="text_body", C=C)

    # 5. Execution: alternate the milestones to create rhythm and gates.
    slide = base(prs, "未来 12 个月执行节奏", "先验证，再产品化，再复制，最后复盘扩张", 5)
    rect(slide, 1.0, 3.65, 11.2, 0.08, fill=C["divider"], C=C)
    quarters = [("Q1", "定位与试点", "验证重点客户"), ("Q2", "产品化交付", "沉淀标准能力"), ("Q3", "规模复制", "扩大伙伴覆盖"), ("Q4", "复盘与扩张", "决定下一阶段")]
    for i, (q, title, body) in enumerate(quarters):
        x = 1.0 + i * 3.1
        color = C["primary"] if i < 3 else C["risk"]
        rect(slide, x, 3.35, 0.55, 0.62, fill=color, C=C)
        text(slide, x, 2.18 if i % 2 == 0 else 4.45, 1.0, 0.4, q, font_size=22, bold=True, color="primary" if i < 3 else "risk", C=C)
        text(slide, x, 2.45 if i % 2 == 0 else 4.98, 2.45, 0.4, title, font_size=18, bold=True, color="text_dark", C=C)
        text(slide, x, 2.92 if i % 2 == 0 else 5.42, 2.45, 0.42, body, font_size=16, color="text_body", C=C)
    text(slide, 1.0, 6.35, 8.5, 0.35, "每一季度都必须留下可复用的下一步，而不是只完成一次性动作。", font_size=16, color="text_muted", C=C)

    # 6. Close: decision request.
    slide = base(prs, "现在需要一个决定", "批准进入执行阶段", 6)
    text(slide, 0.8, 2.05, 8.7, 1.1, "把增长目标，\n变成组织动作。", font_size=36, bold=True, color="text_dark", C=C)
    multiline(slide, 0.85, 4.35, 5.4, 1.1, ["批准三项战略进入 Q1 试点", "指定负责人，启动月度复盘机制"], font_size=19, color="text_body", C=C, line_spacing=1.25)
    rect(slide, 9.55, 2.25, 2.7, 2.7, fill=C["primary"], C=C)
    text(slide, 9.9, 2.8, 2.0, 0.5, "NEXT", font_size=18, bold=True, color="background", C=C)
    text(slide, 9.9, 3.55, 2.0, 0.9, "Q1\n试点", font_size=30, bold=True, color="background", C=C)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUTPUT))
    print(f"Saved: {OUTPUT}")


if __name__ == "__main__":
    main()
