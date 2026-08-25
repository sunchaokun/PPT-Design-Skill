"""Evaluation 003: scientific research deck, Build Mode."""

from pathlib import Path

from pptx_designer import Presentation
from pptx_designer.tools.charts import bar_chart
from pptx_designer.tools.layout import page_number
from pptx_designer.tools.shapes import arrow, oval, rect, rrect
from pptx_designer.tools.text import multiline, text


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output" / "scientific_research_eval.pptx"

C = {
    "background": "#F7F8FA",
    "text_dark": "#1E293B",
    "text_body": "#334155",
    "text_muted": "#64748B",
    "divider": "#CBD5E1",
    "control": "#3B82F6",
    "treatment": "#C2414C",
    "primary": "#1E293B",
    "accent": "#3B82F6",
}


def base(prs, number, title=None, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["background"], C=C)
    if title:
        text(slide, 0.72, 0.55, 8.6, 0.55, title, font_size=27, bold=True, color="text_dark", C=C)
        text(slide, 0.74, 1.12, 9.4, 0.28, subtitle or "", font_size=12, color="text_muted", C=C)
        rect(slide, 0.72, 1.55, 11.9, 0.025, fill=C["divider"], C=C)
    page_number(slide, number, 6, C=C)
    return slide


def caption(slide, value):
    text(slide, 0.85, 6.75, 11.3, 0.3, value, font_size=10, color="text_muted", C=C)


def main():
    prs = Presentation()

    # 1. Formal scientific cover.
    slide = base(prs, 1)
    text(slide, 0.82, 1.25, 7.8, 0.4, "LAB MEETING / FIGURE-LED RESEARCH REPORT", font_size=11, bold=True, color="control", C=C)
    text(slide, 0.82, 2.1, 9.1, 1.35, "单细胞测序中的\n治疗反应信号", font_size=35, bold=True, color="text_dark", C=C)
    text(slide, 0.85, 4.15, 6.8, 0.6, "From sample preparation to response modeling", font_size=17, color="text_body", C=C)
    rect(slide, 0.85, 5.2, 1.4, 0.06, fill=C["treatment"], C=C)
    multiline(slide, 0.85, 5.65, 5.8, 0.75, ["Illustrative lab dataset", "Smith et al., 2024 / protocol note"], font_size=12, color="text_muted", C=C, line_spacing=1.1)
    # A restrained, content-bearing signal map gives the cover a visual anchor.
    rect(slide, 9.15, 1.65, 0.025, 4.25, fill=C["divider"], C=C)
    text(slide, 9.55, 1.65, 2.5, 0.3, "CELL-STATE SIGNAL MAP", font_size=10, bold=True, color="text_muted", C=C)
    for x, y, color in [
        (9.65, 2.35, C["control"]), (10.35, 2.05, C["control"]), (11.05, 2.55, C["control"]),
        (9.95, 3.15, C["control"]), (10.75, 3.45, C["treatment"]), (11.45, 3.05, C["treatment"]),
        (9.55, 4.05, C["control"]), (10.35, 4.35, C["treatment"]), (11.1, 4.05, C["treatment"]),
        (10.0, 5.1, C["treatment"]), (10.8, 5.35, C["treatment"]), (11.5, 4.85, C["treatment"]),
    ]:
        oval(slide, x, y, 0.34, 0.34, fill=color, C=C)
    text(slide, 9.55, 5.8, 2.7, 0.35, "control → treatment response", font_size=11, color="text_body", C=C)

    # 2. Question and hypothesis: a scientific relationship, not a hero page.
    slide = base(prs, 2, "01 / Research question", "What changes when the treatment condition is introduced?")
    text(slide, 0.85, 2.05, 4.7, 1.0, "Can single-cell signals\nseparate response states?", font_size=27, bold=True, color="text_dark", C=C)
    text(slide, 0.88, 3.55, 4.4, 0.7, "Hypothesis: treatment shifts the response index across cell populations.", font_size=16, color="text_body", C=C)
    rect(slide, 5.8, 2.55, 1.55, 0.9, fill="#DBEAFE", line=C["control"], C=C)
    text(slide, 5.98, 2.82, 1.2, 0.3, "CONTROL", font_size=12, bold=True, color="control", C=C, align="center")
    arrow(slide, 7.7, 2.87, 0.8, 0.25, fill=C["divider"], C=C)
    rect(slide, 8.8, 2.55, 1.55, 0.9, fill="#FEE2E2", line=C["treatment"], C=C)
    text(slide, 8.94, 2.82, 1.3, 0.3, "TREATMENT", font_size=12, bold=True, color="treatment", C=C, align="center")
    rect(slide, 6.0, 4.15, 4.2, 0.025, fill=C["divider"], C=C)
    text(slide, 6.0, 4.5, 4.3, 0.42, "same assay / different response state", font_size=15, color="text_muted", C=C)
    rect(slide, 0.85, 5.45, 11.45, 0.62, fill="#FFFFFF", line=C["divider"], C=C)
    text(slide, 1.1, 5.63, 1.2, 0.22, "SIGNAL", font_size=10, bold=True, color="control", C=C)
    text(slide, 2.25, 5.58, 2.3, 0.28, "single-cell expression", font_size=13, color="text_body", C=C)
    text(slide, 5.05, 5.63, 1.2, 0.22, "READOUT", font_size=10, bold=True, color="treatment", C=C)
    text(slide, 6.2, 5.58, 2.3, 0.28, "response index", font_size=13, color="text_body", C=C)
    text(slide, 9.05, 5.63, 1.3, 0.22, "TEST", font_size=10, bold=True, color="text_muted", C=C)
    text(slide, 10.25, 5.58, 1.7, 0.28, "replicate + validate", font_size=13, color="text_body", C=C)
    caption(slide, "Figure question: the comparison is illustrative and does not establish clinical efficacy.")

    # 3. Methods: three stages with explicit flow.
    slide = base(prs, 3, "02 / Methods", "Three stages connect the sample to an interpretable response index")
    stages = [
        (0.9, "A", "SAMPLE PREPARATION", "normalize input\nand preserve cell state", C["control"]),
        (4.65, "B", "SINGLE-CELL SEQUENCING", "capture expression\nacross populations", "#64748B"),
        (8.4, "C", "RESPONSE MODELING", "estimate response\nwith uncertainty", C["treatment"]),
    ]
    for i, (x, letter, title, body, color) in enumerate(stages):
        oval(slide, x, 2.25, 0.7, 0.7, fill=color, C=C)
        text(slide, x + 0.2, 2.47, 0.3, 0.25, letter, font_size=16, bold=True, color="#FFFFFF", C=C, align="center")
        text(slide, x, 3.35, 2.65, 0.48, title, font_size=15, bold=True, color="text_dark", C=C)
        multiline(slide, x, 4.05, 2.6, 0.8, body.split("\n"), font_size=15, color="text_body", C=C, line_spacing=1.1)
        if i < 2:
            arrow(slide, x + 1.0, 2.48, 1.8, 0.25, fill=C["divider"], C=C)
    caption(slide, "Methods overview. Stage labels are structural; protocol details remain in the source record.")

    # 4. Results: figure-led comparison with caption and source note.
    slide = base(prs, 4, "03 / Results", "Treatment response index is higher in the illustrative dataset")
    text(slide, 0.85, 1.95, 1.1, 0.35, "FIGURE 1", font_size=12, bold=True, color="text_dark", C=C)
    bar_chart(slide, 2.0, 2.25, [("Control", 0.42, "0.42"), ("Treatment", 0.78, "0.78")], max_width=6.4, bar_height=0.48, C={**C, "primary": C["control"], "accent": C["treatment"], "bg_tint": "#E2E8F0"})
    rrect(slide, 9.35, 2.1, 2.8, 2.25, fill="#FFFFFF", line=C["divider"], C=C)
    text(slide, 9.7, 2.45, 2.1, 0.35, "READOUT", font_size=12, bold=True, color="text_muted", C=C)
    text(slide, 9.7, 3.05, 1.9, 0.8, "0.78", font_size=30, bold=True, color="treatment", C=C)
    text(slide, 9.7, 3.82, 1.9, 0.35, "treatment index", font_size=12, color="text_body", C=C)
    # The lower band turns the comparison into an interpretable result, not a
    # pair of isolated bars. All values are explicitly framed as illustrative.
    rect(slide, 0.9, 4.95, 3.65, 0.05, fill=C["control"], C=C)
    text(slide, 0.9, 5.25, 3.3, 0.28, "DIFFERENCE", font_size=10, bold=True, color="control", C=C)
    text(slide, 0.9, 5.62, 3.3, 0.48, "0.36 index points", font_size=22, bold=True, color="text_dark", C=C)
    rect(slide, 4.85, 4.95, 3.65, 0.05, fill="#64748B", C=C)
    text(slide, 4.85, 5.25, 3.3, 0.28, "INTERPRETATION", font_size=10, bold=True, color="text_muted", C=C)
    text(slide, 4.85, 5.62, 3.3, 0.48, "direction matches hypothesis", font_size=16, bold=True, color="text_dark", C=C)
    rect(slide, 8.8, 4.95, 3.35, 0.05, fill=C["treatment"], C=C)
    text(slide, 8.8, 5.25, 3.0, 0.28, "BOUNDARY", font_size=10, bold=True, color="treatment", C=C)
    text(slide, 8.8, 5.62, 3.0, 0.48, "illustrative; replicate", font_size=16, bold=True, color="text_dark", C=C)
    caption(slide, "Figure 1. Illustrative control/treatment response index. Source: illustrative lab dataset; protocol adapted from Smith et al., 2024.")

    # 5. Discussion: evidence and limitation, no marketing CTA.
    slide = base(prs, 5, "04 / Discussion", "The result is promising, but interpretation remains bounded")
    # Three evidence bands create a complete discussion structure and use the
    # lower canvas for reasoning rather than decorative filler.
    panels = [
        (0.85, 3.75, 3.75, C["control"], "WHAT THE FIGURE SUPPORTS", ["Treatment shows a higher", "response index.", "The signal merits follow-up modeling."]),
        (4.8, 3.75, 3.75, "#64748B", "WHAT REMAINS UNCERTAIN", ["The dataset is illustrative,", "not a clinical conclusion.", "No causal effect is established."]),
        (8.75, 3.75, 3.75, C["treatment"], "NEXT VALIDATION", ["Replicate across donors", "and batches.", "Test robustness across cell states."]),
    ]
    for x, w, h, color, heading, lines in panels:
        rrect(slide, x, 2.05, w, 3.75, fill="#FFFFFF", line=C["divider"], C=C)
        rect(slide, x, 2.05, w, 0.08, fill=color, C=C)
        text(slide, x + 0.25, 2.42, w - 0.5, 0.35, heading, font_size=11, bold=True, color=color if color != "#64748B" else "text_muted", C=C)
        multiline(slide, x + 0.25, 3.15, w - 0.5, 1.7, lines, font_size=16, color="text_dark", C=C, line_spacing=1.25)
        evidence = {
            "WHAT THE FIGURE SUPPORTS": "Observed direction: 0.42 → 0.78",
            "WHAT REMAINS UNCERTAIN": "Evidence type: illustrative dataset",
            "NEXT VALIDATION": "Validation unit: donor × batch",
        }[heading]
        text(slide, x + 0.25, 4.65, w - 0.5, 0.3, evidence, font_size=11, color="text_muted", C=C)
        rect(slide, x + 0.25, 5.25, w - 0.5, 0.025, fill=C["divider"], C=C)
    text(slide, 1.1, 5.42, 11.0, 0.25, "Discussion rule: separate an observed signal, a plausible interpretation, and a claim that still requires validation.", font_size=11, color="text_muted", C=C, align="center")
    caption(slide, "Discussion guardrail: distinguish an observed signal from a validated treatment effect.")

    # 6. References and next experiment.
    slide = base(prs, 6, "05 / Next step", "Preserve traceability before expanding the claim")
    text(slide, 0.85, 2.0, 4.4, 0.4, "NEXT EXPERIMENT", font_size=13, bold=True, color="control", C=C)
    text(slide, 0.85, 2.55, 5.0, 0.9, "Replicate across donors\nand batches.", font_size=28, bold=True, color="text_dark", C=C)
    rect(slide, 7.0, 2.0, 0.05, 3.25, fill=C["divider"], C=C)
    multiline(slide, 7.45, 2.0, 4.7, 1.6, ["Source note", "Illustrative lab dataset", "Protocol adapted from Smith et al., 2024."], font_size=16, color="text_body", C=C, line_spacing=1.15)
    rule_y = 4.7
    rect(slide, 7.45, rule_y, 4.7, 0.025, fill=C["divider"], C=C)
    text(slide, 7.45, 5.0, 4.7, 0.4, "Figure 1 remains a signal, not a conclusion.", font_size=16, bold=True, color="text_dark", C=C)
    caption(slide, "End of report / questions and discussion")

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUTPUT))
    print(f"Saved: {OUTPUT}")


if __name__ == "__main__":
    main()
