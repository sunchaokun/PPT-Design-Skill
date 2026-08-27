"""Flagship case 005: AI Infrastructure Economics, Build Mode."""

from pathlib import Path

from pptx_designer import Presentation
from pptx_designer.tools.charts import bar_chart, comparison_bars
from pptx_designer.tools.layout import page_number
from pptx_designer.tools.shapes import arrow, diamond, oval, rect, rrect
from pptx_designer.tools.text import multiline, text


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output" / "ai_infrastructure_economics.pptx"

C = {
    "paper": "#F5F0E8",
    "paper2": "#E9E1D4",
    "ink": "#172033",
    "muted": "#6B7280",
    "rule": "#C9C0B2",
    "cobalt": "#245BDB",
    "red": "#C6473B",
    "green": "#6B8E5A",
    "blue_tint": "#DDE7FF",
    "red_tint": "#F3DAD3",
    "green_tint": "#DCE7D3",
    "white": "#FFFFFF",
}


def tx(slide, x, y, w, h, value, size=14, color="ink", bold=False, name="Aptos", align=None):
    kw = dict(font_size=size, color=color, bold=bold, font_name=name, C=C)
    if align is not None:
        kw["align"] = align
    return text(slide, x, y, w, h, value, **kw)


def multi(slide, x, y, w, h, lines, size=14, color="muted", spacing=1.1):
    return multiline(slide, x, y, w, h, lines, font_size=size, color=color, C=C, line_spacing=spacing)


def rule(slide, x, y, w, h=0.018, color="rule"):
    rect(slide, x, y, w, h, fill=C[color], C=C)


def tag(slide, x, y, w, value, color="cobalt", fill="paper"):
    rrect(slide, x, y, w, 0.27, fill=C[fill], line=C[color], C=C)
    tx(slide, x, y + 0.05, w, 0.15, value, 9, color, True, "Consolas", "center")


def box(slide, x, y, w, h, title, body=None, accent="cobalt", fill="white", title_size=14):
    rrect(slide, x, y, w, h, fill=C[fill], line=C[accent], C=C)
    rect(slide, x, y, 0.05, h, fill=C[accent], C=C)
    tx(slide, x + 0.18, y + 0.16, w - 0.32, 0.25, title, title_size, "ink", True)
    if body:
        multi(slide, x + 0.18, y + 0.55, w - 0.32, h - 0.62, body if isinstance(body, list) else body.split("\n"), 10, "muted", 1.0)


def base(prs, number, section, title, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["paper"], C=C)
    for x in [0.7, 3.1, 5.5, 7.9, 10.3, 12.7]:
        rule(slide, x, 0.45, 0.012, 6.2, "paper2")
    tag(slide, 0.72, 0.45, 2.0, section, "cobalt", "paper")
    tx(slide, 0.72, 0.9, 11.7, 0.5, title, 23, "ink", True)
    if subtitle:
        tx(slide, 0.74, 1.29, 11.6, 0.25, subtitle, 11, "muted")
    page_number(slide, number, 12, C=C)
    tx(slide, 11.4, 0.55, 1.1, 0.2, "INFRA / 02", 9, "muted", False, "Consolas", "right")
    return slide


def footer(slide, value):
    tx(slide, 0.73, 7.05, 11.6, 0.18, value, 8, "muted", False, "Consolas")


def main():
    prs = Presentation()

    # 1 Cover
    s = base(prs, 1, "INFRASTRUCTURE LEDGER", "", None)
    tx(s, 0.76, 1.0, 5.8, 1.0, "AI INFRASTRUCTURE\nECONOMICS", 32, "ink", True)
    tx(s, 0.78, 2.22, 5.8, 0.3, "Capital moves first. Constraints decide what ships.", 14, "muted")
    tx(s, 0.76, 2.55, 5.0, 0.45, "A report on capacity,\ncapital, and the physical stack.", 19, "red", True)
    rule(s, 0.76, 3.75, 1.45, 0.06, "red")
    multi(s, 0.78, 4.15, 4.8, 0.7, ["A source-led editorial data study", "for strategy and platform operators."], 13, "muted")
    # Infrastructure silhouette
    tx(s, 7.15, 1.9, 4.5, 0.25, "THE VALUE CHAIN IS PHYSICAL", 10, "muted", True, "Consolas")
    layers = [(2.35, 4.85, "APPLICATIONS", "value", "green"), (3.0, 4.25, "PLATFORM", "distribution", "cobalt"), (3.65, 3.65, "COMPUTE", "scarcity", "red"), (4.3, 3.05, "POWER + PLACE", "constraint", "ink")]
    for y, w, label, sub, color in layers:
        x = 7.25 + (4.85 - w) / 2
        rrect(s, x, y, w, 0.48, fill=C["paper2"], line=C[color], C=C)
        tx(s, x + 0.12, y + 0.13, w - 0.24, 0.16, label, 10, color, True, "Consolas", "center")
        tx(s, x + w + 0.15, y + 0.14, 1.3, 0.16, sub, 10, "muted", False, "Consolas")
    tx(s, 0.78, 6.7, 5.4, 0.2, "PUBLIC DISCLOSURES / ILLUSTRATIVE SYSTEMS ANALYSIS", 9, "muted", False, "Consolas")
    footer(s, "SOURCES: ALPHABET IR · MICROSOFT IR · SEE SOURCE NOTES IN PAGE PLAN")

    # 2 Stack
    s = base(prs, 2, "01 / THE STACK", "AI is not one market. It is a stack of dependencies.", "The margin pool can move upward; the bottleneck often stays physical.")
    stack = [
        (1.0, 1.95, 11.1, "07 / APPLICATIONS", "workflow · distribution · willingness to pay", "green"),
        (1.45, 2.75, 10.2, "06 / MODELS", "training · inference · evaluation", "cobalt"),
        (1.9, 3.55, 9.3, "05 / SOFTWARE", "orchestration · serving · observability", "cobalt"),
        (2.35, 4.35, 8.4, "04 / SYSTEMS", "accelerator · memory · network fabric", "red"),
        (2.8, 5.15, 7.5, "03 / DATA CENTERS", "land · cooling · power delivery", "red"),
    ]
    for x, y, w, label, body, accent in stack:
        rrect(s, x, y, w, 0.6, fill=C["white"], line=C[accent], C=C)
        tx(s, x + 0.2, y + 0.17, 2.0, 0.18, label, 10, accent, True, "Consolas")
        tx(s, x + 2.25, y + 0.17, w - 2.45, 0.18, body, 12, "ink")
    tx(s, 10.0, 6.03, 2.2, 0.25, "dependency ↑", 11, "red", True, "Consolas", "center")
    arrow(s, 11.0, 2.0, 0.25, 3.75, fill=C["red"], C=C)
    footer(s, "INTERPRETATION / THE ECONOMICS OF AI ARE DISTRIBUTED ACROSS THE STACK")

    # 3 Reported capex field
    s = base(prs, 3, "02 / CAPITAL", "Capital is accelerating — but the periods are not interchangeable.", "Reported figures are shown with company and fiscal-period labels, not as a false league table.")
    # Native editable bars keep the evidence field separate from the reading
    # panel and make the source labels legible at presentation scale.
    reported = [("Alphabet FY2025", 91.4, "$91.4B", "cobalt"), ("Microsoft FY25 Q4", 24.2, "$24.2B", "red"), ("Alphabet FY2026 outlook", 180, "$175–185B", "muted")]
    bar_left, bar_top, bar_width, bar_h, max_value = 1.75, 2.05, 6.25, 0.52, 180
    for i, (label, value, display, accent) in enumerate(reported):
        y = bar_top + i * 0.72
        tx(s, 0.55, y + 0.16, 1.05, 0.16, label, 9, "ink", False, "Aptos", "right")
        rect(s, bar_left, y, bar_width, bar_h, fill=C["paper2"], C=C)
        rect(s, bar_left, y, bar_width * value / max_value, bar_h, fill=C[accent], C=C)
        tx(s, bar_left + bar_width * value / max_value - 0.1, y + 0.16, 0.85, 0.16, display, 9, "ink", True, "Consolas", "right")
    box(s, 8.7, 2.0, 3.45, 2.65, "READ THIS AS", ["directional evidence", "not a single-period ranking", "different disclosure periods", "different business mix"], "red", "paper2")
    rule(s, 0.95, 5.2, 11.3, 0.025, "rule")
    tx(s, 0.98, 5.48, 2.2, 0.2, "WHAT IS REPORTED", 10, "cobalt", True, "Consolas")
    tx(s, 4.2, 5.48, 2.2, 0.2, "WHAT IT SUGGESTS", 10, "red", True, "Consolas")
    tx(s, 7.45, 5.48, 2.2, 0.2, "WHAT IT DOES NOT PROVE", 10, "green", True, "Consolas")
    multi(s, 0.98, 5.85, 2.8, 0.55, ["Large technical-infrastructure", "commitments are visible."], 12, "ink")
    multi(s, 4.2, 5.85, 2.8, 0.55, ["Capacity is becoming a", "strategic balance-sheet item."], 12, "ink")
    multi(s, 7.45, 5.85, 3.8, 0.55, ["CapEx alone does not tell us", "who captures the margin."], 12, "ink")
    footer(s, "SOURCES: ALPHABET 2025 Q4 EARNINGS CALL · MICROSOFT FY25 Q4 EARNINGS · FIGURES IN USD")

    # 4 Composition
    s = base(prs, 4, "03 / COMPOSITION", "Where the money lands changes the depreciation story.", "One disclosed split makes the physical stack visible.")
    tx(s, 1.0, 2.1, 2.2, 0.35, "ALPHABET FY2025", 11, "muted", True, "Consolas")
    tx(s, 1.0, 2.62, 3.2, 0.8, "$91.4B", 36, "ink", True)
    tx(s, 1.0, 3.55, 3.4, 0.35, "technical infrastructure CapEx", 13, "muted")
    # 60/40 composition band
    tx(s, 5.0, 2.1, 5.4, 0.25, "APPROXIMATE 2025 COMPOSITION", 10, "muted", True, "Consolas")
    rect(s, 5.0, 2.8, 6.6, 0.72, fill=C["cobalt"], C=C)
    rect(s, 8.96, 2.8, 2.64, 0.72, fill=C["red"], C=C)
    tx(s, 5.25, 3.03, 3.4, 0.2, "60%  SERVERS", 13, "white", True, "Consolas")
    tx(s, 9.2, 3.03, 2.1, 0.2, "40%  DC + NETWORK", 11, "white", True, "Consolas")
    box(s, 5.0, 4.1, 3.0, 1.45, "SHORTER CYCLE", ["servers", "faster refresh", "utilization pressure"], "cobalt", "blue_tint")
    box(s, 8.6, 4.1, 3.0, 1.45, "LONGER DURATION", ["data centers", "networking", "depreciation horizon"], "red", "red_tint")
    tx(s, 1.0, 5.15, 3.2, 0.3, "The stack has different clocks.", 17, "red", True)
    multi(s, 1.0, 5.65, 3.4, 0.65, ["A server cycle can move faster", "than a building or grid connection."], 12, "muted")
    footer(s, "SOURCE: ALPHABET 2025 Q4 EARNINGS CALL / APPROXIMATELY 60% SERVERS, 40% DATA CENTERS + NETWORKING")

    # 5 Constraints map
    s = base(prs, 5, "04 / BOTTLENECK", "The constraint is not “compute.” It is the coordination of four physical systems.", "Land, energy, networking, and servers arrive on different timelines.")
    constraints = [(0.95, 2.25, "LAND", "permitted + buildable", "red"), (3.95, 2.25, "ENERGY", "predictable supply", "red"), (6.95, 2.25, "NETWORK", "fabric + interconnect", "cobalt"), (9.95, 2.25, "SERVERS", "GPU + components", "cobalt")]
    for x, y, title, body, accent in constraints:
        box(s, x, y, 2.35, 1.55, title, [body, "lead time ≠ demand"], accent, "white", 15)
    for x in [3.45, 6.45, 9.45]:
        arrow(s, x, 2.9, 0.28, 0.22, fill=C["rule"], C=C)
    rule(s, 1.1, 4.65, 10.8, 0.03, "red")
    tx(s, 1.15, 5.0, 10.4, 0.35, "CAPACITY ONLY SHIPS WHEN THE SLOWEST LAYER ARRIVES.", 18, "red", True, "Consolas", "center")
    multi(s, 2.0, 5.65, 9.2, 0.55, ["This is why supply-chain risk belongs inside the AI strategy — not in a footnote after the model roadmap."], 13, "ink", 1.1)
    footer(s, "SOURCE: MICROSOFT FY2025 ANNUAL REPORT / DATA CENTER LAND, ENERGY, NETWORKING, SERVERS")

    # 6 Value flywheel
    s = base(prs, 6, "05 / VALUE", "Capacity becomes economics only through utilization.", "The flywheel is simple to draw — and difficult to keep turning.")
    center = (6.55, 3.85)
    oval(s, 5.35, 3.0, 2.4, 1.2, fill=C["paper2"], line=C["red"], C=C)
    tx(s, 5.65, 3.42, 1.8, 0.25, "OPERATING\nLEVERAGE", 14, "red", True, "Consolas", "center")
    fly = [(1.0, 2.25, "CAPACITY", "hardware + power", "cobalt"), (9.8, 2.25, "REVENUE", "serving demand", "green"), (9.8, 4.75, "CASH", "depreciation + margin", "red"), (1.0, 4.75, "UTILIZATION", "load + scheduling", "cobalt")]
    for x, y, title, body, accent in fly:
        box(s, x, y, 2.45, 1.0, title, [body], accent, "white")
    arrow(s, 3.55, 2.65, 1.55, 0.22, fill=C["cobalt"], C=C)
    arrow(s, 10.0, 3.65, 0.22, 0.7, fill=C["green"], C=C)
    arrow(s, 8.0, 5.1, 1.55, 0.22, fill=C["red"], C=C)
    arrow(s, 2.95, 4.0, 0.22, 0.7, fill=C["cobalt"], C=C)
    tx(s, 4.7, 6.2, 3.7, 0.25, "NO UTILIZATION → NO LEVERAGE", 11, "red", True, "Consolas", "center")
    footer(s, "INTERPRETATION / THE UNIT ECONOMICS OF AI ARE A CAPACITY-MANAGEMENT PROBLEM")

    # 7 Training vs inference curves
    s = base(prs, 7, "06 / ECONOMICS", "Training is episodic. Inference is continuous.", "Illustrative curves show why utilization and serving efficiency can matter more than a single benchmark.")
    # chart axes
    rule(s, 1.2, 5.75, 7.3, 0.025, "ink")
    rule(s, 1.2, 2.25, 0.025, 3.5, "ink")
    tx(s, 1.0, 1.85, 1.0, 0.2, "COST", 10, "muted", True, "Consolas")
    tx(s, 7.9, 5.98, 1.0, 0.2, "TIME", 10, "muted", True, "Consolas")
    # illustrative line segments using bars and dots
    rule(s, 1.25, 3.0, 1.4, 0.04, "red"); rule(s, 2.65, 3.55, 1.4, 0.04, "red"); rule(s, 4.05, 4.0, 1.4, 0.04, "red"); rule(s, 5.45, 4.45, 1.4, 0.04, "red"); rule(s, 6.85, 4.85, 1.4, 0.04, "red")
    rule(s, 1.25, 5.1, 1.4, 0.04, "cobalt"); rule(s, 2.65, 4.95, 1.4, 0.04, "cobalt"); rule(s, 4.05, 4.8, 1.4, 0.04, "cobalt"); rule(s, 5.45, 4.65, 1.4, 0.04, "cobalt"); rule(s, 6.85, 4.5, 1.4, 0.04, "cobalt")
    tx(s, 1.4, 2.7, 2.0, 0.2, "TRAINING / burst", 11, "red", True, "Consolas")
    tx(s, 1.4, 5.25, 2.3, 0.2, "INFERENCE / stream", 11, "cobalt", True, "Consolas")
    box(s, 9.35, 2.35, 2.75, 1.35, "OPERATOR QUESTION", ["Can the load be", "scheduled, shifted,", "or pooled?"], "green", "green_tint")
    box(s, 9.35, 4.15, 2.75, 1.35, "ILLUSTRATIVE ONLY", ["Shape of curve", "is conceptual", "not a benchmark"], "red", "red_tint")
    footer(s, "ANALYTICAL MODEL / ILLUSTRATIVE CURVES — NOT A PERFORMANCE CLAIM")

    # 8 Flow of value
    s = base(prs, 8, "07 / MARGIN", "The margin pool is a flow, not a single layer.", "Compute creates the possibility; platforms and applications decide where value compounds.")
    stages = [(0.95, 2.5, 2.3, "POWER + PLACE", "fixed assets", "red"), (3.65, 2.5, 2.3, "COMPUTE", "scarce capacity", "red"), (6.35, 2.5, 2.3, "PLATFORM", "orchestration", "cobalt"), (9.05, 2.5, 2.3, "APPLICATION", "workflow value", "green")]
    for x, y, w, title, body, accent in stages:
        box(s, x, y, w, 1.35, title, [body], accent, "white", 13)
    for x, color in [(3.28, "red"), (5.98, "cobalt"), (8.68, "green")]:
        arrow(s, x, 3.05, 0.32, 0.22, fill=C[color], C=C)
    rule(s, 1.0, 4.9, 10.4, 0.03, "rule")
    tx(s, 1.0, 5.25, 2.6, 0.25, "CAPITAL INTENSITY", 10, "red", True, "Consolas")
    tx(s, 4.2, 5.25, 2.6, 0.25, "OPERATING LEVERAGE", 10, "cobalt", True, "Consolas")
    tx(s, 7.45, 5.25, 2.6, 0.25, "CUSTOMER LOCK-IN", 10, "green", True, "Consolas")
    multi(s, 1.0, 5.7, 2.6, 0.55, ["low flexibility", "high fixed cost"], 12, "muted")
    multi(s, 4.2, 5.7, 2.6, 0.55, ["shared capacity", "better scheduling"], 12, "muted")
    multi(s, 7.45, 5.7, 2.6, 0.55, ["workflow fit", "switching friction"], 12, "muted")
    footer(s, "INTERPRETATION / VALUE MOVES UP THE STACK, BUT DEPENDENCIES MOVE WITH IT")

    # 9 Risk matrix
    s = base(prs, 9, "08 / RISK", "The dangerous asset is capacity you cannot use.", "A simple matrix makes the hidden risk visible: high lead time plus low utilization.")
    rule(s, 1.5, 5.75, 7.2, 0.025, "ink"); rule(s, 1.5, 2.2, 0.025, 3.55, "ink")
    tx(s, 3.4, 6.05, 3.8, 0.2, "UTILIZATION  →", 10, "cobalt", True, "Consolas", "center")
    tx(s, 0.7, 3.4, 0.25, 1.5, "L\nE\nA\nD\nT\nI\nM\nE", 9, "red", True, "Consolas", "center")
    box(s, 2.0, 4.3, 2.65, 1.0, "UNDERUSE", ["capacity exists", "demand is soft"], "green", "green_tint")
    box(s, 5.0, 4.3, 2.65, 1.0, "HEALTHY", ["matched load", "optionality"], "cobalt", "blue_tint")
    box(s, 2.0, 2.75, 2.65, 1.0, "STRANDED", ["fixed cost", "no flexibility"], "red", "red_tint")
    box(s, 5.0, 2.75, 2.65, 1.0, "BOTTLENECK", ["demand exists", "supply lags"], "red", "red_tint")
    rrect(s, 9.1, 2.45, 3.0, 3.0, fill=C["paper2"], line=C["red"], C=C)
    tx(s, 9.42, 2.8, 2.25, 0.2, "WATCH ITEM", 10, "red", True, "Consolas")
    tx(s, 9.42, 3.35, 2.25, 0.75, "Lead time\ncan outlive\ndemand.", 23, "ink", True)
    multi(s, 9.42, 4.55, 2.2, 0.5, ["Treat capacity as an option", "until demand is proven."], 11, "muted")
    footer(s, "ILLUSTRATIVE DECISION MATRIX / RISK IS A JOINT FUNCTION OF LEAD TIME AND UTILIZATION")

    # 10 Buy/rent/build
    s = base(prs, 10, "09 / CHOICE", "There are three ways to buy capacity. None is free.", "The right answer depends on demand certainty, capital appetite, and control requirements.")
    headers = [(1.0, "BUY", "control"), (4.25, "RENT", "optionality"), (7.5, "BUILD", "differentiation")]
    for x, title, sub in headers:
        tx(s, x, 2.0, 2.4, 0.35, title, 22, "ink", True)
        tx(s, x, 2.45, 2.4, 0.18, sub, 10, "muted", True, "Consolas")
        rule(s, x, 2.85, 2.35, 0.035, "cobalt" if title == "RENT" else "red" if title == "BUILD" else "ink")
    rows = [("CAPITAL", "high", "variable", "very high"), ("SPEED", "medium", "fast", "slow"), ("CONTROL", "high", "shared", "highest"), ("RISK", "demand", "vendor", "execution")]
    for i, (label, a, b, c) in enumerate(rows):
        y = 3.35 + i * 0.6
        tx(s, 1.0, y, 1.0, 0.18, label, 9, "muted", True, "Consolas")
        tx(s, 2.1, y, 1.1, 0.2, a, 12, "ink", False, "Consolas")
        tx(s, 4.25, y, 1.1, 0.2, b, 12, "cobalt", True, "Consolas")
        tx(s, 7.5, y, 1.1, 0.2, c, 12, "red", True, "Consolas")
        rule(s, 1.0, y + 0.33, 8.85, 0.012, "rule")
    rrect(s, 10.15, 2.0, 2.0, 3.95, fill=C["paper2"], line=C["green"], C=C)
    tx(s, 10.45, 2.4, 1.4, 0.2, "OPERATOR TEST", 9, "green", True, "Consolas", "center")
    multi(s, 10.45, 3.0, 1.4, 1.6, ["How certain", "is demand?", "How unique", "is the workload?"], 14, "ink", 1.35)
    footer(s, "DECISION MODEL / CAPACITY STRATEGY IS A PORTFOLIO, NOT A SINGLE PROCUREMENT MODE")

    # 11 Scorecard
    s = base(prs, 11, "10 / CONTROL", "Run infrastructure like an operating system, not a warehouse.", "Five measures reveal whether the asset is earning its keep.")
    score = [("CAPACITY", "what exists", "cobalt"), ("UTILIZATION", "what is used", "green"), ("LATENCY", "what users feel", "cobalt"), ("DEPRECIATION", "what time costs", "red"), ("POWER", "what scale requires", "red")]
    for i, (label, body, accent) in enumerate(score):
        x = 0.95 + i * 2.4
        box(s, x, 2.25, 2.05, 1.15, label, [body], accent, "white", 11)
        tx(s, x + 0.2, 3.75, 1.65, 0.2, f"0{i + 1}", 10, accent, True, "Consolas")
        rule(s, x + 0.2, 4.2, 1.65, 0.035, accent)
    tx(s, 1.0, 5.1, 11.0, 0.32, "THE SCORECARD TURNS CAPEX INTO A FEEDBACK SYSTEM.", 18, "ink", True, "Consolas", "center")
    multi(s, 2.0, 5.7, 9.2, 0.55, ["If one measure is missing, the operator is managing a blind spot — not a platform."], 13, "muted", 1.1)
    footer(s, "OPERATOR CONTROL / MEASURE CAPACITY, UTILIZATION, USER EXPERIENCE, TIME, AND POWER TOGETHER")

    # 12 Close
    s = base(prs, 12, "11 / CLOSE", "The winning infrastructure is not the biggest. It is the most legible.", "Capital creates the option. Constraints define the clock. Control decides the return.")
    tx(s, 0.9, 2.0, 5.1, 1.4, "Capital creates\nthe option.", 31, "ink", True)
    tx(s, 0.9, 3.75, 5.1, 1.1, "Constraints define\nthe clock.", 25, "red", True)
    tx(s, 0.9, 5.15, 5.1, 0.8, "Control decides\nthe return.", 22, "green", True)
    rrect(s, 7.0, 2.0, 4.8, 3.8, fill=C["ink"], line=C["ink"], C=C)
    tx(s, 7.35, 2.45, 4.0, 0.2, "THE OPERATOR’S QUESTION", 10, "paper", True, "Consolas")
    multi(s, 7.35, 3.05, 3.9, 1.8, ["Where is the bottleneck?", "What can we flex?", "What must we own?"], 20, "paper", 1.35)
    rule(s, 7.35, 5.2, 3.8, 0.04, "green")
    tx(s, 7.35, 5.45, 3.8, 0.2, "MAKE THE STACK LEGIBLE.", 11, "green", True, "Consolas")
    footer(s, "AI INFRASTRUCTURE ECONOMICS / END OF REPORT")

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUTPUT))
    print(f"Saved: {OUTPUT}")


if __name__ == "__main__":
    main()
