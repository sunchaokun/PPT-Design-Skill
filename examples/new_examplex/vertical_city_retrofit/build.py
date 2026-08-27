"""Flagship case 007: Vertical City Retrofit — two-page style prototype."""
from pathlib import Path

from pptx_designer import Presentation
from pptx_designer.tools.images import cover_image
from pptx_designer.tools.layout import page_number
from pptx_designer.tools.shapes import arrow, oval, rect
from pptx_designer.tools.text import multiline, text

ROOT = Path(__file__).resolve().parent
ASSET = ROOT / "assets"
OUT = ROOT / "output" / "vertical_city_retrofit_style_sample.pptx"
C = {
    "paper": "#F1EEE8", "ink": "#20272B", "muted": "#6F7775",
    "rule": "#C9C2B8", "brick": "#B85C42", "green": "#597461",
    "blue": "#1F6F8B", "sand": "#E5DED2", "white": "#FFFFFF",
}


def tx(s, x, y, w, h, v, size=12, color="ink", bold=False, name="Aptos", align=None):
    kw = dict(font_size=size, color=color, bold=bold, font_name=name, C=C)
    if align is not None:
        kw["align"] = align
    return text(s, x, y, w, h, v, **kw)


def ml(s, x, y, w, h, lines, size=11, color="muted", spacing=1.12):
    return multiline(s, x, y, w, h, lines, font_size=size, color=color, C=C, line_spacing=spacing)


def line(s, x, y, w, h=0.018, color="rule"):
    rect(s, x, y, w, h, fill=C[color], C=C)


def base(prs, n, section, title, sub):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 7.5, fill=C["paper"], C=C)
    for x in [0.68, 3.18, 5.68, 8.18, 10.68, 12.68]:
        line(s, x, 0.45, 0.012, 6.25, "sand")
    line(s, 0.72, 0.44, 1.95, 0.035, "blue")
    tx(s, 0.72, 0.58, 2.2, 0.16, section, 8.5, "blue", True, "Consolas")
    tx(s, 0.74, 0.92, 11.8, 0.45, title, 24, "ink", True)
    tx(s, 0.76, 1.34, 11.6, 0.18, sub, 10, "muted")
    page_number(s, n, 14, C=C)
    tx(s, 11.35, 0.6, 1.2, 0.14, "CITY / 01", 8, "muted", False, "Consolas", "right")
    return s


def footer(s, v):
    tx(s, 0.76, 7.05, 11.5, 0.14, v, 7.5, "muted", False, "Consolas")


def main():
    prs = Presentation()
    OUT.parent.mkdir(parents=True, exist_ok=True)

    # 1 cover — editorial architecture image with quiet left field.
    s = prs.slides.add_slide(prs.slide_layouts[6])
    cover_image(s, 0, 0, 13.333, 7.5, str(ASSET / "vertical_city_cover.png"))
    line(s, 0.82, 0.82, 2.95, 0.045, "blue")
    tx(s, 0.84, 1.10, 4.8, 0.18, "URBAN SYSTEMS / CASE 07", 9, "blue", True, "Consolas")
    tx(s, 0.84, 1.62, 4.9, 1.15, "VERTICAL CITY\nRETROFIT", 29, "ink", True)
    tx(s, 0.86, 3.22, 4.65, 0.52, "The next city is not built from scratch.\nIt is upgraded in place.", 14, "ink", True)
    ml(s, 0.88, 4.25, 4.5, 0.88, ["A strategy for aging high-rise housing", "envelope · energy · shared space · care"], 10.5, "muted", 1.25)
    tx(s, 0.88, 6.42, 4.5, 0.14, "CONCEPTUAL HERO VISUAL / NOT A REAL PROJECT", 7.3, "blue", True, "Consolas")
    tx(s, 0.88, 6.82, 4.5, 0.12, "VERTICAL CITY RETROFIT · 2026", 7.2, "muted", False, "Consolas")

    # 2 inner page — architectural diagnosis, no card grid.
    s = base(prs, 2, "01 / THE VISIBLE PROBLEM", "An aging tower is a stack of hidden pressures", "The façade is only the visible layer; the retrofit brief begins where energy, access, and care intersect.")

    # Editable sectional tower: floors, core, and intervention markers.
    tx(s, 0.82, 1.82, 4.1, 0.16, "SECTIONAL DIAGNOSIS / EXISTING CONDITION", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.08, 4.45, 0.025, "blue")
    rect(s, 1.45, 2.42, 2.25, 3.75, fill=C["sand"], line=C["ink"], C=C)
    rect(s, 2.35, 2.68, 0.38, 3.18, fill=C["ink"], C=C)
    for i in range(8):
        y = 2.60 + i * 0.43
        line(s, 1.55, y, 2.05, 0.014, "rule")
        line(s, 1.68, y + 0.11, 0.42, 0.06, "brick")
        line(s, 2.95, y + 0.11, 0.50, 0.06, "green")
    rect(s, 1.28, 6.18, 2.6, 0.18, fill=C["ink"], C=C)
    # Datum lines and callouts.
    for y, label, detail, col in [
        (2.76, "HEAT", "envelope loss", "brick"),
        (3.90, "ACCESS", "lift + threshold", "blue"),
        (5.06, "CARE", "isolated shared space", "green"),
        (6.00, "ENERGY", "plant at end-of-life", "brick"),
    ]:
        line(s, 3.72, y, 0.72, 0.014, col)
        arrow(s, 4.42, y - 0.03, 0.30, 0, C[col], C=C)
        tx(s, 4.92, y - 0.11, 0.9, 0.14, label, 8.2, col, True, "Consolas")
        tx(s, 5.88, y - 0.11, 1.8, 0.14, detail, 9.2, "ink", False)

    # Right-side editorial argument and a pressure index, using rules rather than boxes.
    tx(s, 7.86, 1.84, 4.2, 0.18, "THE RETROFIT BRIEF", 9, "brick", True, "Consolas")
    line(s, 7.86, 2.12, 4.32, 0.03, "brick")
    tx(s, 7.86, 2.48, 4.0, 0.62, "Upgrade the\nwhole section.", 22, "ink", True)
    ml(s, 7.88, 3.36, 3.9, 0.78, [
        "A façade-only fix can lower heat loss,",
        "but it cannot repair the social and",
        "operational systems around the home.",
    ], 11.5, "muted", 1.18)
    tx(s, 7.86, 4.56, 4.0, 0.16, "PRESSURE INDEX / ILLUSTRATIVE", 8.5, "blue", True, "Consolas")
    line(s, 7.86, 4.84, 4.05, 0.018, "rule")
    for i, (label, value, col) in enumerate([("THERMAL", 0.86, "brick"), ("ACCESS", 0.62, "blue"), ("SOCIAL", 0.74, "green")]):
        y = 5.12 + i * 0.36
        tx(s, 7.88, y, 0.85, 0.14, label, 7.8, "muted", True, "Consolas")
        rect(s, 8.88, y + 0.03, 2.62, 0.09, fill=C["sand"], C=C)
        rect(s, 8.88, y + 0.03, 2.62 * value, 0.09, fill=C[col], C=C)
        tx(s, 11.68, y - 0.01, 0.35, 0.15, f"{int(value*100)}", 8.5, col, True, "Consolas", "right")
    line(s, 7.86, 6.30, 4.05, 0.025, "blue")
    tx(s, 7.86, 6.55, 4.4, 0.22, "The unit of change is the building section.", 11.5, "blue", True)
    footer(s, "Scenario framing is illustrative; visual language separates diagnosis from future intervention.")

    # 3 inner page — exploded axonometric / operating section.
    s = base(prs, 3, "02 / SYSTEMS MAP", "One tower, five systems", "The retrofit is not a façade package. It is a coordinated intervention across the whole vertical section.")
    tx(s, 0.82, 1.82, 4.0, 0.16, "EXPLODED AXONOMETRIC / OPERATING SECTION", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.08, 11.45, 0.025, "blue")

    # A larger architectural object with offset slabs makes the page read as a building drawing.
    rect(s, 3.18, 2.52, 3.45, 3.16, fill=C["sand"], line=C["ink"], C=C)
    rect(s, 4.66, 2.72, 0.52, 2.74, fill=C["ink"], C=C)
    for y in [2.86, 3.34, 3.82, 4.30, 4.78, 5.26]:
        line(s, 3.32, y, 3.17, 0.014, "rule")
    # Offset intervention slabs / technical layers.
    for x, y, w, col in [(2.86, 2.42, 4.08, "brick"), (3.04, 3.22, 3.90, "blue"), (3.22, 4.02, 3.72, "green"), (3.40, 4.82, 3.54, "brick")]:
        line(s, x, y, w, 0.055, col)
        line(s, x + .12, y + .11, w - .24, 0.012, col)
    rect(s, 2.92, 5.84, 4.02, 0.16, fill=C["ink"], C=C)
    tx(s, 4.00, 6.16, 2.2, 0.16, "ONE SECTION / FIVE OPERATING LAYERS", 8.2, "blue", True, "Consolas", "center")

    callouts = [
        (0.96, 2.54, "01", "ENVELOPE", "thermal skin", "brick", 3.18),
        (0.96, 3.34, "02", "ENERGY", "plant + distribution", "blue", 3.04),
        (0.96, 4.14, "03", "MOBILITY", "lift + threshold", "green", 3.22),
        (7.90, 4.46, "04", "SHARED SPACE", "lobby + terrace", "brick", 6.94),
        (7.90, 5.18, "05", "CARE", "neighbourhood support", "green", 6.94),
    ]
    for x, y, n, label, detail, col, anchor in callouts:
        tx(s, x, y, .34, .16, n, 8, col, True, "Consolas")
        tx(s, x + .48, y, 1.45, .16, label, 8.4, col, True, "Consolas")
        tx(s, x + .48, y + .25, 1.72, .16, detail, 9.0, "muted")
        if x < 2:
            line(s, x + 1.98, y + .08, anchor - (x + 1.98), 0.014, col)
        else:
            line(s, anchor, y + .08, x - anchor - .18, 0.014, col)

    tx(s, 8.02, 2.28, 3.8, 0.18, "FROM BUILDING TO OPERATING MODEL", 9, "brick", True, "Consolas")
    line(s, 8.02, 2.56, 3.8, 0.03, "brick")
    tx(s, 8.02, 2.90, 3.75, 0.58, "Value appears\nin the connections.", 21, "ink", True)
    ml(s, 8.04, 3.74, 3.55, 0.52, ["The same section can reduce loss,", "restore access, and create care capacity."], 10.8, "muted", 1.15)
    line(s, 8.02, 5.62, 3.8, 0.025, "blue")
    tx(s, 8.02, 5.92, 3.7, 0.22, "The tower behaves like a city when its layers work together.", 11.3, "blue", True)
    footer(s, "Conceptual axonometric; system labels define the retrofit vocabulary and are not measured project outcomes.")

    # 4 inner page — horizontal delivery path, deliberately unlike both prior pages.
    s = base(prs, 4, "03 / DELIVERY PATH", "The intervention moves in five deliberate steps", "A retrofit succeeds when each move unlocks the next one without interrupting the life of the building.")
    tx(s, 0.82, 1.86, 3.6, 0.16, "OCCUPIED-BUILDING SEQUENCE", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.12, 11.55, 0.025, "blue")
    stages = [
        (0.92, "01", "MEASURE", "baseline the\nwhole section", "thermal + social", "brick"),
        (3.22, "02", "WRAP", "reduce loss\nfirst", "envelope + plant", "blue"),
        (5.52, "03", "OPEN", "repair the\nthreshold", "lift + lobby", "green"),
        (7.82, "04", "SHARE", "activate\ncommon life", "terrace + care", "brick"),
        (10.12, "05", "ADAPT", "operate\nand learn", "data + feedback", "blue"),
    ]
    for i, (x, n, verb, body, detail, col) in enumerate(stages):
        y = 3.06 - i * 0.16
        rect(s, x, y, 1.72, 1.46, fill=C["sand"], line=C[col], C=C)
        line(s, x, y, 1.72, 0.08, col)
        tx(s, x + 0.16, y + 0.22, 0.34, 0.16, n, 8, col, True, "Consolas")
        tx(s, x + 0.16, y + 0.52, 1.3, 0.18, verb, 10.5, col, True, "Aptos Display")
        tx(s, x + 0.16, y + 0.82, 1.35, 0.38, body, 11, "ink", True)
        tx(s, x + 0.16, y + 1.28, 1.35, 0.14, detail, 7.8, "muted", False, "Consolas")
        if i < 4:
            arrow(s, x + 1.82, y + 0.64, 0.38, 0, C["blue"], C=C)
    line(s, 0.92, 5.30, 10.92, 0.035, "ink")
    tx(s, 0.92, 5.62, 2.2, 0.16, "THE DELIVERY TEST", 8.5, "brick", True, "Consolas")
    tx(s, 0.92, 5.98, 10.8, 0.28, "Can the building become more comfortable, more connected, and more useful while people remain inside it?", 15, "ink", True)
    tx(s, 0.92, 6.48, 10.7, 0.16, "The sequence is the value proposition: reduce loss → restore access → create shared capacity → learn through operation.", 9.5, "blue", True)
    footer(s, "Conceptual delivery path; stage logic is a strategy hypothesis, not a measured project schedule.")

    # 5 inner page — data-led energy constraint, intentionally different from the drawings.
    s = base(prs, 5, "04 / ENERGY CONSTRAINT", "Energy is the first constraint", "Before adding new capacity, reduce the amount of energy the building needs to perform everyday life.")
    tx(s, 0.82, 1.82, 4.4, 0.16, "ILLUSTRATIVE LOAD INDEX / EXISTING = 100", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.08, 7.20, 0.025, "blue")
    for x, lab in [(2.28, "0"), (3.60, "25"), (4.92, "50"), (6.24, "75"), (7.56, "100")]:
        line(s, x, 2.48, 0.012, 3.12, "rule")
        tx(s, x - 0.22, 5.82, 0.58, 0.14, lab, 7.2, "muted", False, "Consolas", "center")
    rows = [("EXISTING", 1.00, "uncoordinated demand", "brick"), ("WRAP", 0.68, "envelope first", "blue"), ("REWIRE", 0.54, "plant + controls", "green"), ("SHARE", 0.47, "shared load / care", "blue")]
    for i, (lab, value, detail, col) in enumerate(rows):
        y = 2.70 + i * 0.72
        tx(s, 0.86, y + 0.04, 1.18, 0.16, lab, 8.5, col, True, "Consolas")
        rect(s, 2.28, y, 5.28, 0.26, fill=C["sand"], C=C)
        rect(s, 2.28, y, 5.28 * value, 0.26, fill=C[col], C=C)
        tx(s, 7.64, y + 0.02, 0.62, 0.16, str(int(value * 100)), 9.5, col, True, "Aptos Display", "right")
        tx(s, 0.86, y + 0.28, 1.18, 0.14, detail, 7.8, "muted", False, "Consolas")
    tx(s, 0.86, 6.38, 7.1, 0.20, "Illustrative sequence: reduce loss → rewire capacity → share the gain.", 11.5, "blue", True)
    tx(s, 8.74, 1.84, 3.2, 0.18, "THE STRATEGIC READ", 9, "brick", True, "Consolas")
    line(s, 8.74, 2.12, 3.1, 0.03, "brick")
    tx(s, 8.74, 2.56, 3.25, 0.82, "Reduce loss\nbefore adding capacity.", 22, "ink", True)
    ml(s, 8.76, 3.62, 3.05, 0.74, ["A plant upgrade can improve", "efficiency; an envelope upgrade", "changes the demand itself."], 11.2, "muted", 1.16)
    line(s, 8.74, 4.68, 3.1, 0.025, "blue")
    for y, n, word, desc, col in [(5.02, "01", "DEMAND", "lower the peak", "brick"), (5.48, "02", "CAPACITY", "right-size the plant", "blue"), (5.94, "03", "RESILIENCE", "keep options open", "green")]:
        tx(s, 8.76, y, 0.34, 0.15, n, 8.2, col, True, "Consolas")
        tx(s, 9.34, y, 1.0, 0.15, word, 8.2, col, True, "Consolas")
        tx(s, 10.66, y - 0.01, 1.2, 0.16, desc, 9.7, "ink")
    line(s, 8.74, 6.38, 3.1, 0.025, "blue")
    tx(s, 8.74, 6.58, 3.4, 0.18, "The first retrofit is a demand decision.", 10.8, "blue", True)
    footer(s, "Illustrative load index for visual explanation only; not a measured building-energy result.")

    # 6 inner page — before/after ground-floor plan, spatial rather than card-led.
    s = base(prs, 6, "05 / SHARED GROUND", "The ground floor is the social switch", "A small change at the threshold can turn circulation space into shared civic capacity.")
    tx(s, 0.82, 1.82, 4.4, 0.16, "GROUND-FLOOR PLAN / BEFORE → AFTER", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.08, 11.45, 0.025, "blue")

    # Existing plan.
    tx(s, 0.98, 2.46, 2.9, 0.18, "01 / EXISTING", 9, "brick", True, "Consolas")
    tx(s, 0.98, 2.78, 2.9, 0.18, "sealed threshold", 12, "ink", True)
    rect(s, 0.98, 3.24, 4.05, 2.04, fill=C["sand"], line=C["ink"], C=C)
    # Walls and dead-end circulation.
    line(s, 1.20, 3.52, 3.60, 0.05, "ink")
    line(s, 1.20, 4.20, 3.60, 0.05, "ink")
    line(s, 1.20, 4.88, 3.60, 0.05, "ink")
    line(s, 2.10, 3.52, 0.05, 1.36, "ink")
    line(s, 3.30, 4.20, 0.05, 0.68, "ink")
    rect(s, 1.28, 3.62, 0.56, 0.38, fill=C["brick"], C=C)
    rect(s, 3.52, 4.34, 0.70, 0.32, fill=C["brick"], C=C)
    tx(s, 1.28, 5.52, 3.6, 0.16, "lobby / parking edge / no shared program", 8.2, "muted", False, "Consolas")
    tx(s, 0.98, 5.98, 4.1, 0.18, "Movement passes through. Nothing stays.", 11.2, "brick", True)

    # Transformation arrow and principle.
    arrow(s, 5.24, 4.18, 0.82, 0, C["blue"], C=C)
    tx(s, 5.20, 4.56, 0.98, 0.16, "OPEN", 8.2, "blue", True, "Consolas", "center")
    line(s, 5.68, 3.22, 0.025, 2.02, "blue")

    # Proposed plan.
    tx(s, 6.34, 2.46, 3.2, 0.18, "02 / SHARED GROUND", 9, "green", True, "Consolas")
    tx(s, 6.34, 2.78, 3.2, 0.18, "porous threshold", 12, "ink", True)
    rect(s, 6.34, 3.24, 4.05, 2.04, fill=C["sand"], line=C["ink"], C=C)
    # Open plan with connected program bands.
    line(s, 6.56, 3.52, 3.60, 0.05, "ink")
    line(s, 6.56, 4.78, 3.60, 0.05, "ink")
    rect(s, 6.70, 3.68, 0.78, 0.72, fill=C["green"], C=C)
    rect(s, 7.66, 3.68, 1.05, 0.72, fill=C["blue"], C=C)
    rect(s, 8.88, 3.68, 1.20, 0.72, fill=C["brick"], C=C)
    rect(s, 7.02, 4.96, 2.10, 0.18, fill=C["green"], C=C)
    line(s, 7.10, 4.52, 2.70, 0.025, "blue")
    arrow(s, 7.28, 4.34, 0.62, 0, C["blue"], C=C)
    arrow(s, 8.44, 4.34, 0.62, 0, C["blue"], C=C)
    tx(s, 6.64, 5.52, 3.8, 0.16, "lobby / care / terrace / flexible room", 8.2, "muted", False, "Consolas")
    tx(s, 6.34, 5.98, 4.1, 0.18, "Movement meets. Something can happen.", 11.2, "green", True)

    # Bottom editorial thesis.
    line(s, 0.98, 6.48, 10.08, 0.025, "blue")
    tx(s, 0.98, 6.68, 10.8, 0.18, "The retrofit does not only improve the envelope; it increases the number of useful encounters the building can support.", 10.5, "blue", True)
    footer(s, "Conceptual plan comparison; program labels describe a strategy hypothesis, not a built project outcome.")

    # 7 inner page — a day-in-the-life storyboard, time as the organizing principle.
    s = base(prs, 7, "06 / DAILY LIFE", "A vertical city is measured in moments", "The retrofit succeeds when the building supports more useful moments across an ordinary day.")
    tx(s, 0.82, 1.84, 3.8, 0.16, "DAY STORYBOARD / OCCUPIED BUILDING", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.12, 11.45, 0.025, "blue")
    line(s, 1.12, 4.12, 10.42, 0.035, "ink")
    moments = [(1.28, "07:30", "LEAVE", "threshold becomes\nlegible", "brick"), (3.72, "12:15", "MEET", "lobby becomes\nshared ground", "green"), (6.16, "17:45", "RETURN", "lift becomes\na social hinge", "blue"), (8.60, "20:30", "CARE", "terrace becomes\nvisible support", "brick"), (10.78, "22:00", "REST", "comfort becomes\nquiet infrastructure", "green")]
    for i, (x, time, verb, body, col) in enumerate(moments):
        oval(s, x, 3.80, 0.64, 0.64, fill=C[col], C=C)
        tx(s, x + 0.07, 4.03, 0.50, 0.12, str(i + 1).zfill(2), 8, "paper", True, "Consolas", "center")
        tx(s, x - 0.20, 3.26, 1.04, 0.16, time, 8.5, col, True, "Consolas", "center")
        tx(s, x - 0.32, 4.76, 1.28, 0.18, verb, 10.5, col, True, "Aptos Display", "center")
        tx(s, x - 0.42, 5.12, 1.48, 0.42, body, 9.5, "ink", True, "Aptos", "center")
        if i < 4: arrow(s, x + 0.82, 4.08, 1.46, 0, C["blue"], C=C)
    line(s, 0.98, 2.76, 10.80, 0.014, "rule")
    tx(s, 0.98, 2.48, 3.4, 0.16, "FROM MOVEMENT TO CAPACITY", 8.5, "brick", True, "Consolas")
    tx(s, 0.98, 6.22, 10.8, 0.20, "A retrofit is not finished when the construction ends; it is finished when daily life has more room to work.", 12, "blue", True)
    footer(s, "Conceptual day storyboard; moments describe intended user experience, not measured post-occupancy results.")

    # 8 inner page — three retrofit archetypes as architectural objects.
    s = base(prs, 8, "07 / ARCHETYPES", "There is no single retrofit", "The right depth depends on capital, disruption, and the amount of shared capacity the building needs to recover.")
    tx(s, 0.82, 1.84, 3.8, 0.16, "THREE INTERVENTION DEPTHS", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.12, 11.45, 0.025, "blue")
    archetypes = [(1.04, "LIGHT", "protect", "skin + controls", 1.62, "brick"), (4.34, "DEEP", "reconnect", "skin + threshold + care", 2.30, "blue"), (7.64, "DISTRICT", "share", "building + neighbourhood", 3.00, "green")]
    for x, name, verb, detail, h, col in archetypes:
        y = 5.72 - h
        rect(s, x, y, 2.10, h, fill=C["sand"], line=C[col], C=C)
        rect(s, x + 0.86, y + 0.18, 0.38, h - 0.36, fill=C["ink"], C=C)
        for j in range(max(2, int(h * 3))):
            yy = y + 0.34 + j * 0.38
            line(s, x + 0.18, yy, 1.74, 0.012, "rule")
        line(s, x, y, 2.10, 0.07, col)
        tx(s, x, 5.98, 2.10, 0.18, name, 10, col, True, "Consolas", "center")
        tx(s, x, 6.30, 2.10, 0.18, verb, 13, "ink", True, "Aptos Display", "center")
        tx(s, x, 6.62, 2.10, 0.14, detail, 7.8, "muted", False, "Consolas", "center")
    line(s, 1.04, 3.02, 8.70, 0.025, "blue")
    tx(s, 9.98, 2.62, 2.0, 0.16, "SELECTION LOGIC", 8.5, "brick", True, "Consolas")
    line(s, 9.98, 2.90, 2.0, 0.03, "brick")
    tx(s, 9.98, 3.30, 2.2, 0.68, "Depth follows\nshared need.", 19, "ink", True)
    ml(s, 9.98, 4.24, 2.0, 0.68, ["Light protects.", "Deep reconnects.", "District shares."], 10.8, "muted", 1.18)
    line(s, 9.98, 5.32, 2.0, 0.025, "blue")
    tx(s, 9.98, 5.62, 2.1, 0.45, "Choose the smallest intervention that changes the operating model.", 10.2, "blue", True)
    footer(s, "Conceptual archetypes; intervention depth and labels are strategic options, not project specifications.")

    # 9 inner page — phasing while residents remain in place.
    s = base(prs, 9, "08 / PHASING", "The building stays occupied while the work moves", "Construction phasing is part of the design: every intervention must have a resident-safe next move.")
    tx(s, 0.82, 1.84, 3.8, 0.16, "RESIDENT-IN-PLACE PHASING", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.12, 11.45, 0.025, "blue")
    tx(s, 0.98, 2.66, 1.8, 0.16, "MONTHS", 8, "muted", True, "Consolas")
    line(s, 2.18, 2.78, 9.40, 0.018, "rule")
    for x, lab in [(2.18, "0"), (4.52, "6"), (6.86, "12"), (9.20, "18"), (11.58, "24")]:
        line(s, x, 2.62, 0.014, 3.55, "rule")
        tx(s, x - 0.16, 2.36, 0.34, 0.14, lab, 8, "muted", True, "Consolas", "center")
    phases = [("01", "PREPARE", 2.18, 2.10, "survey + resident pact", "brick"), ("02", "WRAP", 3.92, 2.76, "envelope + plant", "blue"), ("03", "OPEN", 6.48, 2.20, "thresholds + shared ground", "green"), ("04", "LEARN", 8.70, 2.28, "operate + adjust", "brick")]
    for n, name, x, w, detail, col in phases:
        y = 3.34 + (int(n) - 1) * 0.40
        rect(s, x, y, w, 0.46, fill=C[col], C=C)
        tx(s, x + 0.12, y + 0.14, 0.48, 0.12, n, 7.4, "paper", True, "Consolas")
        tx(s, x + 0.56, y + 0.12, w - 0.68, 0.16, name, 9, "paper", True, "Consolas")
        tx(s, x, y + 0.60, w, 0.16, detail, 8.2, "muted", False, "Consolas")
    tx(s, 0.98, 5.58, 1.8, 0.18, "RESIDENT IMPACT", 8.5, "brick", True, "Consolas")
    line(s, 2.18, 5.70, 9.40, 0.025, "brick")
    tx(s, 2.18, 6.08, 2.25, 0.18, "noise / access", 10, "ink", True)
    arrow(s, 4.56, 6.14, 0.70, 0, C["blue"], C=C)
    tx(s, 5.42, 6.08, 2.25, 0.18, "temporary disruption", 10, "ink", True)
    arrow(s, 7.80, 6.14, 0.70, 0, C["blue"], C=C)
    tx(s, 8.68, 6.08, 2.35, 0.18, "permanent capacity", 10, "blue", True)
    footer(s, "Conceptual phasing sequence; durations and impacts are illustrative planning assumptions.")

    # 10 inner page — waterfall-like economics and sensitivity.
    s = base(prs, 10, "09 / ECONOMICS", "Sequence changes the economics", "The same interventions can produce different value depending on what happens first.")
    tx(s, 0.82, 1.84, 4.0, 0.16, "ILLUSTRATIVE VALUE BRIDGE / INDEX", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.12, 6.90, 0.025, "blue")
    line(s, 1.06, 5.74, 6.18, 0.025, "ink")
    bars = [("BASE", 1.00, 2.18, "brick"), ("WRAP", -0.24, 3.22, "blue"), ("REWIRE", -0.16, 4.10, "green"), ("SHARE", 0.22, 4.84, "brick"), ("VALUE", 0.82, 6.00, "blue")]
    for i, (lab, delta, x, col) in enumerate(bars):
        height = abs(delta) * 1.65 + 0.34
        y = 5.74 - height if delta >= 0 else 5.74
        rect(s, x, y, 0.78, height, fill=C[col], C=C)
        tx(s, x - 0.12, 5.96 if delta >= 0 else 5.42, 1.02, 0.16, lab, 7.8, col, True, "Consolas", "center")
        tx(s, x - 0.12, y - 0.24 if delta >= 0 else 5.64, 1.02, 0.16, ("+" if delta > 0 else "") + str(int(delta * 100)), 8.5, col, True, "Aptos Display", "center")
        if i < 4: line(s, x + 0.78, y + height / 2, 0.26, 0.014, "rule")
    tx(s, 0.98, 6.36, 6.2, 0.18, "Illustrative index: demand reduction creates the headroom for shared capacity.", 10.5, "blue", True)
    tx(s, 8.26, 1.84, 3.7, 0.18, "THE ECONOMIC READ", 9, "brick", True, "Consolas")
    line(s, 8.26, 2.12, 3.7, 0.03, "brick")
    tx(s, 8.26, 2.52, 3.6, 0.78, "Order is\na financial lever.", 22, "ink", True)
    ml(s, 8.28, 3.54, 3.3, 0.72, ["Reduce peak demand.", "Right-size capacity.", "Invest in shared use."], 11.2, "muted", 1.18)
    line(s, 8.26, 4.66, 3.7, 0.025, "blue")
    tx(s, 8.26, 4.96, 3.5, 0.16, "SENSITIVITY / WHAT MOVES FIRST", 8.5, "blue", True, "Consolas")
    for y, lab, val, col in [(5.34, "CAPEX", "high", "brick"), (5.72, "DISRUPTION", "medium", "blue"), (6.10, "SHARED VALUE", "later", "green")]:
        tx(s, 8.28, y, 1.14, 0.15, lab, 8.0, col, True, "Consolas")
        tx(s, 9.72, y, 1.52, 0.15, val, 10.2, "ink", True)
    footer(s, "Illustrative value bridge; no project IRR, payback, or measured savings are being claimed.")

    # 11 inner page — decision matrix with trade-offs made visible.
    s = base(prs, 11, "10 / DECISION", "Choose by trade-off, not by score", "A credible retrofit brief makes the conflicts visible before it recommends a depth.")
    tx(s, 0.82, 1.84, 3.7, 0.16, "OPTION MATRIX / STRATEGIC FIT", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.12, 11.45, 0.025, "blue")
    tx(s, 1.18, 2.62, 1.2, 0.16, "HIGH", 8, "muted", True, "Consolas")
    tx(s, 1.18, 5.62, 1.2, 0.16, "LOW", 8, "muted", True, "Consolas")
    line(s, 1.76, 2.76, 5.70, 0.018, "rule"); line(s, 1.76, 5.42, 5.70, 0.018, "rule")
    line(s, 1.76, 2.76, 0.018, 2.66, "rule"); line(s, 7.44, 2.76, 0.018, 2.66, "rule")
    for x, lab in [(1.76, "LOW"), (4.58, "MEDIUM"), (7.28, "HIGH")]:
        tx(s, x - 0.32, 5.66, 0.70, 0.14, lab, 7.8, "muted", True, "Consolas", "center")
    tx(s, 3.02, 6.02, 3.4, 0.16, "CAPITAL + DISRUPTION", 8.3, "blue", True, "Consolas", "center")
    tx(s, 0.94, 3.98, 0.6, 0.16, "SHARED\nVALUE", 7.8, "blue", True, "Consolas", "center")
    options = [(3.12, 4.86, "LIGHT", "protect", "brick"), (5.14, 3.98, "DEEP", "reconnect", "blue"), (6.66, 3.16, "DISTRICT", "share", "green")]
    for x, y, name, verb, col in options:
        oval(s, x, y, 0.54, 0.54, fill=C[col], C=C)
        tx(s, x - 0.46, y + 0.70, 1.46, 0.16, name, 8.2, col, True, "Consolas", "center")
        tx(s, x - 0.46, y + 0.96, 1.46, 0.16, verb, 9.5, "ink", True, "Aptos Display", "center")
    tx(s, 8.42, 2.62, 3.4, 0.18, "THE DECISION RULE", 9, "brick", True, "Consolas")
    line(s, 8.42, 2.90, 3.4, 0.03, "brick")
    tx(s, 8.42, 3.34, 3.4, 0.72, "Pick the option\nthat changes the system.", 19, "ink", True)
    ml(s, 8.44, 4.32, 3.1, 0.72, ["Light: minimum disruption", "Deep: balanced value", "District: maximum capacity"], 10.8, "muted", 1.18)
    line(s, 8.42, 5.42, 3.4, 0.025, "blue")
    tx(s, 8.42, 5.76, 3.4, 0.36, "Trade-offs are not a weakness; they are the decision.", 10.8, "blue", True)
    footer(s, "Conceptual option matrix; positions express strategic trade-offs, not a quantified project ranking.")

    # 12 inner page — operating model network.
    s = base(prs, 12, "11 / OPERATING MODEL", "The retrofit needs an operator", "Capital can upgrade the building once; an operating model keeps the new capacity useful.")
    tx(s, 0.82, 1.84, 3.8, 0.16, "ACTORS / FLOWS / ACCOUNTABILITY", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.12, 11.45, 0.025, "blue")
    oval(s, 5.06, 3.26, 2.54, 1.20, fill=C["ink"], C=C)
    tx(s, 5.26, 3.58, 2.12, 0.18, "BUILDING\nOPERATOR", 13, "paper", True, "Aptos Display", "center")
    tx(s, 5.26, 4.04, 2.12, 0.14, "measure / maintain / adapt", 7.8, "sand", False, "Consolas", "center")
    actors = [(1.08, 2.64, "RESIDENTS", "feedback", "brick"), (1.08, 4.82, "OWNER", "capital", "blue"), (9.98, 2.64, "ENERGY TEAM", "performance", "green"), (9.98, 4.82, "CARE PARTNERS", "services", "brick")]
    for x, y, name, flow, col in actors:
        oval(s, x, y, 1.72, 0.62, fill=C["paper"], line=C[col], C=C)
        tx(s, x + 0.06, y + 0.18, 1.60, 0.14, name, 7.8, col, True, "Consolas", "center")
        tx(s, x + 0.06, y + 0.88, 1.60, 0.14, flow, 8.5, "muted", False, "Consolas", "center")
        if x < 3: arrow(s, x + 1.86, y + 0.28, 2.02, 0, C[col], C=C)
        else: arrow(s, 7.76, y + 0.28, 2.02, 0, C[col], C=C)
    line(s, 6.32, 4.46, 0.014, 0.88, "blue")
    tx(s, 4.06, 5.58, 5.30, 0.18, "ACCOUNTABILITY LOOP", 8.5, "blue", True, "Consolas", "center")
    line(s, 3.74, 5.90, 5.94, 0.025, "blue")
    tx(s, 3.78, 6.26, 5.86, 0.18, "data → decision → maintenance → lived outcome → new data", 10.2, "blue", True, "Aptos", "center")
    footer(s, "Conceptual operating model; actor roles describe the governance hypothesis for a retrofit programme.")

    # 13 inner page — 36-month delivery path with gates.
    s = base(prs, 13, "12 / DELIVERY", "A 36-month path with evidence at every gate", "The programme becomes investable when each horizon has a decision, an owner, and proof of progress.")
    tx(s, 0.82, 1.84, 3.8, 0.16, "THREE HORIZONS / FOUR GATES", 8.5, "blue", True, "Consolas")
    line(s, 0.82, 2.12, 11.45, 0.025, "blue")
    line(s, 1.18, 4.16, 10.72, 0.045, "ink")
    for x, lab in [(1.18, "0"), (4.72, "12"), (8.26, "24"), (11.90, "36")]:
        line(s, x, 3.72, 0.018, 0.88, "blue")
        tx(s, x - 0.30, 3.30, 0.60, 0.16, lab, 9, "blue", True, "Aptos Display", "center")
        tx(s, x - 0.40, 4.54, 0.80, 0.14, "MONTHS", 7, "muted", True, "Consolas", "center")
    horizons = [(1.18, 3.54, "DIAGNOSE", "baseline + resident pact", "brick"), (4.72, 3.54, "DELIVER", "wrap + open + share", "blue"), (8.26, 3.54, "OPERATE", "measure + adapt", "green")]
    for x, w, name, detail, col in horizons:
        rect(s, x, 2.68, w, 0.38, fill=C[col], C=C)
        tx(s, x, 2.79, w, 0.14, name, 8.5, "paper", True, "Consolas", "center")
        tx(s, x, 5.00, w, 0.18, detail, 9.0, "ink", True, "Aptos", "center")
    gates = [(2.90, "G1", "brief locked", "brick"), (6.44, "G2", "pilot signed off", "blue"), (9.98, "G3", "shared space live", "green"), (11.90, "G4", "evidence review", "brick")]
    for x, n, label, col in gates:
        oval(s, x, 5.70, 0.54, 0.54, fill=C[col], C=C)
        tx(s, x, 5.90, 0.54, 0.12, n, 7.6, "paper", True, "Consolas", "center")
        tx(s, x - 0.48, 6.42, 1.50, 0.14, label, 8.0, col, True, "Consolas", "center")
    tx(s, 0.98, 6.88, 10.8, 0.16, "A gate is not a date. It is a decision supported by evidence.", 10.8, "blue", True)
    footer(s, "Conceptual 36-month path; milestones and gates are illustrative programme design assumptions.")

    # 14 closing — return to the cover image as a conclusion, not a repeat cover.
    s = prs.slides.add_slide(prs.slide_layouts[6])
    cover_image(s, 7.04, 0.70, 5.45, 5.92, str(ASSET / "vertical_city_cover.png"))
    rect(s, 0, 0, 6.78, 7.5, fill=C["paper"], C=C)
    line(s, 0.84, 0.88, 2.92, 0.045, "blue")
    tx(s, 0.86, 1.18, 3.7, 0.18, "THE TAKEAWAY / CASE 07", 9, "blue", True, "Consolas")
    tx(s, 0.86, 1.72, 5.2, 1.12, "Upgrade in place.\nKeep life in place.", 28, "ink", True)
    ml(s, 0.88, 3.28, 4.92, 0.98, ["A retrofit becomes strategic when", "the building is treated as a connected", "vertical city—not a façade project."], 14, "muted", 1.20)
    line(s, 0.88, 4.76, 4.55, 0.03, "blue")
    for y, n, textv, col in [(5.18, "01", "Reduce loss before adding capacity.", "brick"), (5.62, "02", "Design the shared ground.", "green"), (6.06, "03", "Operate the new capacity.", "blue")]:
        tx(s, 0.90, y, 0.34, 0.15, n, 8.2, col, True, "Consolas")
        tx(s, 1.48, y - 0.01, 4.4, 0.18, textv, 10.5, "ink", True)
    tx(s, 7.08, 6.86, 5.2, 0.14, "CONCEPTUAL HERO VISUAL / NOT A REAL PROJECT", 7.2, "blue", True, "Consolas")
    tx(s, 0.88, 7.02, 5.2, 0.14, "VERTICAL CITY RETROFIT · 2026", 7.2, "muted", False, "Consolas")

    prs.save(str(OUT))
    print(OUT)


if __name__ == "__main__":
    main()
