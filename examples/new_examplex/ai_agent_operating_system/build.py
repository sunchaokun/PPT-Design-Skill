"""Flagship case 004: AI Agent Operating System, Build Mode."""

from pathlib import Path

from pptx_designer import Presentation
from pptx_designer.tools.charts import comparison_bars
from pptx_designer.tools.layout import page_number
from pptx_designer.tools.shapes import arrow, diamond, oval, rect, rrect
from pptx_designer.tools.text import multiline, text


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output" / "agent_operating_system.pptx"

C = {
    "bg": "#07111F",
    "panel": "#0E1D30",
    "panel2": "#12263C",
    "ink": "#F3F7FB",
    "muted": "#8EA5BC",
    "grid": "#20364B",
    "line": "#35516A",
    "cyan": "#36D8E8",
    "lime": "#B9F27C",
    "amber": "#FFBF69",
    "coral": "#FF6B6B",
    "white": "#FFFFFF",
}


def t(slide, left, top, width, height, value, size=14, color="ink", bold=False, name="Aptos", align=None):
    kwargs = dict(font_size=size, color=color, bold=bold, font_name=name, C=C)
    if align is not None:
        kwargs["align"] = align
    return text(slide, left, top, width, height, value, **kwargs)


def ml(slide, left, top, width, height, lines, size=14, color="muted", spacing=1.1):
    return multiline(slide, left, top, width, height, lines, font_size=size, color=color, C=C, line_spacing=spacing)


def line(slide, left, top, width, height=0.018, color="line"):
    rect(slide, left, top, width, height, fill=C[color], C=C)


def chip(slide, left, top, width, value, color="cyan", fill="panel2"):
    rrect(slide, left, top, width, 0.28, fill=C[fill], line=C[color], C=C)
    t(slide, left, top + 0.055, width, 0.16, value, 9, color, True, "Consolas", "center")


def node(slide, left, top, width, height, title, body=None, accent="cyan", tag=None, fill="panel"):
    rrect(slide, left, top, width, height, fill=C[fill], line=C[accent], C=C)
    rect(slide, left, top, 0.055, height, fill=C[accent], C=C)
    t(slide, left + 0.18, top + 0.16, width - 0.32, 0.25, title, 14, "ink", True)
    if body:
        body_height = max(0.18, height - (0.92 if tag and height >= 1.15 else 0.55))
        ml(slide, left + 0.18, top + 0.53, width - 0.32, body_height, body if isinstance(body, list) else body.split("\n"), 10 if height < 1.15 else 11, "muted", 1.0)
    if tag and height >= 1.15:
        chip(slide, left + 0.18, top + height - 0.39, min(width - 0.36, 1.25), tag, accent, fill)


def base(prs, number, kicker, title, subtitle=None, dark=True):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["bg"], C=C)
    # blueprint grid: restrained, structural, and consistent across the deck
    for x in [0.65, 3.05, 5.45, 7.85, 10.25, 12.65]:
        line(slide, x, 0.42, 0.012, 6.35, "grid")
    for y in [1.45, 3.05, 4.65, 6.25]:
        line(slide, 0.45, y, 12.45, 0.012, "grid")
    chip(slide, 0.7, 0.45, 2.1, kicker, "cyan", "bg")
    if title:
        t(slide, 0.7, 0.9, 11.5, 0.52, title, 22, "ink", True)
    if subtitle:
        t(slide, 0.72, 1.29, 10.7, 0.25, subtitle, 11, "muted")
    page_number(slide, number, 12, C=C)
    t(slide, 12.1, 0.55, 0.55, 0.2, "AOS / 01", 9, "muted", False, "Consolas", "right")
    return slide


def footer(slide, value):
    t(slide, 0.72, 7.05, 11.3, 0.18, value, 8, "muted", False, "Consolas")


def main():
    prs = Presentation()

    # 1 — cover: system map as the visual hook
    slide = base(prs, 1, "SYSTEMS BLUEPRINT", None)
    t(slide, 0.72, 1.18, 6.1, 1.7, "AI AGENT\nOPERATING SYSTEM", 35, "ink", True)
    t(slide, 0.76, 3.18, 5.3, 0.65, "From prompt chains to\ncontrolled autonomy.", 19, "cyan", True)
    line(slide, 0.75, 4.35, 1.35, 0.06, "lime")
    ml(slide, 0.76, 4.65, 4.4, 0.9, ["A field guide to choosing the", "smallest architecture that can carry the work."], 13, "muted", 1.18)
    # system map
    t(slide, 7.3, 1.22, 4.3, 0.25, "THE SYSTEM IS THE PRODUCT", 10, "muted", True, "Consolas")
    node(slide, 7.25, 1.75, 4.8, 0.82, "USER INTENT", ["objective · context · constraints"], "cyan", "INPUT")
    arrow(slide, 9.35, 2.68, 0.35, 0.28, fill=C["cyan"], C=C)
    node(slide, 6.7, 3.08, 5.9, 1.12, "CONTROL PLANE", ["route · plan · approve · evaluate"], "amber", "DECIDE")
    arrow(slide, 9.35, 4.4, 0.35, 0.28, fill=C["lime"], C=C)
    node(slide, 7.25, 4.82, 4.8, 0.82, "EXECUTION PLANE", ["tools · memory · state · outputs"], "lime", "ACT")
    t(slide, 0.76, 6.75, 5.4, 0.2, "A decision-led technical editorial / 12 pages", 9, "muted", False, "Consolas")
    footer(slide, "SOURCE FRAME: Anthropic, Building Effective AI Agents; Trustworthy Agents in Practice")

    # 2 — thesis: three layers
    slide = base(prs, 2, "01 / THESIS", "Autonomy is not a mood. It is a system boundary.", "The design question is not “how agentic?” — it is “where should control live?”")
    layers = [
        (1.0, 2.0, 11.1, 0.9, "WORK", "What must be done", "task · context · definition of done", "cyan"),
        (1.65, 3.2, 10.45, 1.0, "CONTROL", "Who decides what happens next", "route · plan · permissions · escalation", "amber"),
        (2.3, 4.55, 9.8, 1.05, "LEARNING", "How the system gets better", "trace · eval · feedback · release gate", "lime"),
    ]
    for x, y, w, h, title, head, body, accent in layers:
        rrect(slide, x, y, w, h, fill=C["panel"], line=C[accent], C=C)
        t(slide, x + 0.25, y + 0.23, 1.3, 0.25, title, 11, accent, True, "Consolas")
        t(slide, x + 1.8, y + 0.18, 4.0, 0.32, head, 18, "ink", True)
        t(slide, x + 6.4, y + 0.25, 3.9, 0.22, body, 11, "muted", False, "Consolas")
    t(slide, 0.98, 6.2, 11.1, 0.35, "A system earns autonomy only when its boundaries, feedback, and failure modes are visible.", 16, "ink", True, align="center")
    footer(slide, "DESIGN PRINCIPLE / MAKE THE BOUNDARY VISIBLE")

    # 3 — decision matrix
    slide = base(prs, 3, "02 / DECISION", "Choose architecture by problem shape, not by fashion.", "Two questions predict most of the design: can the steps be known, and can the work split safely?")
    # axes
    line(slide, 1.4, 5.8, 8.1, 0.025, "line")
    line(slide, 1.4, 2.25, 0.025, 3.55, "line")
    arrow(slide, 9.0, 5.68, 0.5, 0.22, fill=C["cyan"], C=C)
    arrow(slide, 1.3, 1.75, 0.22, 0.5, fill=C["cyan"], C=C)
    t(slide, 3.35, 6.03, 4.4, 0.25, "PARALLELISM  →", 10, "cyan", True, "Consolas", "center")
    t(slide, 0.7, 3.1, 0.25, 2.0, "U\nN\nC\nE\nR\nT\nA\nI\nN\nT\nY", 9, "cyan", True, "Consolas", "center")
    # quadrants
    node(slide, 1.85, 4.45, 3.05, 1.0, "WORKFLOW", ["known steps · low variance"], "line", "CHAIN")
    node(slide, 5.55, 4.45, 3.1, 1.0, "PARALLEL", ["independent branches"], "lime", "FAN-OUT")
    node(slide, 1.85, 2.65, 3.05, 1.0, "SINGLE AGENT", ["open-ended · tool use"], "cyan", "LOOP")
    node(slide, 5.55, 2.65, 3.1, 1.0, "ORCHESTRATED", ["specialists · synthesis"], "amber", "DELEGATE")
    rrect(slide, 9.55, 2.35, 2.55, 3.2, fill=C["panel2"], line=C["coral"], C=C)
    t(slide, 9.85, 2.7, 1.9, 0.3, "ANTI-PATTERN", 10, "coral", True, "Consolas")
    t(slide, 9.85, 3.28, 1.75, 0.8, "More agents\ndoes not mean\nmore system.", 18, "ink", True)
    ml(slide, 9.85, 4.98, 1.8, 0.38, ["Complexity is the cost."], 9, "muted")
    footer(slide, "DECISION RULE / START WITH THE SIMPLEST PATTERN THAT CAN SURVIVE THE WORK")

    # 4 — operating system architecture
    slide = base(prs, 4, "03 / ARCHITECTURE", "The operating system has two planes — and one feedback spine.", "Control decides. Execution acts. Observability closes the loop.")
    # outer boundary
    rrect(slide, 0.85, 1.9, 11.7, 4.65, fill=C["panel"], line=C["line"], C=C)
    chip(slide, 1.1, 2.12, 2.0, "CONTROL PLANE", "amber", "panel")
    node(slide, 1.1, 2.65, 2.15, 1.0, "ROUTER", ["classify intent", "select path"], "cyan", "DECIDE")
    node(slide, 3.55, 2.65, 2.15, 1.0, "ORCHESTRATOR", ["delegate work", "merge context"], "amber", "PLAN")
    node(slide, 6.0, 2.65, 2.15, 1.0, "EVALUATOR", ["score output", "trigger revise"], "lime", "JUDGE")
    arrow(slide, 3.25, 3.0, 0.25, 0.22, fill=C["line"], C=C)
    arrow(slide, 5.7, 3.0, 0.25, 0.22, fill=C["line"], C=C)
    t(slide, 8.62, 2.78, 2.9, 0.45, "permissions\n& escalation", 12, "amber", True, "Consolas", "center")
    line(slide, 1.1, 4.05, 10.95, 0.02, "line")
    chip(slide, 1.1, 4.27, 2.0, "EXECUTION PLANE", "lime", "panel")
    for x, title, body, accent in [(1.1, "SPECIALISTS", "research · code · design", "cyan"), (4.05, "TOOLS", "APIs · files · runtime", "lime"), (7.0, "STATE", "memory · artifacts · trace", "amber")]:
        node(slide, x, 4.72, 2.45, 0.93, title, [body], accent, "ACT")
    arrow(slide, 3.7, 5.0, 0.28, 0.22, fill=C["line"], C=C)
    arrow(slide, 6.65, 5.0, 0.28, 0.22, fill=C["line"], C=C)
    line(slide, 9.95, 4.75, 0.02, 0.9, "line")
    t(slide, 10.25, 4.82, 1.75, 0.45, "TRACE\nSPINE", 12, "lime", True, "Consolas", "center")
    footer(slide, "SYSTEM VIEW / EVERY ACTION SHOULD BE ROUTABLE, PERMISSIONED, AND TRACEABLE")

    # 5 — sequential pattern
    slide = base(prs, 5, "04 / PATTERN 01", "Prompt chaining: make the path explicit.", "Use a sequence when the work is stable enough to describe before execution.")
    stages = [(0.9, "01", "RETRIEVE", "bring the right context", "cyan"), (3.35, "02", "DRAFT", "produce a first artifact", "lime"), (5.8, "03", "CHECK", "validate against criteria", "amber"), (8.25, "04", "PACKAGE", "format for the user", "cyan")]
    for i, (x, no, title, body, accent) in enumerate(stages):
        node(slide, x, 2.55, 2.0, 1.55, title, [body], accent, no)
        if i < 3:
            arrow(slide, x + 2.08, 3.18, 0.28, 0.22, fill=C["line"], C=C)
    line(slide, 1.15, 4.75, 9.0, 0.035, "cyan")
    for x, label in [(1.15, "CONTEXT"), (3.62, "ARTIFACT"), (6.07, "CRITERIA"), (8.52, "OUTPUT")]:
        oval(slide, x, 4.61, 0.28, 0.28, fill=C["cyan"], C=C)
        t(slide, x - 0.25, 5.08, 0.8, 0.18, label, 8, "muted", True, "Consolas", "center")
    rrect(slide, 10.2, 2.45, 2.05, 2.7, fill=C["panel2"], line=C["line"], C=C)
    t(slide, 10.48, 2.75, 1.5, 0.2, "BEST WHEN", 9, "lime", True, "Consolas")
    ml(slide, 10.48, 3.22, 1.45, 1.45, ["steps are known", "handoffs are clear", "failure is local"], 13, "ink", 1.35)
    footer(slide, "PATTERN 01 / LOW VARIANCE · HIGH EXPLAINABILITY")

    # 6 — routing
    slide = base(prs, 6, "05 / PATTERN 02", "Routing: send the task to the right lane.", "A classifier can reduce complexity when the task family is recognizable.")
    node(slide, 0.95, 3.0, 2.2, 1.2, "USER INTENT", ["What kind of work is this?"], "cyan", "INPUT")
    diamond(slide, 4.825, 3.725, 1.55, fill=C["panel2"], line=C["amber"], C=C)
    t(slide, 4.22, 3.38, 1.2, 0.4, "CLASSIFY", 12, "amber", True, "Consolas", "center")
    arrow(slide, 3.25, 3.45, 0.55, 0.22, fill=C["cyan"], C=C)
    lanes = [(6.45, "RESEARCH", "retrieve · cite", "cyan"), (8.8, "BUILD", "code · test", "lime"), (11.15, "REVIEW", "critique · explain", "amber")]
    for x, title, body, accent in lanes:
        arrow(slide, 5.75, 3.1 + (x - 6.45) * 0.08, 0.52, 0.22, fill=C[accent], C=C)
        node(slide, x, 2.45, 1.75, 1.9, title, [body], accent, "LANE")
    line(slide, 4.82, 4.68, 0.02, 1.0, "line")
    t(slide, 3.5, 5.2, 2.6, 0.25, "ROUTING IS A DECISION", 10, "amber", True, "Consolas")
    ml(slide, 3.5, 5.62, 7.2, 0.5, ["The classifier should expose why a lane was chosen — not hide a second agent inside a black box."], 13, "muted")
    footer(slide, "PATTERN 02 / CLASSIFICATION BEFORE EXECUTION")

    # 7 — parallelization
    slide = base(prs, 7, "06 / PATTERN 03", "Parallelization: split only where independence is real.", "Fan out the work, then make synthesis an explicit step.")
    node(slide, 0.9, 3.02, 2.0, 1.1, "QUESTION", ["What should we learn?"], "cyan", "START")
    arrow(slide, 3.05, 3.45, 0.6, 0.22, fill=C["cyan"], C=C)
    branches = [(4.0, 1.98, "A", "MARKET", "signals", "cyan"), (4.0, 3.25, "B", "USERS", "needs", "lime"), (4.0, 4.52, "C", "RISKS", "constraints", "amber")]
    for x, y, letter, title, body, accent in branches:
        oval(slide, x, y, 0.42, 0.42, fill=C[accent], C=C)
        t(slide, x + 0.1, y + 0.12, 0.2, 0.16, letter, 10, "bg", True, "Consolas", "center")
        node(slide, x + 0.62, y - 0.16, 1.75, 0.9, title, [body], accent, "WORKER")
        arrow(slide, 6.55, y + 0.12, 0.55, 0.22, fill=C[accent], C=C)
    node(slide, 7.35, 3.0, 2.2, 1.15, "SYNTHESIZER", ["merge evidence", "resolve conflict"], "lime", "JOIN")
    arrow(slide, 9.7, 3.45, 0.55, 0.22, fill=C["lime"], C=C)
    node(slide, 10.4, 3.0, 2.0, 1.15, "DECISION", ["recommendation", "with provenance"], "amber", "END")
    rrect(slide, 1.0, 5.95, 11.4, 0.52, fill=C["panel2"], line=C["line"], C=C)
    t(slide, 1.25, 6.11, 10.9, 0.18, "FAILURE MODE  /  parallel workers without a synthesis contract produce a pile of answers, not a system.", 10, "coral", True, "Consolas", "center")
    footer(slide, "PATTERN 03 / INDEPENDENCE FIRST · SYNTHESIS SECOND")

    # 8 — orchestrator workers
    slide = base(prs, 8, "07 / PATTERN 04", "Orchestrator–workers: delegate the search space.", "Use hierarchy when the task needs specialists and the plan must adapt as evidence arrives.")
    node(slide, 4.8, 1.9, 3.7, 0.98, "ORCHESTRATOR", ["plan · delegate · synthesize"], "amber", "CONTROL")
    line(slide, 6.62, 2.92, 0.025, 0.55, "amber")
    workers = [(1.0, "RESEARCHER", "find evidence", "cyan"), (4.0, "ANALYST", "compare signals", "lime"), (7.0, "BUILDER", "make artifact", "cyan"), (10.0, "RED TEAM", "challenge result", "coral")]
    for x, title, body, accent in workers:
        line(slide, x + 0.95, 3.48, 0.025, 0.42, accent)
        node(slide, x, 3.95, 2.0, 1.0, title, [body], accent, "SPECIALIST")
    line(slide, 1.95, 5.42, 9.0, 0.025, "line")
    t(slide, 1.0, 5.72, 10.6, 0.26, "The orchestrator owns the objective. Specialists own the search space.", 16, "ink", True, align="center")
    chip(slide, 4.55, 6.3, 4.2, "ADAPTIVE PLAN / SHARED CONTEXT", "amber", "panel2")
    footer(slide, "PATTERN 04 / HIERARCHY CREATES SPECIALIZATION — AND A NEW GOVERNANCE SURFACE")

    # 9 — evaluator optimizer
    slide = base(prs, 9, "08 / PATTERN 05", "Evaluator–optimizer: make quality a loop.", "A useful agent does not only produce an answer; it knows what a better answer means.")
    node(slide, 0.9, 2.48, 2.15, 1.5, "DRAFT", ["artifact / response", "first pass"], "cyan", "WRITE")
    arrow(slide, 3.25, 3.05, 0.62, 0.22, fill=C["cyan"], C=C)
    node(slide, 4.05, 2.48, 2.15, 1.5, "EVALUATE", ["criteria / rubric", "evidence"], "amber", "JUDGE")
    arrow(slide, 6.45, 3.05, 0.62, 0.22, fill=C["amber"], C=C)
    node(slide, 7.25, 2.48, 2.15, 1.5, "REVISE", ["targeted change", "new pass"], "lime", "IMPROVE")
    arrow(slide, 9.65, 3.05, 0.62, 0.22, fill=C["lime"], C=C)
    node(slide, 10.45, 2.48, 1.95, 1.5, "RELEASE", ["meets threshold", "with trace"], "lime", "SHIP")
    # feedback arc
    line(slide, 8.3, 4.3, 0.025, 1.0, "amber")
    line(slide, 2.0, 5.3, 6.3, 0.025, "amber")
    arrow(slide, 1.55, 4.98, 0.45, 0.22, fill=C["amber"], C=C)
    t(slide, 2.35, 5.52, 5.8, 0.25, "FEEDBACK / WHAT FAILED, AND WHY?", 10, "amber", True, "Consolas", "center")
    rrect(slide, 9.45, 4.65, 2.75, 1.0, fill=C["panel2"], line=C["line"], C=C)
    t(slide, 9.75, 4.93, 2.15, 0.2, "QUALITY IS A CONTRACT", 9, "lime", True, "Consolas", "center")
    t(slide, 9.75, 5.28, 2.15, 0.22, "rubric · threshold · stop", 11, "ink", False, "Consolas", "center")
    footer(slide, "PATTERN 05 / EVALUATION TURNS ITERATION INTO A SYSTEM PROPERTY")

    # 10 — human control
    slide = base(prs, 10, "09 / GOVERNANCE", "Human control belongs at the level of intent.", "The safest review point is usually the plan, the permission boundary, and the final release — not every tool call.")
    line(slide, 1.15, 5.7, 10.35, 0.04, "amber")
    points = [(1.25, "OBSERVE", "read-only trace", "cyan"), (3.9, "PLAN", "review intent", "amber"), (6.55, "APPROVE", "grant access", "amber"), (9.2, "INTERVENE", "stop / redirect", "coral")]
    for x, title, body, accent in points:
        oval(slide, x, 5.48, 0.45, 0.45, fill=C[accent], C=C)
        node(slide, x - 0.28, 3.05, 1.95, 1.35, title, [body], accent, "CONTROL")
        line(slide, x + 0.2, 4.42, 0.025, 1.0, accent)
    rrect(slide, 1.2, 2.0, 10.3, 0.62, fill=C["panel2"], line=C["line"], C=C)
    t(slide, 1.45, 2.2, 9.8, 0.2, "A good control surface compresses many low-level actions into a small number of high-value decisions.", 12, "ink", True, align="center")
    footer(slide, "GOVERNANCE / PLAN MODE, PERMISSIONS, AND INTERVENTION ARE PART OF THE UX")

    # 11 — observability and evals
    slide = base(prs, 11, "10 / QUALITY", "If you cannot trace it, you cannot improve it.", "Agent quality lives across turns, tools, state changes, and final artifacts.")
    # event spine
    line(slide, 1.0, 3.15, 10.9, 0.035, "cyan")
    events = [(1.0, "01", "INTENT", "objective"), (3.15, "02", "PLAN", "decision"), (5.3, "03", "TOOLS", "action"), (7.45, "04", "STATE", "change"), (9.6, "05", "OUTPUT", "artifact")]
    for x, no, title, body in events:
        oval(slide, x, 2.92, 0.48, 0.48, fill=C["cyan"], C=C)
        t(slide, x + 0.11, 3.08, 0.24, 0.16, no, 9, "bg", True, "Consolas", "center")
        t(slide, x - 0.15, 3.72, 1.2, 0.2, title, 10, "cyan", True, "Consolas", "center")
        t(slide, x - 0.15, 4.05, 1.2, 0.2, body, 11, "muted", False, "Consolas", "center")
    # release gates
    for x, title, body, accent in [(1.0, "TRACE", "what happened?", "cyan"), (4.2, "EVAL", "did it work?", "lime"), (7.4, "RED TEAM", "how could it fail?", "coral"), (10.0, "RELEASE", "should it ship?", "amber")]:
        node(slide, x, 5.0, 2.1, 0.85, title, [body], accent, "GATE")
    t(slide, 1.0, 1.95, 10.8, 0.3, "OBSERVABILITY IS NOT LOGGING. IT IS THE EVIDENCE LAYER FOR SYSTEM BEHAVIOR.", 11, "lime", True, "Consolas", "center")
    footer(slide, "QUALITY / TRACE THE LOOP, EVALUATE THE BEHAVIOR, RED-TEAM THE BOUNDARY")

    # 12 — close
    slide = base(prs, 12, "11 / CLOSE", "Build the smallest system that can carry the work.", "Complexity is not the proof of intelligence. Fit is.")
    t(slide, 0.85, 1.95, 4.55, 1.25, "Start small.\nEarn autonomy.", 32, "ink", True)
    line(slide, 0.88, 3.55, 1.6, 0.06, "lime")
    ml(slide, 0.9, 3.95, 4.5, 1.0, ["Choose the path.", "Expose the boundary.", "Close the loop."], 16, "muted", 1.25)
    # compact decision table
    rrect(slide, 6.0, 1.8, 6.25, 3.7, fill=C["panel"], line=C["line"], C=C)
    t(slide, 6.3, 2.1, 2.2, 0.22, "IF THE WORK IS…", 10, "muted", True, "Consolas")
    t(slide, 9.15, 2.1, 2.3, 0.22, "START WITH…", 10, "muted", True, "Consolas")
    rows = [("predictable", "workflow", "cyan"), ("branching", "router", "amber"), ("parallel", "fan-out", "lime"), ("open-ended", "agent loop", "cyan"), ("high-stakes", "human gate", "coral")]
    for i, (condition, choice, accent) in enumerate(rows):
        y = 2.55 + i * 0.55
        line(slide, 6.3, y - 0.12, 5.35, 0.012, "grid")
        t(slide, 6.3, y, 2.35, 0.2, condition, 12, "ink", False, "Consolas")
        t(slide, 9.15, y, 2.4, 0.2, choice, 12, accent, True, "Consolas")
    t(slide, 0.88, 6.55, 11.2, 0.28, "THE OPERATING SYSTEM IS THE DECISION SURFACE BETWEEN INTENT AND ACTION.", 12, "cyan", True, "Consolas", "center")
    footer(slide, "AI AGENT OPERATING SYSTEM / END OF BLUEPRINT")

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUTPUT))
    print(f"Saved: {OUTPUT}")


if __name__ == "__main__":
    main()
