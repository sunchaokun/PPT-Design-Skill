"""Generate PPTX artifacts that exercise the documented theme contracts."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pptx_designer
from pptx_designer import Presentation, generate_ppt, merge_vi_design_context, validate_resolved_theme
from pptx_designer.renderer.theme import ThemeComposer
from pptx_designer.tools.shapes import rect, rrect
from pptx_designer.tools.text import text


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(exist_ok=True)


def write_theme(theme: dict) -> str:
    payload = json.dumps(theme, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    (ROOT / "resolved-theme-v1.json").write_text(payload + "\n", encoding="utf-8")
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def build_deck(theme: dict) -> Path:
    prs = Presentation(theme=theme, strict_theme=True)

    cover = prs.slides.add_slide(prs.slide_layouts[6])
    rect(cover, 0, 0, 13.333, 7.5, "background")
    rect(cover, 0.7, 0.8, 0.08, 5.9, "accent")
    text(cover, 1.1, 1.0, 10.8, 1.15, "THEME CONTRACT\nSMOKE TEST", font_size=32, bold=True, color="ink")
    text(
        cover, 1.12, 2.45, 9.8, 0.7,
        "One resolved theme drives editable output across generation modes.",
        font_size=18, color="muted",
    )
    labels = [("VALIDATED", "resolved theme"), ("EDITABLE", "native objects"), ("RENDERED", "PNG reviewed")]
    for index, (headline, detail) in enumerate(labels):
        x = 1.12 + index * 3.7
        rrect(cover, x, 4.35, 3.15, 1.2, fill="surface")
        text(cover, x + 0.25, 4.6, 2.7, 0.28, headline, font_size=13, bold=True, color="accent")
        text(cover, x + 0.25, 4.98, 2.7, 0.26, detail, font_size=12, color="ink")

    flow = prs.slides.add_slide(prs.slide_layouts[6])
    rect(flow, 0, 0, 13.333, 7.5, "background")
    text(flow, 0.85, 0.72, 11.4, 0.6, "ONE THEME, THREE DELIVERY STATES", font_size=24, bold=True, color="ink")
    text(flow, 0.87, 1.42, 9.4, 0.42, "The design record guides decisions; the resolved object drives rendering.", font_size=14, color="muted")
    stages = [("01", "THEME LOCK", "intent + confirmation"), ("02", "RESOLVED THEME", "validated API object"), ("03", "EDITABLE PPTX", "native text + shapes")]
    for index, (number, title, body) in enumerate(stages):
        x = 0.9 + index * 4.15
        rrect(flow, x, 2.55, 3.35, 2.25, fill="surface")
        text(flow, x + 0.28, 2.88, 0.58, 0.35, number, font_size=15, bold=True, color="accent")
        text(flow, x + 0.28, 3.48, 2.7, 0.38, title, font_size=16, bold=True, color="ink")
        text(flow, x + 0.28, 4.04, 2.65, 0.35, body, font_size=12, color="muted")
        if index < 2:
            rect(flow, x + 3.46, 3.54, 0.48, 0.05, "accent")

    output = OUTPUT / "build-theme-contract.pptx"
    prs.save(output)
    return output


def freestyle_deck(theme: dict) -> tuple[Path, dict]:
    output = OUTPUT / "freestyle-theme-contract.pptx"
    result = generate_ppt(
        content={
            "pages": [
                {"goal": "hook", "title": "Theme-guided delivery", "subtitle": "Resolved theme contract verified"},
                {"goal": "data", "title": "Validation evidence", "subtitle": "Three checks completed", "bullets": ["Theme: validated", "Build: editable", "PNG: reviewed"]},
            ]
        },
        theme=theme,
        output=str(output),
    )
    return output, result


def vi_contract(theme: dict) -> dict:
    template = {
        "assets": {"logo": {"path": "brand.svg"}},
        "locks": [{"field": "assets.logo", "mode": "template-locked"}],
    }
    allowed = merge_vi_design_context(template, theme, {"page_role": "content"})
    rejected = merge_vi_design_context(template, theme, {"assets": {"logo": {"path": "other.svg"}}})
    assert not allowed["diagnostics"]["conflicts"]
    assert rejected["assets"]["logo"] == template["assets"]["logo"]
    assert rejected["diagnostics"]["conflicts"]
    return {"allowed_conflicts": allowed["diagnostics"]["conflicts"], "rejected_conflicts": rejected["diagnostics"]["conflicts"]}


def main() -> None:
    theme = ThemeComposer().compose(style="dark-tech", seed=17)
    validate_resolved_theme(theme)
    fingerprint = write_theme(theme)
    build_path = build_deck(theme)
    freestyle_path, freestyle_result = freestyle_deck(theme)
    results = {
        "pptx_designer_version": pptx_designer.__version__,
        "pptx_designer_module_path": pptx_designer.__file__,
        "theme_fingerprint": fingerprint,
        "build_output": str(build_path),
        "freestyle_output": str(freestyle_path),
        "freestyle_theme_application": freestyle_result["theme_application"],
        "vi_contract": vi_contract(theme),
    }
    (OUTPUT / "contract-results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
