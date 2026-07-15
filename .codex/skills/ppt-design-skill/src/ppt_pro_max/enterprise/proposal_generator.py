"""ProposalGenerator — generate 2-3 style preview PPTs for user selection.

Flow:
  1. Get 2-3 sets of design atoms from ThemeComposer
  2. For each set, generate a 4-page preview PPT (hook, problem, features, cta)
  3. Save as proposal_A.pptx, proposal_B.pptx, proposal_C.pptx
  4. Return proposal descriptions for user selection

Differentiation:
  - Proposal A: style closest to user description
  - Proposal B: same mood but different palette
  - Proposal C: different mood alternative
"""

from __future__ import annotations

import os
from typing import Any

from ppt_pro_max.enterprise.brand_spec import BrandSpec
from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
from ppt_pro_max.renderer.theme_composer import COLOR_PALETTES, FONT_PAIRS, ThemeComposer

_MOOD_ALTERNATIVES: dict[str, str] = {
    "professional": "creative",
    "tech": "nature",
    "dark": "warm",
    "warm": "minimal",
    "elegant": "bold",
    "luxury": "fresh",
    "vibrant": "calm",
    "startup": "professional",
    "minimal": "elegant",
    "bold": "elegant",
    "calm": "vibrant",
    "fresh": "luxury",
    "nature": "tech",
    "creative": "professional",
    "corporate": "startup",
    "industrial": "fresh",
    "fintech": "creative",
    "health": "bold",
    "education": "startup",
    "sustainability": "tech",
    "international": "warm",
    "cream": "frosted",
    "frosted": "cream",
    "mckinsey": "creative",
    "consulting": "startup",
    "pastel": "bold",
    "retro": "modern",
    "government": "creative",
    "legal": "warm",
    "pharma": "fresh",
    "realestate": "nature",
    "automotive": "elegant",
    "aviation": "bold",
    "energy": "nature",
    "telecom": "creative",
    "logistics": "fresh",
}

_PREVIEW_PAGES: list[dict[str, Any]] = [
    {
        "goal": "hook",
        "title": "{query}",
        "subtitle": "Proposal preview",
    },
    {
        "goal": "content",
        "title": "Current Challenges",
        "bullets": ["Challenge 1", "Challenge 2", "Challenge 3"],
    },
    {
        "goal": "features",
        "title": "Key Features",
        "cards": [
            {"title": "Feature 1", "body": "Description"},
            {"title": "Feature 2", "body": "Description"},
            {"title": "Feature 3", "body": "Description"},
        ],
    },
    {
        "goal": "cta",
        "title": "Get Started",
        "subtitle": "Contact us today",
    },
]


class ProposalGenerator:

    def generate(
        self,
        query: str,
        style: str | None = None,
        output_dir: str = ".",
        project_dir: str | None = None,
        seed: int | None = None,
    ) -> list[dict[str, Any]]:
        composer = ThemeComposer()

        style_text = style or query
        theme_a = composer.compose(style=style_text, seed=seed)
        atoms_a = theme_a["atoms"]

        moods = atoms_a.get("moods", ["professional"])
        primary_mood = moods[0] if moods else "professional"

        theme_b = composer.compose(style=style_text, mood=primary_mood, seed=(seed + 1) if seed is not None else None)
        atoms_b = theme_b["atoms"]

        if atoms_b["palette"] == atoms_a["palette"]:
            all_palettes = list(COLOR_PALETTES.keys())
            mood_palettes = all_palettes
            for p in mood_palettes:
                if p != atoms_a["palette"]:
                    theme_b = composer.compose(style=style_text, palette=p, seed=(seed + 1) if seed is not None else None)
                    atoms_b = theme_b["atoms"]
                    break

        alt_mood = _MOOD_ALTERNATIVES.get(primary_mood, "creative")
        theme_c = composer.compose(mood=alt_mood, seed=(seed + 2) if seed is not None else None)
        atoms_c = theme_c["atoms"]

        if atoms_c["palette"] == atoms_a["palette"] and atoms_c["palette"] == atoms_b["palette"]:
            all_palettes = list(COLOR_PALETTES.keys())
            for p in all_palettes:
                if p not in (atoms_a["palette"], atoms_b["palette"]):
                    theme_c = composer.compose(mood=alt_mood, palette=p, seed=(seed + 2) if seed is not None else None)
                    atoms_c = theme_c["atoms"]
                    break

        proposals: list[dict[str, Any]] = []
        for idx, (label, theme, atoms) in enumerate([
            ("A", theme_a, atoms_a),
            ("B", theme_b, atoms_b),
            ("C", theme_c, atoms_c),
        ]):
            filename = f"proposal_{label}.pptx"
            filepath = os.path.join(output_dir, filename)

            brand_spec = self._make_brand_spec(atoms)
            renderer = PrecisionRenderer(brand_spec=brand_spec, template_path=self._find_template(project_dir))
            prs = renderer.create_presentation()

            for page_template in _PREVIEW_PAGES:
                page = self._fill_page(page_template, query)
                renderer.render_slide(prs, page)

            renderer.save(prs, filepath)

            description = self._describe(label, atoms, primary_mood, alt_mood)
            proposals.append({
                "id": label,
                "path": filepath,
                "atoms": atoms,
                "description": description,
            })

        return proposals

    def _make_brand_spec(self, atoms: dict[str, Any]) -> BrandSpec:
        palette_data = COLOR_PALETTES.get(atoms["palette"], COLOR_PALETTES["ocean-blue"])
        font_data = FONT_PAIRS.get(atoms["fonts"], FONT_PAIRS["modern-sans"])
        brand_spec = BrandSpec()
        brand_spec.source = "theme_composer"
        brand_spec.colors = dict(palette_data)
        brand_spec.fonts = dict(font_data)
        return brand_spec

    def _find_template(self, project_dir: str | None) -> str | None:
        if not project_dir:
            return None
        template_path = os.path.join(project_dir, "template.pptx")
        if os.path.isfile(template_path):
            return template_path
        return None

    def _fill_page(self, template: dict[str, Any], query: str) -> dict[str, Any]:
        page = dict(template)
        if page.get("title") == "{query}":
            page["title"] = query
        return page

    def _describe(self, label: str, atoms: dict[str, Any], primary_mood: str, alt_mood: str) -> str:
        palette = atoms.get("palette", "unknown")
        moods = atoms.get("moods", [])
        mood_str = ", ".join(moods) if moods else primary_mood
        if label == "A":
            return f"Closest to your style — {palette} palette, {mood_str} mood"
        elif label == "B":
            return f"Same mood, different palette — {palette} palette, {mood_str} mood"
        else:
            return f"Alternative mood ({alt_mood}) — {palette} palette, {mood_str} mood"
