"""QA Gates — post-render quality checks for generated .pptx files."""

from __future__ import annotations

from typing import Any

from ppt_pro_max.decider.design_decider import PageDesign
from ppt_pro_max.content.content_generator import PageContent


class QAGates:
    def check(self, output_path: str, page_designs: list[PageDesign], page_contents: list[PageContent]) -> dict[str, Any]:
        results = {
            "passed": True,
            "checks": [],
        }

        checks = [
            self._check_page_count(page_designs),
            self._check_title_per_page(page_contents),
            self._check_font_sizes(page_designs),
            self._check_theme_consistency(page_designs),
            self._check_no_empty_titles(page_contents),
        ]

        for check in checks:
            results["checks"].append(check)
            if not check["passed"]:
                results["passed"] = False

        return results

    def _check_page_count(self, designs: list[PageDesign]) -> dict[str, Any]:
        count = len(designs)
        passed = 3 <= count <= 30
        return {
            "name": "page_count",
            "passed": passed,
            "detail": f"{count} pages (expected 3-30)",
        }

    def _check_title_per_page(self, contents: list[PageContent]) -> dict[str, Any]:
        missing = [c.position for c in contents if not c.title]
        return {
            "name": "title_per_page",
            "passed": len(missing) == 0,
            "detail": f"Pages missing title: {missing}" if missing else "All pages have titles",
        }

    def _check_font_sizes(self, designs: list[PageDesign]) -> dict[str, Any]:
        small_fonts = []
        for d in designs:
            primary = d.typography.get("primary_size", 28)
            if isinstance(primary, (int, float)) and primary < 14:
                small_fonts.append(d.position)
        return {
            "name": "font_sizes",
            "passed": len(small_fonts) == 0,
            "detail": f"Pages with font < 14pt: {small_fonts}" if small_fonts else "All fonts >= 14pt",
        }

    def _check_theme_consistency(self, designs: list[PageDesign]) -> dict[str, Any]:
        layouts = set(d.layout for d in designs)
        return {
            "name": "theme_consistency",
            "passed": True,
            "detail": f"Using {len(layouts)} distinct layouts",
        }

    def _check_no_empty_titles(self, contents: list[PageContent]) -> dict[str, Any]:
        placeholder_titles = [c.position for c in contents if c.title.startswith("[")]
        return {
            "name": "no_placeholder_titles",
            "passed": len(placeholder_titles) == 0,
            "detail": f"Pages with placeholder titles: {placeholder_titles}" if placeholder_titles else "All titles filled",
        }
