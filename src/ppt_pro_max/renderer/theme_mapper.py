"""Theme Mapper — ui-ux-pro-max design system → PPT theme."""

from __future__ import annotations

from typing import Any

try:
    from pptx import Presentation
    from pptx.oxml.ns import qn
    from pptx.dml.color import RGBColor

    _PPTX_AVAILABLE = True
except ImportError:
    _PPTX_AVAILABLE = False


_COLOR_ROLE_MAP = {
    "primary": "accent1",
    "on-primary": "accent1_light",
    "secondary": "accent2",
    "accent": "accent3",
    "background": "bg1",
    "foreground": "dk1",
    "muted": "bg2",
    "muted-foreground": "dk2",
    "border": "dk2_light",
    "destructive": "accent6",
    "ring": "accent4",
}


_PRESET_THEMES: dict[str, dict[str, Any]] = {
    "professional": {
        "name": "Professional Modern",
        "colors": {
            "primary": "#2563EB",
            "on-primary": "#FFFFFF",
            "secondary": "#64748B",
            "accent": "#F97316",
            "background": "#FFFFFF",
            "foreground": "#1E293B",
            "muted": "#F1F5F9",
            "muted-foreground": "#94A3B8",
            "border": "#E2E8F0",
            "destructive": "#EF4444",
        },
        "typography": {"heading": "Inter", "body": "Inter"},
        "dark_mode": False,
    },
    "dark-tech": {
        "name": "Dark Tech",
        "colors": {
            "primary": "#6366F1",
            "on-primary": "#FFFFFF",
            "secondary": "#94A3B8",
            "accent": "#22D3EE",
            "background": "#0F172A",
            "foreground": "#F8FAFC",
            "muted": "#1E293B",
            "muted-foreground": "#64748B",
            "border": "#334155",
            "destructive": "#EF4444",
        },
        "typography": {"heading": "Space Grotesk", "body": "Inter"},
        "dark_mode": True,
    },
    "warm-elegant": {
        "name": "Warm Elegant",
        "colors": {
            "primary": "#92400E",
            "on-primary": "#FFFFFF",
            "secondary": "#78716C",
            "accent": "#D97706",
            "background": "#FFFBEB",
            "foreground": "#1C1917",
            "muted": "#FEF3C7",
            "muted-foreground": "#A8A29E",
            "border": "#E7E5E4",
            "destructive": "#DC2626",
        },
        "typography": {"heading": "Playfair Display", "body": "Inter"},
        "dark_mode": False,
    },
    "vibrant-startup": {
        "name": "Vibrant Startup",
        "colors": {
            "primary": "#7C3AED",
            "on-primary": "#FFFFFF",
            "secondary": "#64748B",
            "accent": "#EC4899",
            "background": "#F8FAFC",
            "foreground": "#0F172A",
            "muted": "#F1F5F9",
            "muted-foreground": "#94A3B8",
            "border": "#E2E8F0",
            "destructive": "#EF4444",
        },
        "typography": {"heading": "Poppins", "body": "Inter"},
        "dark_mode": False,
    },
    "nature-calm": {
        "name": "Nature Calm",
        "colors": {
            "primary": "#059669",
            "on-primary": "#FFFFFF",
            "secondary": "#64748B",
            "accent": "#34D399",
            "background": "#ECFDF5",
            "foreground": "#064E3B",
            "muted": "#D1FAE5",
            "muted-foreground": "#6B7280",
            "border": "#A7F3D0",
            "destructive": "#DC2626",
        },
        "typography": {"heading": "Lora", "body": "Inter"},
        "dark_mode": False,
    },
}


class ThemeMapper:
    def map(self, design_system: dict[str, Any], theme_name: str | None = None) -> dict[str, Any]:
        if theme_name and theme_name in _PRESET_THEMES:
            return _PRESET_THEMES[theme_name]

        colors = design_system.get("colors", {})
        typography = design_system.get("typography", {})

        return {
            "name": theme_name or "Custom",
            "colors": colors,
            "typography": typography,
            "dark_mode": self._is_dark_mode(colors),
        }

    def apply_theme(self, prs: Presentation, theme: dict[str, Any]) -> None:
        if not _PPTX_AVAILABLE:
            return

        colors = theme.get("colors", {})
        typography = theme.get("typography", {})

        self._apply_default_fonts(prs, typography)
        self._apply_theme_colors(prs, colors)

    def _apply_default_fonts(self, prs: Presentation, typography: dict[str, str]) -> None:
        if not _PPTX_AVAILABLE:
            return

        heading_font = typography.get("heading", "Inter")
        body_font = typography.get("body", "Inter")

        try:
            for master in prs.slide_masters:
                theme_elem = master.element.find(qn("p:cSld"))
                if theme_elem is None:
                    continue
        except Exception:
            pass

    def _apply_theme_colors(self, prs: Presentation, colors: dict[str, str]) -> None:
        if not _PPTX_AVAILABLE:
            return
        try:
            for master in prs.slide_masters:
                for layout in master.slide_layouts:
                    pass
        except Exception:
            pass

    def _is_dark_mode(self, colors: dict[str, str]) -> bool:
        bg = colors.get("background", "#FFFFFF")
        r, g, b = int(bg[1:3], 16), int(bg[3:5], 16), int(bg[5:7], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        return luminance < 0.5

    @staticmethod
    def get_preset_themes() -> dict[str, dict[str, Any]]:
        return _PRESET_THEMES.copy()
