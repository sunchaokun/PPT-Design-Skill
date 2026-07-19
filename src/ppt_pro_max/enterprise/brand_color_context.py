"""BrandColorContext — semantic color view computed from DesignDNA or BrandSpec."""

from __future__ import annotations

from dataclasses import dataclass


def _luminance(hex_color: str) -> float:
    try:
        h = hex_color.lstrip("#")
        r, g, b = int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255
        return 0.299 * r + 0.587 * g + 0.114 * b
    except Exception:
        return 0.5


def _contrast_text(hex_color: str) -> str:
    return "#FFFFFF" if _luminance(hex_color) < 0.5 else "#333333"


@dataclass
class BrandColorContext:
    primary: str = "#333333"
    accent1: str = "#6096E6"
    accent2: str = "#E8A838"
    muted: str = "#F1F5F9"
    foreground: str = "#333333"
    background: str = "#FFFFFF"
    on_primary: str = "#FFFFFF"

    @classmethod
    def from_dna(cls, dna) -> BrandColorContext:
        palette = getattr(dna, "color_palette", None) or {}
        usage = getattr(dna, "actual_colors", None) or {}

        sorted_colors = sorted(usage.items(), key=lambda x: -x[1])
        dark_colors = [c for c, _ in sorted_colors if _luminance(c) < 0.5]
        light_colors = [c for c, _ in sorted_colors if _luminance(c) >= 0.5]

        primary = dark_colors[0] if dark_colors else palette.get("accent1", "#333333")
        accent1 = palette.get("accent1", dark_colors[1] if len(dark_colors) > 1 else "#6096E6")
        muted = light_colors[0] if light_colors else "#F1F5F9"
        foreground = palette.get("dk1", "#333333")
        background = palette.get("lt1", "#FFFFFF")
        on_primary = _contrast_text(primary)

        return cls(
            primary=primary,
            accent1=accent1,
            accent2=palette.get("accent2", "#E8A838"),
            muted=muted,
            foreground=foreground,
            background=background,
            on_primary=on_primary,
        )

    @classmethod
    def from_brand_spec(cls, brand_spec) -> BrandColorContext:
        colors = getattr(brand_spec, "colors", None) or {}
        accent1 = colors.get("accent1", "#6096E6")
        foreground = colors.get("tx1", "#333333")
        primary = colors.get("primary", foreground)
        muted = colors.get("muted", "#F1F5F9")
        background = colors.get("background", "#FFFFFF")

        return cls(
            primary=primary,
            accent1=accent1,
            accent2=colors.get("accent2", "#E8A838"),
            muted=muted,
            foreground=foreground,
            background=background,
            on_primary=_contrast_text(primary),
        )
