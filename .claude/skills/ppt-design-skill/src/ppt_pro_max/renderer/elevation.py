"""Shadow Elevation System — 5-level depth with dark-mode glow fallback."""

from __future__ import annotations

from ppt_pro_max.renderer.visual_effects import apply_glow, apply_shadow

ELEVATION_SCALE: dict[int, dict[str, int]] = {
    0: {"blur_pt": 0, "distance_pt": 0, "alpha_pct": 0},
    1: {"blur_pt": 3, "distance_pt": 1, "alpha_pct": 10},
    2: {"blur_pt": 6, "distance_pt": 2, "alpha_pct": 15},
    3: {"blur_pt": 12, "distance_pt": 4, "alpha_pct": 20},
    4: {"blur_pt": 24, "distance_pt": 8, "alpha_pct": 25},
}


def apply_elevation(shape, level: int, is_dark: bool = False,
                     foreground_hex: str = "#000000", primary_hex: str = "#2563EB") -> None:
    spec = ELEVATION_SCALE.get(level, ELEVATION_SCALE[1])

    if level == 0:
        return

    if is_dark:
        apply_glow(shape,
                   radius_pt=spec["blur_pt"],
                   color=primary_hex,
                   alpha_pct=max(15, spec["alpha_pct"] // 2))
    else:
        apply_shadow(shape,
                     blur_pt=spec["blur_pt"],
                     distance_pt=spec["distance_pt"],
                     direction_deg=90,
                     color=foreground_hex,
                     alpha_pct=spec["alpha_pct"])
