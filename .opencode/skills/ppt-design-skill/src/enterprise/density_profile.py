"""DensityProfile — density-driven font/spacing/content adjustments."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class DensityProfile:
    title_size: float
    subtitle_size: float
    body_size: float
    bullet_size: float
    line_spacing: float
    max_bullets: int
    max_bullet_chars: int
    image_width_ratio: float


_DENSITY_PROFILES: dict[int, DensityProfile] = {
    1: DensityProfile(title_size=36, subtitle_size=20, body_size=18, bullet_size=16, line_spacing=1.5, max_bullets=3, max_bullet_chars=40, image_width_ratio=0.5),
    2: DensityProfile(title_size=34, subtitle_size=19, body_size=17, bullet_size=15, line_spacing=1.4, max_bullets=4, max_bullet_chars=50, image_width_ratio=0.45),
    3: DensityProfile(title_size=32, subtitle_size=18, body_size=16, bullet_size=14, line_spacing=1.3, max_bullets=5, max_bullet_chars=60, image_width_ratio=0.4),
    4: DensityProfile(title_size=30, subtitle_size=17, body_size=15, bullet_size=13, line_spacing=1.2, max_bullets=5, max_bullet_chars=70, image_width_ratio=0.38),
    5: DensityProfile(title_size=28, subtitle_size=16, body_size=14, bullet_size=12, line_spacing=1.15, max_bullets=6, max_bullet_chars=80, image_width_ratio=0.35),
    6: DensityProfile(title_size=26, subtitle_size=15, body_size=13, bullet_size=11, line_spacing=1.1, max_bullets=7, max_bullet_chars=90, image_width_ratio=0.32),
    7: DensityProfile(title_size=24, subtitle_size=14, body_size=12, bullet_size=10, line_spacing=1.05, max_bullets=8, max_bullet_chars=100, image_width_ratio=0.3),
    8: DensityProfile(title_size=22, subtitle_size=13, body_size=11, bullet_size=9, line_spacing=1.0, max_bullets=10, max_bullet_chars=120, image_width_ratio=0.25),
    9: DensityProfile(title_size=20, subtitle_size=12, body_size=10, bullet_size=8, line_spacing=0.95, max_bullets=12, max_bullet_chars=140, image_width_ratio=0.22),
    10: DensityProfile(title_size=18, subtitle_size=11, body_size=9, bullet_size=7, line_spacing=0.9, max_bullets=15, max_bullet_chars=160, image_width_ratio=0.2),
}


def get_density_profile(density: int) -> DensityProfile:
    density = max(1, min(10, density))
    return _DENSITY_PROFILES[density]


def apply_density_to_bullets(bullets: list[str], profile: DensityProfile) -> list[str]:
    result: list[str] = []
    for b in bullets[:profile.max_bullets]:
        if len(b) > profile.max_bullet_chars:
            b = b[: profile.max_bullet_chars - 1] + "…"
        result.append(b)
    return result
