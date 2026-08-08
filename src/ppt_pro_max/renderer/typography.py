"""Typography Scale System — 5-level type hierarchy with density/mode variants."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class TypeScale:
    display: int = 52
    title: int = 28
    subtitle: int = 20
    body: int = 14
    caption: int = 11

    display_lh: float = 1.1
    title_lh: float = 1.2
    subtitle_lh: float = 1.3
    body_lh: float = 1.5
    caption_lh: float = 1.4

    display_tracking: float = -0.02
    title_tracking: float = -0.01
    subtitle_tracking: float = 0.0
    body_tracking: float = 0.0
    caption_tracking: float = 0.02

    @classmethod
    def for_density(cls, density: int) -> TypeScale:
        factor = 1.0 - (density - 1) * 0.025
        return cls(
            display=max(36, int(52 * factor)),
            title=max(20, int(28 * factor)),
            subtitle=max(16, int(20 * factor)),
            body=max(11, int(14 * factor)),
            caption=max(9, int(11 * factor)),
        )

    @classmethod
    def for_mode(cls, mode: str) -> TypeScale:
        if mode == "presenting":
            return cls(display=64, title=32, subtitle=22, body=18, caption=12)
        elif mode == "reading":
            return cls(display=44, title=24, subtitle=18, body=14, caption=11)
        else:
            return cls()
