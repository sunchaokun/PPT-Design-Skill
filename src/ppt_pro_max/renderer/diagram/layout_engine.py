"""Region — rectangular drawing area for diagrams."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Region:
    left: float
    top: float
    width: float
    height: float

    @property
    def center_x(self) -> float:
        return self.left + self.width / 2

    @property
    def center_y(self) -> float:
        return self.top + self.height / 2

    @property
    def right(self) -> float:
        return self.left + self.width

    @property
    def bottom(self) -> float:
        return self.top + self.height

    def subregion(self, left_pct: float, top_pct: float, width_pct: float, height_pct: float) -> Region:
        return Region(
            left=self.left + self.width * left_pct,
            top=self.top + self.height * top_pct,
            width=self.width * width_pct,
            height=self.height * height_pct,
        )

    def inset(self, margin: float) -> Region:
        return Region(
            left=self.left + margin,
            top=self.top + margin,
            width=max(0, self.width - 2 * margin),
            height=max(0, self.height - 2 * margin),
        )
