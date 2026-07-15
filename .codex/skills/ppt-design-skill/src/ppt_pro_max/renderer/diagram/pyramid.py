from __future__ import annotations

from ppt_pro_max.renderer.diagram.base import BaseDiagram


class PyramidDiagram(BaseDiagram):

    def compute_layout(self) -> None:
        levels = self.data.get("levels", self.data.get("stages", []))
        if not levels:
            return

        n = len(levels)
        gap = min(self.style.node_gap_inches * 0.3, self.region.height * 0.02)
        level_h = (self.region.height - gap * (n - 1)) / n
        level_h = min(level_h, 1.2)

        for i, level in enumerate(levels):
            width_pct = 1.0 - (i / max(n, 1)) * 0.8
            level_w = self.region.width * width_pct
            x = self.region.center_x - level_w / 2
            y = self.region.top + i * (level_h + gap)

            label = level.get("label", level.get("name", f"Level {i + 1}"))

            self._nodes.append({
                "x": x, "y": y, "width": level_w, "height": level_h,
                "label": label,
                "shape": "rectangle",
                "fill_role": level.get("fill_role", self.style.node_fill),
                "font_color_role": level.get("font_color_role", self.style.node_font_color),
            })

            if i < n - 1:
                self._connectors.append({
                    "x1": self.region.center_x,
                    "y1": y + level_h,
                    "x2": self.region.center_x,
                    "y2": y + level_h + gap,
                })
