from __future__ import annotations

import math

from ppt_pro_max.renderer.diagram.base import BaseDiagram


class VennDiagram(BaseDiagram):

    def compute_layout(self) -> None:
        sets = self.data.get("sets", self.data.get("nodes", []))
        if not sets:
            return

        n = len(sets)
        if n < 2 or n > 3:
            return

        cx = self.region.center_x
        cy = self.region.center_y
        radius = min(self.region.width, self.region.height) * 0.25

        if n == 2:
            offset = radius * 0.6
            positions = [
                (cx - offset, cy),
                (cx + offset, cy),
            ]
        else:
            positions = []
            for i in range(3):
                angle = 2 * math.pi * i / 3 - math.pi / 2
                offset = radius * 0.5
                positions.append((
                    cx + offset * math.cos(angle),
                    cy + offset * math.sin(angle),
                ))

        for i, (sx, sy) in enumerate(positions):
            label = sets[i].get("label", sets[i].get("name", f"Set {i + 1}"))

            self._nodes.append({
                "x": sx - radius, "y": sy - radius,
                "width": radius * 2, "height": radius * 2,
                "label": label,
                "shape": "oval",
                "fill_role": sets[i].get("fill_role", self.style.node_fill),
                "font_color_role": sets[i].get("font_color_role", self.style.node_font_color),
            })
