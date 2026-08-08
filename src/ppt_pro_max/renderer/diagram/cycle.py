from __future__ import annotations

import math

from ppt_pro_max.renderer.diagram.base import BaseDiagram


class CycleDiagram(BaseDiagram):

    def compute_layout(self) -> None:
        nodes_data = self.data.get("nodes", self.data.get("stages", []))
        if not nodes_data:
            return

        n = len(nodes_data)
        if n < 2:
            return

        cx = self.region.center_x
        cy = self.region.center_y
        radius = min(self.region.width, self.region.height) * 0.35

        node_w = min(self.region.width * 0.25, 2.5)
        node_h = min(self.region.height * 0.2, 1.0)

        for i, nd in enumerate(nodes_data):
            angle = 2 * math.pi * i / n - math.pi / 2
            nx = cx + radius * math.cos(angle) - node_w / 2
            ny = cy + radius * math.sin(angle) - node_h / 2

            label = nd.get("label", nd.get("name", f"Step {i + 1}"))

            self._nodes.append({
                "x": nx, "y": ny, "width": node_w, "height": node_h,
                "label": label,
                "fill_role": nd.get("fill_role", self.style.node_fill),
                "font_color_role": nd.get("font_color_role", self.style.node_font_color),
            })

            next_i = (i + 1) % n
            next_angle = 2 * math.pi * next_i / n - math.pi / 2

            x1 = cx + radius * math.cos(angle)
            y1 = cy + radius * math.sin(angle)
            x2 = cx + radius * math.cos(next_angle)
            y2 = cy + radius * math.sin(next_angle)

            self._connectors.append({
                "x1": x1, "y1": y1, "x2": x2, "y2": y2,
            })
