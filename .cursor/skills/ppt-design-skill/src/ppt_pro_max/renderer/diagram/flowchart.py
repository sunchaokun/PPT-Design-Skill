"""FlowchartDiagram — horizontal/vertical flow diagram."""

from __future__ import annotations

from typing import Any

from ppt_pro_max.renderer.diagram.base import BaseDiagram
from ppt_pro_max.renderer.diagram.text_measurer import estimate_node_size


class FlowchartDiagram(BaseDiagram):

    def compute_layout(self) -> None:
        nodes_data = self.data.get("nodes", [])
        if not nodes_data:
            return

        n = len(nodes_data)
        direction = self.data.get("direction")
        if direction is None:
            direction = "horizontal" if n <= 6 else "vertical"

        if direction == "horizontal":
            self._layout_horizontal(nodes_data)
        else:
            self._layout_vertical(nodes_data)

    def _layout_horizontal(self, nodes_data: list[dict[str, Any]]) -> None:
        n = len(nodes_data)
        gap = min(self.style.node_gap_inches, self.region.width * 0.03)
        node_w = (self.region.width - (n - 1) * gap) / n
        node_w = min(node_w, 3.0)
        node_h = min(self.region.height * 0.35, 1.2)

        max_label_w = node_w - self.style.padding_inches * 2
        for i, nd in enumerate(nodes_data):
            label = nd.get("label", nd.get("title", f"Step {i + 1}"))
            est_w, est_h = estimate_node_size(
                label, self.style.node_font_size_pt, max_label_w,
            )
            actual_h = max(node_h, est_h)

            x = self.region.left + i * (node_w + gap)
            y = self.region.center_y - actual_h / 2

            self._nodes.append({
                "x": x, "y": y, "width": node_w, "height": actual_h,
                "label": label,
                "fill_role": nd.get("fill_role", self.style.node_fill),
                "font_color_role": nd.get("font_color_role", self.style.node_font_color),
            })

            if i < n - 1:
                self._connectors.append({
                    "x1": x + node_w,
                    "y1": y + actual_h / 2,
                    "x2": x + node_w + gap,
                    "y2": y + actual_h / 2,
                })

    def _layout_vertical(self, nodes_data: list[dict[str, Any]]) -> None:
        n = len(nodes_data)
        gap = min(self.style.node_gap_inches + 0.1, self.region.height * 0.03)
        node_h = (self.region.height - (n - 1) * gap) / n
        node_h = min(node_h, 1.0)
        node_w = min(self.region.width * 0.6, 4.0)

        max_label_w = node_w - self.style.padding_inches * 2
        for i, nd in enumerate(nodes_data):
            label = nd.get("label", nd.get("title", f"Step {i + 1}"))
            est_w, est_h = estimate_node_size(
                label, self.style.node_font_size_pt, max_label_w,
            )
            actual_h = max(node_h, est_h)

            x = self.region.center_x - node_w / 2
            y = self.region.top + i * (actual_h + gap)

            self._nodes.append({
                "x": x, "y": y, "width": node_w, "height": actual_h,
                "label": label,
                "fill_role": nd.get("fill_role", self.style.node_fill),
                "font_color_role": nd.get("font_color_role", self.style.node_font_color),
            })

            if i < n - 1:
                self._connectors.append({
                    "x1": x + node_w / 2,
                    "y1": y + actual_h,
                    "x2": x + node_w / 2,
                    "y2": y + actual_h + gap,
                })
