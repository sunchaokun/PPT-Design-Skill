"""FunnelDiagram — decreasing-width funnel stages."""

from __future__ import annotations

from ppt_pro_max.renderer.diagram.base import BaseDiagram


class FunnelDiagram(BaseDiagram):

    def compute_layout(self) -> None:
        stages = self.data.get("stages", [])
        if not stages:
            return

        n = len(stages)
        gap = min(self.style.node_gap_inches * 0.5, self.region.height * 0.02)
        stage_h = (self.region.height - (n - 1) * gap) / n
        stage_h = min(stage_h, 1.2)

        for i, stage in enumerate(stages):
            width_pct = 1.0 - (i / max(n - 1, 1)) * 0.7
            stage_w = self.region.width * width_pct
            x = self.region.center_x - stage_w / 2
            y = self.region.top + i * (stage_h + gap)

            label = stage.get("label", stage.get("name", f"Stage {i + 1}"))

            self._nodes.append({
                "x": x, "y": y, "width": stage_w, "height": stage_h,
                "label": label,
                "shape": "rectangle",
                "fill_role": stage.get("fill_role", self.style.node_fill),
                "font_color_role": stage.get("font_color_role", self.style.node_font_color),
            })

            if i < n - 1:
                self._connectors.append({
                    "x1": self.region.center_x,
                    "y1": y + stage_h,
                    "x2": self.region.center_x,
                    "y2": y + stage_h + gap,
                })
