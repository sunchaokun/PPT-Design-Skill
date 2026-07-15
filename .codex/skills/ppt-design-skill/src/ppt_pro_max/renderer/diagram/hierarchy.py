from __future__ import annotations

from typing import Any

from ppt_pro_max.renderer.diagram.base import BaseDiagram


class HierarchyDiagram(BaseDiagram):

    def compute_layout(self) -> None:
        nodes_data = self.data.get("nodes", [])
        if not nodes_data:
            return

        levels: dict[int, list[dict[str, Any]]] = {}
        for nd in nodes_data:
            level = nd.get("level", 0)
            levels.setdefault(level, []).append(nd)

        if not levels:
            return

        max_level = max(levels.keys())
        gap_v = min(self.style.node_gap_inches + 0.2, self.region.height * 0.04)
        level_h = (self.region.height - gap_v * max_level) / (max_level + 1)
        level_h = min(level_h, 1.2)

        node_positions: dict[str, tuple[float, float, float, float]] = {}

        for lvl in range(max_level + 1):
            items = levels.get(lvl, [])
            n = len(items)
            if n == 0:
                continue

            node_w = min(self.region.width / max(n, 1) - self.style.node_gap_inches, 3.0)
            total_w = node_w * n + self.style.node_gap_inches * (n - 1)
            start_x = self.region.center_x - total_w / 2
            y = self.region.top + lvl * (level_h + gap_v)

            for i, nd in enumerate(items):
                x = start_x + i * (node_w + self.style.node_gap_inches)
                label = nd.get("label", nd.get("name", f"L{lvl}-{i + 1}"))
                node_id = nd.get("id", label)

                self._nodes.append({
                    "x": x, "y": y, "width": node_w, "height": level_h,
                    "label": label,
                    "fill_role": nd.get("fill_role", self.style.node_fill),
                    "font_color_role": nd.get("font_color_role", self.style.node_font_color),
                })
                node_positions[node_id] = (x, y, node_w, level_h)

        for nd in nodes_data:
            parent_id = nd.get("parent")
            child_id = nd.get("id", nd.get("label", ""))
            if parent_id and parent_id in node_positions and child_id in node_positions:
                px, py, pw, ph = node_positions[parent_id]
                cx, cy, cw, ch = node_positions[child_id]
                self._connectors.append({
                    "x1": px + pw / 2, "y1": py + ph,
                    "x2": cx + cw / 2, "y2": cy,
                })
