from __future__ import annotations

from ppt_pro_max.renderer.diagram.base import BaseDiagram


class SwotDiagram(BaseDiagram):

    _QUADRANT_COLORS = {
        "strengths": "primary",
        "weaknesses": "secondary",
        "opportunities": "accent",
        "threats": "muted-foreground",
    }

    def compute_layout(self) -> None:
        quadrants = self.data.get("quadrants", self.data.get("nodes", []))
        if not quadrants:
            quadrants = [
                {"label": "Strengths", "items": []},
                {"label": "Weaknesses", "items": []},
                {"label": "Opportunities", "items": []},
                {"label": "Threats", "items": []},
            ]

        n = len(quadrants)
        if n < 4:
            while len(quadrants) < 4:
                quadrants.append({"label": f"Quadrant {len(quadrants) + 1}", "items": []})

        gap = 0.15
        half_w = (self.region.width - gap) / 2
        half_h = (self.region.height - gap) / 2

        positions = [
            (self.region.left, self.region.top),
            (self.region.left + half_w + gap, self.region.top),
            (self.region.left, self.region.top + half_h + gap),
            (self.region.left + half_w + gap, self.region.top + half_h + gap),
        ]

        keys = ["strengths", "weaknesses", "opportunities", "threats"]

        for i, (qx, qy) in enumerate(positions):
            q = quadrants[i]
            label = q.get("label", keys[i].title())
            fill_role = q.get("fill_role", self._QUADRANT_COLORS.get(keys[i], self.style.node_fill))

            self._nodes.append({
                "x": qx, "y": qy, "width": half_w, "height": half_h,
                "label": label,
                "shape": "rectangle",
                "fill_role": fill_role,
                "font_color_role": q.get("font_color_role", self.style.node_font_color),
                "font_size_pt": self.style.cell_header_font_size_pt,
                "font_weight": "bold",
            })

            items = q.get("items", [])
            if items:
                item_text = "\n".join(f"• {item}" if not item.startswith(("• ", "- ")) else item for item in items[:8])
                self._nodes.append({
                    "x": qx + 0.1, "y": qy + 0.4, "width": half_w - 0.2, "height": half_h - 0.5,
                    "label": item_text,
                    "shape": "rectangle",
                    "fill_role": "background",
                    "font_color_role": "foreground",
                    "font_size_pt": self.style.cell_body_font_size_pt,
                })
