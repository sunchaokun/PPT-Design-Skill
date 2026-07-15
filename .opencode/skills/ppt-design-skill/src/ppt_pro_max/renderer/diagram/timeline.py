"""TimelineDiagram — horizontal axis with alternating labels."""

from __future__ import annotations

from ppt_pro_max.renderer.diagram.base import BaseDiagram


class TimelineDiagram(BaseDiagram):

    def compute_layout(self) -> None:
        events = self.data.get("events", [])
        if not events:
            return

        n = len(events)
        spacing = self.region.width / (n + 1)
        axis_y = self.region.center_y

        dot_radius = 0.08
        for i, event in enumerate(events):
            cx = self.region.left + spacing * (i + 1)

            self._nodes.append({
                "x": cx - dot_radius,
                "y": axis_y - dot_radius,
                "width": dot_radius * 2,
                "height": dot_radius * 2,
                "label": "",
                "shape": "oval",
                "fill_role": self.style.node_fill,
                "font_color_role": self.style.node_font_color,
            })

            date_label = event.get("date", event.get("year", ""))
            desc_label = event.get("label", event.get("title", ""))

            above = (i % 2 == 0)
            label_x = cx - spacing * 0.4
            label_w = spacing * 0.8

            if above:
                label_y = axis_y - 0.3 - 0.6
            else:
                label_y = axis_y + 0.3

            text_parts = []
            if date_label:
                text_parts.append(date_label)
            if desc_label:
                text_parts.append(desc_label)
            label_text = "\n".join(text_parts) if text_parts else f"Event {i + 1}"

            self._nodes.append({
                "x": label_x,
                "y": label_y,
                "width": label_w,
                "height": 0.6,
                "label": label_text,
                "shape": "rectangle",
                "fill_role": "background",
                "font_color_role": "foreground",
                "font_size_pt": self.style.label_font_size_pt,
            })

            if i < n - 1:
                next_cx = self.region.left + spacing * (i + 2)
                self._connectors.append({
                    "x1": cx + dot_radius,
                    "y1": axis_y,
                    "x2": next_cx - dot_radius,
                    "y2": axis_y,
                })
