from __future__ import annotations

from ppt_pro_max.renderer.diagram.base import BaseDiagram


class MatrixDiagram(BaseDiagram):

    def compute_layout(self) -> None:
        rows = self.data.get("rows", [])
        cols = self.data.get("cols", [])
        cells = self.data.get("cells", [])

        if not rows or not cols:
            return

        n_rows = len(rows)
        n_cols = len(cols)

        header_h = 0.5
        label_w = 1.5
        gap = 0.05

        cell_w = (self.region.width - label_w - gap * n_cols) / n_cols
        cell_h = (self.region.height - header_h - gap * n_rows) / n_rows

        for j, col in enumerate(cols):
            cx = self.region.left + label_w + j * (cell_w + gap)
            self._nodes.append({
                "x": cx, "y": self.region.top,
                "width": cell_w, "height": header_h,
                "label": col.get("label", col.get("name", f"Col {j + 1}")),
                "shape": "rectangle",
                "fill_role": "primary",
                "font_color_role": "on-primary",
                "font_size_pt": self.style.cell_header_font_size_pt,
                "font_weight": "bold",
            })

        for i, row in enumerate(rows):
            ry = self.region.top + header_h + i * (cell_h + gap)
            self._nodes.append({
                "x": self.region.left, "y": ry,
                "width": label_w, "height": cell_h,
                "label": row.get("label", row.get("name", f"Row {i + 1}")),
                "shape": "rectangle",
                "fill_role": "secondary",
                "font_color_role": "on-primary",
                "font_size_pt": self.style.cell_header_font_size_pt,
                "font_weight": "bold",
            })

            for j in range(n_cols):
                cx = self.region.left + label_w + j * (cell_w + gap)
                cell_value = ""
                if i < len(cells) and j < len(cells[i]):
                    cell_value = str(cells[i][j]) if cells[i][j] is not None else ""

                self._nodes.append({
                    "x": cx, "y": ry,
                    "width": cell_w, "height": cell_h,
                    "label": cell_value,
                    "shape": "rectangle",
                    "fill_role": "background",
                    "font_color_role": "foreground",
                    "font_size_pt": self.style.cell_body_font_size_pt,
                })
