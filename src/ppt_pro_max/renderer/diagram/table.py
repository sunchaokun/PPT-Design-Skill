from __future__ import annotations

from ppt_pro_max.renderer.diagram.base import BaseDiagram


class TableDiagram(BaseDiagram):

    def compute_layout(self) -> None:
        headers = self.data.get("headers", [])
        rows = self.data.get("rows", [])

        if not headers and not rows:
            return

        if not headers and rows:
            headers = [f"Col {i + 1}" for i in range(len(rows[0]))] if rows else []

        n_cols = len(headers)
        if n_cols == 0:
            return

        header_h = 0.45
        gap = 0.03

        n_rows = len(rows)
        if n_rows == 0:
            row_h = 0.5
        else:
            row_h = (self.region.height - header_h - gap * (n_rows + 1)) / n_rows
            row_h = min(row_h, 0.8)

        col_w = self.region.width / n_cols

        for j, hdr in enumerate(headers):
            self._nodes.append({
                "x": self.region.left + j * col_w,
                "y": self.region.top,
                "width": col_w - gap,
                "height": header_h,
                "label": str(hdr),
                "shape": "rectangle",
                "fill_role": "primary",
                "font_color_role": "on-primary",
                "font_size_pt": self.style.cell_header_font_size_pt,
                "font_weight": "bold",
            })

        for i, row in enumerate(rows):
            ry = self.region.top + header_h + gap + i * (row_h + gap)
            fill = "background" if i % 2 == 0 else "muted"
            for j in range(n_cols):
                val = row[j] if j < len(row) else ""
                self._nodes.append({
                    "x": self.region.left + j * col_w,
                    "y": ry,
                    "width": col_w - gap,
                    "height": row_h,
                    "label": str(val),
                    "shape": "rectangle",
                    "fill_role": fill,
                    "font_color_role": "foreground",
                    "font_size_pt": self.style.cell_body_font_size_pt,
                })
