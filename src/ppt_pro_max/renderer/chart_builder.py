"""Chart Builder — python-pptx chart generation with theme colors."""

from __future__ import annotations

from typing import Any

try:
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.dml.color import RGBColor

    _PPTX_AVAILABLE = True
except ImportError:
    _PPTX_AVAILABLE = False


_CHART_TYPE_MAP: dict[str, Any] = {}
if _PPTX_AVAILABLE:
    _CHART_TYPE_MAP = {
        "Bar Chart Vertical": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "Bar Chart Horizontal": XL_CHART_TYPE.BAR_CLUSTERED,
        "Line Chart": XL_CHART_TYPE.LINE,
        "Area Chart": XL_CHART_TYPE.AREA,
        "Pie Chart": XL_CHART_TYPE.PIE,
        "Doughnut Chart": XL_CHART_TYPE.DOUGHNUT,
        "Scatter Plot": XL_CHART_TYPE.XY_SCATTER,
        "Radar Chart": XL_CHART_TYPE.RADAR_FILLED,
    }


class ChartBuilder:
    def build(self, slide, chart_type: str, data: dict[str, Any], style: dict[str, Any], position: dict[str, float]) -> Any:
        if not _PPTX_AVAILABLE:
            return None

        pptx_type = _CHART_TYPE_MAP.get(chart_type)
        if pptx_type is None:
            return None

        chart_data = CategoryChartData()
        chart_data.categories = data.get("labels", ["Q1", "Q2", "Q3", "Q4"])

        values = data.get("values", [10, 25, 45, 80])
        if isinstance(values, list) and isinstance(values[0], list):
            for i, series in enumerate(values):
                chart_data.add_series(f"Series {i + 1}", series)
        else:
            chart_data.add_series("Data", values)

        try:
            from pptx.util import Inches

            chart_frame = slide.shapes.add_chart(
                pptx_type,
                Inches(position.get("x", 1.5)),
                Inches(position.get("y", 1.5)),
                Inches(position.get("width", 10.333)),
                Inches(position.get("height", 4.5)),
                chart_data,
            )

            chart = chart_frame.chart
            chart.has_legend = False

            plot = chart.plots[0]
            colors = style.get("colors", {})
            primary_color = colors.get("primary", "#2563EB")

            try:
                series = plot.series[0]
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = RGBColor.from_string(primary_color.lstrip("#"))
            except Exception:
                pass

            return chart_frame
        except Exception:
            return None
