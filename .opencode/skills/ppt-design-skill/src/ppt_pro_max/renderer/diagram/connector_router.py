from __future__ import annotations

from typing import Any


def route_connector(
    x1: float, y1: float, x2: float, y2: float,
    nodes: list[dict[str, Any]],
    padding: float = 0.1,
) -> list[tuple[float, float]]:
    if not _line_intersects_any(x1, y1, x2, y2, nodes, padding):
        return [(x1, y1), (x2, y2)]

    mid_x = (x1 + x2) / 2
    waypoints = [(x1, y1), (mid_x, y1), (mid_x, y2), (x2, y2)]

    if not _path_intersects_any(waypoints, nodes, padding):
        return waypoints

    for offset in [1.0, -1.0, 2.0, -2.0]:
        alt_x = mid_x + offset
        alt_waypoints = [(x1, y1), (alt_x, y1), (alt_x, y2), (x2, y2)]
        if not _path_intersects_any(alt_waypoints, nodes, padding):
            return alt_waypoints

    return waypoints


def _line_intersects_any(
    x1: float, y1: float, x2: float, y2: float,
    nodes: list[dict[str, Any]], padding: float,
) -> bool:
    for node in nodes:
        nx = node["x"] - padding
        ny = node["y"] - padding
        nw = node["width"] + 2 * padding
        nh = node["height"] + 2 * padding
        if _line_intersects_rect(x1, y1, x2, y2, nx, ny, nw, nh):
            if not _is_endpoint_inside(x1, y1, x2, y2, nx, ny, nw, nh):
                return True
    return False


def _line_intersects_rect(
    x1: float, y1: float, x2: float, y2: float,
    rx: float, ry: float, rw: float, rh: float,
) -> bool:
    edges = [
        (rx, ry, rx + rw, ry),
        (rx + rw, ry, rx + rw, ry + rh),
        (rx + rw, ry + rh, rx, ry + rh),
        (rx, ry + rh, rx, ry),
    ]
    for ex1, ey1, ex2, ey2 in edges:
        if _segments_intersect(x1, y1, x2, y2, ex1, ey1, ex2, ey2):
            return True

    if rx <= x1 <= rx + rw and ry <= y1 <= ry + rh:
        return True
    if rx <= x2 <= rx + rw and ry <= y2 <= ry + rh:
        return True

    return False


def _is_endpoint_inside(
    x1: float, y1: float, x2: float, y2: float,
    rx: float, ry: float, rw: float, rh: float,
) -> bool:
    p1_inside = rx <= x1 <= rx + rw and ry <= y1 <= ry + rh
    p2_inside = rx <= x2 <= rx + rw and ry <= y2 <= ry + rh
    return p1_inside or p2_inside


def _path_intersects_any(
    waypoints: list[tuple[float, float]],
    nodes: list[dict[str, Any]], padding: float,
) -> bool:
    for i in range(len(waypoints) - 1):
        x1, y1 = waypoints[i]
        x2, y2 = waypoints[i + 1]
        if _line_intersects_any(x1, y1, x2, y2, nodes, padding):
            return True
    return False


def _segments_intersect(
    ax1: float, ay1: float, ax2: float, ay2: float,
    bx1: float, by1: float, bx2: float, by2: float,
) -> bool:
    def cross(ox, oy, px, py, qx, qy):
        return (px - ox) * (qy - oy) - (py - oy) * (qx - ox)

    d1 = cross(bx1, by1, bx2, by2, ax1, ay1)
    d2 = cross(bx1, by1, bx2, by2, ax2, ay2)
    d3 = cross(ax1, ay1, ax2, ay2, bx1, by1)
    d4 = cross(ax1, ay1, ax2, ay2, bx2, by2)

    if ((d1 > 0 and d2 < 0) or (d1 < 0 and d2 > 0)) and \
       ((d3 > 0 and d4 < 0) or (d3 < 0 and d4 > 0)):
        return True

    if d1 == 0 and _on_segment(bx1, by1, bx2, by2, ax1, ay1):
        return True
    if d2 == 0 and _on_segment(bx1, by1, bx2, by2, ax2, ay2):
        return True
    if d3 == 0 and _on_segment(ax1, ay1, ax2, ay2, bx1, by1):
        return True
    if d4 == 0 and _on_segment(ax1, ay1, ax2, ay2, bx2, by2):
        return True

    return False


def _on_segment(px: float, py: float, qx: float, qy: float, rx: float, ry: float) -> bool:
    return (min(px, qx) <= rx <= max(px, qx) and
            min(py, qy) <= ry <= max(py, qy))
