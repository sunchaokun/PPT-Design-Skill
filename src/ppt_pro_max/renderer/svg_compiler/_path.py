"""SVG path parser — converts SVG path ``d`` attribute to cubic Bezier segments.

Supports all SVG 1.1 path commands: M/L/H/V/C/S/Q/T/A/Z (absolute + relative).
Arcs are converted to cubic Bezier approximations via ``arc_to_cubics``.
"""
from __future__ import annotations

import math
import re


def arc_to_cubics(
    x0: float,
    y0: float,
    rx: float,
    ry: float,
    rot: float,
    large: int,
    sweep: int,
    x1: float,
    y1: float,
) -> list[tuple[tuple[float, float], tuple[float, float], tuple[float, float], tuple[float, float]]]:
    """Convert an SVG arc to a list of cubic Bezier segments.

    Each segment is ``((p0), (c1), (c2), (p1))``.
    Handles the "full circle" idiom (nearly coincident endpoints) correctly.
    """
    rx, ry = abs(rx), abs(ry)
    if rx < 1e-9 or ry < 1e-9:
        return [((x0, y0), (x0, y0), (x1, y1), (x1, y1))]

    phi = math.radians(rot)
    cp, sp = math.cos(phi), math.sin(phi)
    dx, dy = (x0 - x1) / 2, (y0 - y1) / 2
    x1p, y1p = cp * dx + sp * dy, -sp * dx + cp * dy

    lam = x1p * x1p / (rx * rx) + y1p * y1p / (ry * ry)
    if lam > 1:
        r = math.sqrt(lam)
        rx *= r
        ry *= r

    num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p
    den = rx * rx * y1p * y1p + ry * ry * x1p * x1p
    coef = (-1.0 if large == sweep else 1.0) * math.sqrt(max(0.0, num / den))
    cxp = coef * (rx * y1p / ry)
    cyp = coef * (-ry * x1p / rx)
    cx = cp * cxp - sp * cyp + (x0 + x1) / 2
    cy = sp * cxp + cp * cyp + (y0 + y1) / 2

    ux, uy = (x1p - cxp) / rx, (y1p - cyp) / ry
    vx, vy = (-x1p - cxp) / rx, (-y1p - cyp) / ry
    th1 = math.atan2(uy, ux)
    dth = math.atan2(vy, vx) - th1

    if sweep == 0 and dth > 0:
        dth -= 2 * math.pi
    if sweep == 1 and dth < 0:
        dth += 2 * math.pi

    if large == 1 and abs(dth) < math.pi:
        dth = dth - 2 * math.pi if dth >= 0 else dth + 2 * math.pi
    if large == 0 and abs(dth) > math.pi:
        dth = dth + 2 * math.pi if dth < 0 else dth - 2 * math.pi

    n = max(1, math.ceil(abs(dth) / (math.pi / 2)))
    segs: list[tuple[tuple[float, float], tuple[float, float], tuple[float, float], tuple[float, float]]] = []
    for i in range(n):
        a1 = th1 + dth * i / n
        a2 = th1 + dth * (i + 1) / n
        alpha = math.sin(a2 - a1) * (math.sqrt(4 + 3 * math.tan((a2 - a1) / 2) ** 2) - 1) / 3
        p1 = (
            cx + rx * math.cos(a1) - alpha * rx * math.sin(a1),
            cy + ry * math.sin(a1) + alpha * ry * math.cos(a1),
        )
        p2 = (
            cx + rx * math.cos(a2) + alpha * rx * math.sin(a2),
            cy + ry * math.sin(a2) - alpha * ry * math.cos(a2),
        )
        p3 = (cx + rx * math.cos(a2), cy + ry * math.sin(a2))
        segs.append(((x0, y0), p1, p2, p3))
        x0, y0 = p3
    return segs


_TOKEN_RE = re.compile(
    r"([MmLlHhVvCcSsQqTtAaZz])|([-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?)"
)


def parse_path(d: str) -> tuple[list[tuple[str, list[float]]], tuple[float, float] | None]:
    """Parse an SVG path ``d`` string into a list of ``(command, args)`` tuples.

    Returns ``(cmds, start_point)`` where *start_point* is the first ``M`` coordinate
    (or ``None`` if the path is empty).
    """
    seq: list[str | float] = []
    for m in _TOKEN_RE.finditer(d):
        seq.append(m.group(1) or float(m.group(2)))

    cmds: list[tuple[str, list[float]]] = []
    cur = "M"
    last: tuple[float, float] | None = None
    start: tuple[float, float] | None = None
    prev_c: tuple[float, float] | None = None
    prev_cmd: str | None = None
    pos = 0

    while pos < len(seq):
        tok = seq[pos]
        if isinstance(tok, str):
            cur = tok
            pos += 1
            if cur in "Zz":
                cmds.append(("Z", []))
                continue
        args: list[float] = []
        while pos < len(seq) and isinstance(seq[pos], float):
            args.append(seq[pos])
            pos += 1
        if cur in "Zz":
            continue

        C = cur.upper()
        rel = cur.islower()

        if C == "M":
            for k in range(0, len(args), 2):
                if k + 1 >= len(args):
                    break
                x, y = args[k], args[k + 1]
                if rel and last is not None:
                    x += last[0]
                    y += last[1]
                cmds.append(("M", [x, y]))
                last = (x, y)
                start = (x, y)
                prev_cmd = "M"
        elif C == "L":
            for k in range(0, len(args), 2):
                if k + 1 >= len(args):
                    break
                x, y = args[k], args[k + 1]
                if rel and last is not None:
                    x += last[0]
                    y += last[1]
                cmds.append(("L", [x, y]))
                last = (x, y)
                prev_cmd = "L"
        elif C == "H":
            for x in args:
                if rel and last is not None:
                    x += last[0]
                cmds.append(("L", [x, last[1] if last else 0]))
                last = (x, last[1] if last else 0)
                prev_cmd = "L"
        elif C == "V":
            for y in args:
                if rel and last is not None:
                    y += last[1]
                cmds.append(("L", [last[0] if last else 0, y]))
                last = (last[0] if last else 0, y)
                prev_cmd = "L"
        elif C == "C":
            for k in range(0, len(args), 6):
                if k + 5 >= len(args):
                    break
                x1, y1, x2, y2, x, y = args[k : k + 6]
                if rel and last is not None:
                    x1 += last[0]
                    y1 += last[1]
                    x2 += last[0]
                    y2 += last[1]
                    x += last[0]
                    y += last[1]
                cmds.append(("C", [x1, y1, x2, y2, x, y]))
                prev_c = (x2, y2)
                last = (x, y)
                prev_cmd = "C"
        elif C == "S":
            for k in range(0, len(args), 4):
                if k + 3 >= len(args):
                    break
                x2, y2, x, y = args[k : k + 4]
                if rel and last is not None:
                    x2 += last[0]
                    y2 += last[1]
                    x += last[0]
                    y += last[1]
                if prev_cmd in ("C", "S") and last is not None and prev_c is not None:
                    x1 = 2 * last[0] - prev_c[0]
                    y1 = 2 * last[1] - prev_c[1]
                elif last is not None:
                    x1, y1 = last
                else:
                    x1, y1 = 0, 0
                cmds.append(("C", [x1, y1, x2, y2, x, y]))
                prev_c = (x2, y2)
                last = (x, y)
                prev_cmd = "S"
        elif C == "Q":
            for k in range(0, len(args), 4):
                if k + 3 >= len(args):
                    break
                qx, qy, x, y = args[k : k + 4]
                if rel and last is not None:
                    qx += last[0]
                    qy += last[1]
                    x += last[0]
                    y += last[1]
                if last is not None:
                    x1 = last[0] + 2 / 3 * (qx - last[0])
                    y1 = last[1] + 2 / 3 * (qy - last[1])
                else:
                    x1, y1 = qx, qy
                x2 = x + 2 / 3 * (qx - x)
                y2 = y + 2 / 3 * (qy - y)
                cmds.append(("C", [x1, y1, x2, y2, x, y]))
                prev_c = (qx, qy)
                last = (x, y)
                prev_cmd = "Q"
        elif C == "T":
            for k in range(0, len(args), 2):
                if k + 1 >= len(args):
                    break
                x, y = args[k], args[k + 1]
                if rel and last is not None:
                    x += last[0]
                    y += last[1]
                if prev_cmd in ("Q", "T") and last is not None and prev_c is not None:
                    qx = 2 * last[0] - prev_c[0]
                    qy = 2 * last[1] - prev_c[1]
                elif last is not None:
                    qx, qy = last
                else:
                    qx, qy = 0, 0
                if last is not None:
                    x1 = last[0] + 2 / 3 * (qx - last[0])
                    y1 = last[1] + 2 / 3 * (qy - last[1])
                else:
                    x1, y1 = qx, qy
                x2 = x + 2 / 3 * (qx - x)
                y2 = y + 2 / 3 * (qy - y)
                cmds.append(("C", [x1, y1, x2, y2, x, y]))
                prev_c = (qx, qy)
                last = (x, y)
                prev_cmd = "T"
        elif C == "A":
            for k in range(0, len(args), 7):
                if k + 6 >= len(args):
                    break
                rx, ry, rot, la, sw, x, y = args[k : k + 7]
                if rel and last is not None:
                    x += last[0]
                    y += last[1]
                cmds.append(("A", [rx, ry, rot, int(la), int(sw), x, y]))
                last = (x, y)
                prev_cmd = "A"

    return cmds, start


def to_beziers(
    cmds: list[tuple[str, list[float]]],
) -> list[list[tuple[tuple[float, float], tuple[float, float], tuple[float, float], tuple[float, float]]]]:
    """Convert parsed path commands to a list of subpaths of cubic Bezier segments.

    Each subpath is a list of ``((p0), (c1), (c2), (p1))`` tuples.
    Line segments are represented as degenerate cubics with control points equal to endpoints.
    """
    subs: list[list[tuple[tuple[float, float], tuple[float, float], tuple[float, float], tuple[float, float]]]] = []
    cur: list[tuple[tuple[float, float], tuple[float, float], tuple[float, float], tuple[float, float]]] = []
    start: tuple[float, float] | None = None
    last: tuple[float, float] | None = None

    for cmd, args in cmds:
        if cmd == "M":
            if cur:
                subs.append(cur)
            cur = []
            last = start = tuple(args)
        elif cmd == "L":
            if last is not None:
                cur.append((last, last, tuple(args), tuple(args)))
            last = tuple(args)
        elif cmd == "C":
            if last is not None:
                cur.append((last, (args[0], args[1]), (args[2], args[3]), (args[4], args[5])))
            last = (args[4], args[5])
        elif cmd == "A":
            if last is not None:
                for seg in arc_to_cubics(last[0], last[1], *args):
                    cur.append(seg)
                last = (args[5], args[6])
        elif cmd == "Z":
            if last is not None and start is not None and last != start:
                cur.append((last, last, start, start))
            last = start
            if cur:
                subs.append(cur)
            cur = []

    if cur:
        subs.append(cur)
    return subs
