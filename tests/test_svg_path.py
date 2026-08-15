"""Tests for svg_compiler._path — SVG path parsing + arc conversion."""

from ppt_pro_max.renderer.svg_compiler._path import arc_to_cubics, parse_path, to_beziers


class TestParsePathMoveto:
    def test_absolute_M(self):
        cmds, start = parse_path("M 10 20")
        assert cmds == [("M", [10.0, 20.0])]
        assert start == (10.0, 20.0)

    def test_relative_m(self):
        cmds, _ = parse_path("M 10 20 m 5 5")
        assert cmds[-1] == ("M", [15.0, 25.0])

    def test_multiple_M_coords(self):
        cmds, _ = parse_path("M 10 20 30 40")
        assert cmds[0] == ("M", [10.0, 20.0])
        assert cmds[1] == ("M", [30.0, 40.0])


class TestParsePathLineto:
    def test_absolute_L(self):
        cmds, _ = parse_path("M 0 0 L 10 20")
        assert cmds[-1] == ("L", [10.0, 20.0])

    def test_relative_l(self):
        cmds, _ = parse_path("M 0 0 l 10 20")
        assert cmds[-1] == ("L", [10.0, 20.0])

    def test_H(self):
        cmds, _ = parse_path("M 0 0 H 50")
        assert cmds[-1] == ("L", [50.0, 0.0])

    def test_h(self):
        cmds, _ = parse_path("M 10 0 h 50")
        assert cmds[-1] == ("L", [60.0, 0.0])

    def test_V(self):
        cmds, _ = parse_path("M 0 0 V 50")
        assert cmds[-1] == ("L", [0.0, 50.0])

    def test_v(self):
        cmds, _ = parse_path("M 0 10 v 50")
        assert cmds[-1] == ("L", [0.0, 60.0])


class TestParsePathCubic:
    def test_absolute_C(self):
        cmds, _ = parse_path("M 0 0 C 10 20 30 40 50 60")
        assert cmds[-1] == ("C", [10.0, 20.0, 30.0, 40.0, 50.0, 60.0])

    def test_relative_c(self):
        cmds, _ = parse_path("M 10 10 c 5 5 10 10 15 15")
        assert cmds[-1] == ("C", [15.0, 15.0, 20.0, 20.0, 25.0, 25.0])

    def test_S_smooth(self):
        cmds, _ = parse_path("M 0 0 C 10 20 30 40 50 60 S 70 80 90 100")
        # S reflects previous control point: (2*50-30, 2*60-40) = (70, 80)
        assert cmds[-1] == ("C", [70.0, 80.0, 70.0, 80.0, 90.0, 100.0])


class TestParsePathQuadratic:
    def test_Q(self):
        cmds, _ = parse_path("M 0 0 Q 25 50 50 0")
        # Q is converted to C internally
        assert cmds[-1][0] == "C"

    def test_T(self):
        cmds, _ = parse_path("M 0 0 Q 25 50 50 0 T 100 0")
        # T reflects previous Q control point
        assert cmds[-1][0] == "C"


class TestParsePathArc:
    def test_A(self):
        cmds, _ = parse_path("M 0 0 A 50 50 0 1 1 100 0")
        assert cmds[-1] == ("A", [50.0, 50.0, 0.0, 1, 1, 100.0, 0.0])

    def test_relative_a(self):
        cmds, _ = parse_path("M 0 0 a 50 50 0 1 1 100 0")
        assert cmds[-1] == ("A", [50.0, 50.0, 0.0, 1, 1, 100.0, 0.0])


class TestParsePathClose:
    def test_Z(self):
        cmds, _ = parse_path("M 0 0 L 10 0 L 10 10 Z")
        assert cmds[-1] == ("Z", [])

    def test_z(self):
        cmds, _ = parse_path("M 0 0 L 10 0 L 10 10 z")
        assert cmds[-1] == ("Z", [])


class TestParsePathEmpty:
    def test_empty(self):
        cmds, start = parse_path("")
        assert cmds == []
        assert start is None


class TestArcToCubics:
    def test_zero_radius(self):
        segs = arc_to_cubics(0, 0, 0, 0, 0, 0, 0, 10, 10)
        assert len(segs) == 1
        assert segs[0][3] == (10.0, 10.0)

    def test_quarter_circle(self):
        segs = arc_to_cubics(0, 1, 1, 1, 0, 0, 1, 1, 0)
        assert len(segs) >= 1
        # endpoint should be (1, 0)
        last = segs[-1]
        assert abs(last[3][0] - 1.0) < 1e-6
        assert abs(last[3][1] - 0.0) < 1e-6

    def test_full_circle_idiom(self):
        """The 'a r,r 0 1,0 dx,0' idiom for full circles."""
        segs = arc_to_cubics(0, 0, 50, 50, 0, 1, 0, 0.01, 0)
        assert len(segs) >= 2  # should produce multiple segments

    def test_semicircle(self):
        segs = arc_to_cubics(0, 0, 50, 50, 0, 1, 1, 100, 0)
        assert len(segs) >= 1
        last = segs[-1]
        assert abs(last[3][0] - 100.0) < 1e-6
        assert abs(last[3][1] - 0.0) < 1e-6


class TestToBeziers:
    def test_line_to_bezier(self):
        cmds, _ = parse_path("M 0 0 L 10 20")
        subs = to_beziers(cmds)
        assert len(subs) == 1
        assert len(subs[0]) == 1
        seg = subs[0][0]
        assert seg[0] == (0.0, 0.0)
        assert seg[3] == (10.0, 20.0)

    def test_cubic_passthrough(self):
        cmds, _ = parse_path("M 0 0 C 10 20 30 40 50 60")
        subs = to_beziers(cmds)
        assert len(subs) == 1
        assert subs[0][0] == ((0.0, 0.0), (10.0, 20.0), (30.0, 40.0), (50.0, 60.0))

    def test_close_path(self):
        cmds, _ = parse_path("M 0 0 L 10 0 L 10 10 Z")
        subs = to_beziers(cmds)
        assert len(subs) == 1  # Z closes the subpath

    def test_multiple_subpaths(self):
        cmds, _ = parse_path("M 0 0 L 10 0 Z M 20 20 L 30 20 Z")
        subs = to_beziers(cmds)
        assert len(subs) == 2

    def test_arc_in_path(self):
        cmds, _ = parse_path("M 0 0 A 50 50 0 1 1 100 0")
        subs = to_beziers(cmds)
        assert len(subs) == 1
        assert len(subs[0]) >= 1
