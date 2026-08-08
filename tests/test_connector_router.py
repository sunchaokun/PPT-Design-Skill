"""Tests for connector_router — intelligent connector routing around nodes."""

from __future__ import annotations

from ppt_pro_max.renderer.diagram.connector_router import (
    route_connector,
    _line_intersects_rect,
    _segments_intersect,
)


class TestRouteConnector:

    def test_straight_line_no_obstruction(self):
        nodes = [{"x": 5, "y": 5, "width": 1, "height": 1}]
        result = route_connector(0, 0, 10, 0, nodes)
        assert len(result) == 2
        assert result[0] == (0, 0)
        assert result[1] == (10, 0)

    def test_reroute_around_node(self):
        nodes = [{"x": 4, "y": -0.5, "width": 2, "height": 1}]
        result = route_connector(0, 0, 10, 0, nodes)
        assert len(result) >= 3

    def test_no_nodes(self):
        result = route_connector(0, 0, 5, 5, [])
        assert len(result) == 2

    def test_same_point(self):
        result = route_connector(3, 3, 3, 3, [])
        assert len(result) == 2

    def test_diagonal_no_obstruction(self):
        nodes = [{"x": 10, "y": 10, "width": 2, "height": 2}]
        result = route_connector(0, 0, 5, 5, nodes)
        assert len(result) == 2


class TestLineIntersectsRect:

    def test_line_through_rect(self):
        assert _line_intersects_rect(0, 0, 10, 0, 4, -1, 2, 2)

    def test_line_misses_rect(self):
        assert not _line_intersects_rect(0, 0, 10, 0, 4, 5, 2, 2)

    def test_line_touches_corner(self):
        assert _line_intersects_rect(0, 0, 5, 5, 2, 2, 1, 1)

    def test_endpoint_inside_rect(self):
        assert _line_intersects_rect(3, 0, 3, 5, 2, 1, 3, 3)


class TestSegmentsIntersect:

    def test_crossing_segments(self):
        assert _segments_intersect(0, 0, 5, 5, 0, 5, 5, 0)

    def test_parallel_segments(self):
        assert not _segments_intersect(0, 0, 5, 0, 0, 1, 5, 1)

    def test_collinear_non_overlapping(self):
        assert not _segments_intersect(0, 0, 2, 0, 3, 0, 5, 0)

    def test_touching_endpoints(self):
        assert _segments_intersect(0, 0, 3, 0, 3, 0, 5, 0)
