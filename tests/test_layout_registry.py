"""Tests for layout registry."""

from ppt_pro_max.renderer.layout_registry import LayoutRegistry, SLIDE_WIDTH, SLIDE_HEIGHT, MASTER_LAYOUTS


def test_slide_dimensions():
    assert SLIDE_WIDTH == 13.333
    assert SLIDE_HEIGHT == 7.5


def test_all_layouts_defined():
    assert len(MASTER_LAYOUTS) >= 12


def test_layout_ids_unique():
    ids = [v["id"] for v in MASTER_LAYOUTS.values()]
    assert len(ids) == len(set(ids))


def test_get_layout():
    registry = LayoutRegistry()
    layout = registry.get_layout("title-slide")
    assert layout is not None
    assert "title" in layout["placeholders"]


def test_get_layout_by_goal():
    registry = LayoutRegistry()
    layout = registry.get_layout_by_goal("hook")
    assert layout is not None
    assert layout["name"] == "Title Slide"


def test_get_layout_by_goal_fallback():
    registry = LayoutRegistry()
    layout = registry.get_layout_by_goal("unknown-goal")
    assert layout is not None
    assert layout["name"] == "Content With Title"


def test_list_layouts():
    registry = LayoutRegistry()
    layouts = registry.list_layouts()
    assert len(layouts) >= 12
    assert "title-slide" in layouts


def test_placeholder_coordinates_within_bounds():
    for layout in MASTER_LAYOUTS.values():
        for ph_name, ph in layout.get("placeholders", {}).items():
            if ph.get("type") in ("image", "chart"):
                continue
            assert ph["x"] >= 0, f"{layout['name']}/{ph_name}: x={ph['x']} < 0"
            assert ph["y"] >= 0, f"{layout['name']}/{ph_name}: y={ph['y']} < 0"
            assert ph["x"] + ph["width"] <= SLIDE_WIDTH + 0.1, f"{layout['name']}/{ph_name}: x+w={ph['x']+ph['width']} > {SLIDE_WIDTH}"
            assert ph["y"] + ph["height"] <= SLIDE_HEIGHT + 0.1, f"{layout['name']}/{ph_name}: y+h={ph['y']+ph['height']} > {SLIDE_HEIGHT}"
