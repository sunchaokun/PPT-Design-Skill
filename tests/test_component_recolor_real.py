"""Generate test PPTs with real library components + green brand colors."""

import os
from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from ppt_pro_max.enterprise.component_renderer import ComponentRenderer
from ppt_pro_max.enterprise.brand_spec import BrandSpec
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

GREEN_BRAND = BrandSpec(colors={
    "primary": "#00AA00",
    "accent": "#66CC00",
    "muted": "#E6F5E6",
    "foreground": "#1A3A1A",
    "on-primary": "#FFFFFF",
    "background": "#FFFFFF",
})

OUT_DIR = os.path.join(os.path.dirname(__file__), "_test_output")
os.makedirs(OUT_DIR, exist_ok=True)

OLD_COLORS = {"7B9CFE", "85898F", "5A6783", "3D485D", "B5B5B5", "A3CF62",
              "354F73", "8C6844", "B9A799", "344E71", "2F4768", "2B4261"}


def _test_category(cat, user_texts):
    lib = ComponentLibrary(find_db_path())
    results = lib.search(type="group", category=cat, limit=1)
    assert results, f"No {cat} components in library"
    comp = results[0]
    xml_parts = lib.load_xml(comp["id"])
    assert "group" in xml_parts

    renderer = ComponentRenderer()
    filled = renderer._fill_group_data(xml_parts, user_texts)
    styled = renderer._apply_brand_colors(filled, GREEN_BRAND)

    root = etree.fromstring(styled["group"])

    fill_vals = set()
    for sf in root.iter(f"{{{_A}}}solidFill"):
        srgb = sf.find(f"{{{_A}}}srgbClr")
        if srgb is not None:
            fill_vals.add(srgb.get("val", "").upper())

    leaked = fill_vals & OLD_COLORS
    assert not leaked, f"{cat}: old colors leaked: {leaked}"

    all_t = [t for t in root.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
    texts = [t.text.strip() for t in all_t]

    for ut in user_texts:
        assert ut in texts, f"{cat}: user text '{ut}' not found in {texts}"

    lib.close()
    print(f"  {cat}: id={comp['id']}, fills={fill_vals}, texts={texts[:8]}, total={len(all_t)}")
    return True


def test_recolor_hierarchy():
    assert _test_category("hierarchy", ["CEO", "CTO", "CFO", "VP Eng", "VP Sales"])


def test_recolor_process():
    assert _test_category("process", ["Plan", "Design", "Build", "Launch"])


def test_recolor_swot():
    assert _test_category("swot", ["Strengths", "Weaknesses", "Opportunities", "Threats"])


def test_recolor_infographic():
    assert _test_category("infographic", ["Revenue", "Growth", "Profit", "Market Share"])


def test_recolor_timeline():
    lib = ComponentLibrary(find_db_path())
    results = lib.search(type="group", category="timeline", limit=1)
    if not results:
        lib.close()
        print("  timeline: SKIPPED (no components)")
        return
    lib.close()
    _test_category("timeline", ["Q1", "Q2", "Q3", "Q4"])


def test_recolor_chart():
    lib = ComponentLibrary(find_db_path())
    results = lib.search(type="group", category="chart", limit=1)
    if not results:
        lib.close()
        print("  chart: SKIPPED (no components)")
        return
    lib.close()
    _test_category("chart", ["Sales", "Profit"])


def test_multiple_components_same_brand():
    """Test that multiple components from different categories all get the same green brand."""
    lib = ComponentLibrary(find_db_path())
    renderer = ComponentRenderer()

    all_fill_vals = set()
    for cat, texts in [
        ("hierarchy", ["CEO", "CTO", "CFO"]),
        ("process", ["Step 1", "Step 2", "Step 3"]),
        ("infographic", ["Metric A", "Metric B"]),
    ]:
        results = lib.search(type="group", category=cat, limit=1)
        if not results:
            continue
        xml_parts = lib.load_xml(results[0]["id"])
        filled = renderer._fill_group_data(xml_parts, texts)
        styled = renderer._apply_brand_colors(filled, GREEN_BRAND)
        root = etree.fromstring(styled["group"])
        for sf in root.iter(f"{{{_A}}}solidFill"):
            srgb = sf.find(f"{{{_A}}}srgbClr")
            if srgb is not None:
                all_fill_vals.add(srgb.get("val", "").upper())

    leaked = all_fill_vals & OLD_COLORS
    assert not leaked, f"Leaked old colors across components: {leaked}"
    assert "00AA00" in all_fill_vals, "Primary green not found"
    lib.close()
    print(f"  multi-brand: fills={all_fill_vals}, leaked={leaked}")


if __name__ == "__main__":
    print("=== Hierarchy ===")
    test_recolor_hierarchy()
    print("=== Process ===")
    test_recolor_process()
    print("=== SWOT ===")
    test_recolor_swot()
    print("=== Infographic ===")
    test_recolor_infographic()
    print("=== Timeline ===")
    test_recolor_timeline()
    print("=== Chart ===")
    test_recolor_chart()
    print("=== Multi-component brand consistency ===")
    test_multiple_components_same_brand()
    print("ALL DONE")
