"""Analyze problematic components in detail."""

from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from ppt_pro_max.enterprise.component_renderer import ComponentRenderer, _is_placeholder, _is_title_text, _shape_primary_slots
from ppt_pro_max.enterprise.brand_spec import BrandSpec
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

GREEN_BRAND = BrandSpec(colors={
    "primary": "#00AA00", "accent": "#66CC00", "muted": "#E6F5E6",
    "foreground": "#1A3A1A", "on-primary": "#FFFFFF", "background": "#FFFFFF",
})

lib = ComponentLibrary(find_db_path())
renderer = ComponentRenderer()

for cat in ["hierarchy", "infographic", "timeline"]:
    results = lib.search(type="group", category=cat, limit=1)
    if not results:
        continue
    comp = results[0]
    xml_parts = lib.load_xml(comp["id"])
    root = etree.fromstring(xml_parts["group"])

    print(f"\n{'='*60}")
    print(f"{cat} (id={comp['id']}, nodes={comp['node_count']})")
    print(f"{'='*60}")

    # 1. Check fill types
    fill_types = set()
    for child in root.iter():
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if "Fill" in tag or tag == "fill":
            fill_types.add(tag)
    print(f"Fill element types: {fill_types}")

    # 2. Check all color mechanisms
    srgb_vals = {}
    for srgb in root.iter(f"{{{_A}}}srgbClr"):
        v = srgb.get("val", "").upper()
        srgb_vals[v] = srgb_vals.get(v, 0) + 1
    print(f"srgbClr: {dict(sorted(srgb_vals.items(), key=lambda x:-x[1])[:10])}")

    # 3. Check gradFill colors
    grad_stops = {}
    for gs in root.iter(f"{{{_A}}}gs"):
        colors = []
        for srgb in gs.iter(f"{{{_A}}}srgbClr"):
            colors.append(srgb.get("val", "").upper())
        for scheme in gs.iter(f"{{{_A}}}schemeClr"):
            colors.append("scheme:" + scheme.get("val", ""))
        if colors:
            grad_stops[",".join(colors)] = grad_stops.get(",".join(colors), 0) + 1
    if grad_stops:
        print(f"gradFill stops: {dict(sorted(grad_stops.items(), key=lambda x:-x[1])[:8])}")

    # 4. Check ln (line) colors
    ln_colors = {}
    for ln in root.iter(f"{{{_A}}}ln"):
        for srgb in ln.iter(f"{{{_A}}}srgbClr"):
            v = srgb.get("val", "").upper()
            ln_colors[v] = ln_colors.get(v, 0) + 1
    if ln_colors:
        print(f"Line colors: {dict(sorted(ln_colors.items(), key=lambda x:-x[1])[:8])}")

    # 5. After recolor+fill
    texts_map = {
        "hierarchy": ["CEO", "CTO", "CFO", "VP Eng", "VP Sales"],
        "infographic": ["Revenue", "Growth", "Profit", "Market Share"],
        "timeline": ["Q1", "Q2", "Q3", "Q4"],
    }
    filled = renderer._fill_group_data(xml_parts, texts_map[cat])
    styled = renderer._apply_brand_colors(filled, GREEN_BRAND)
    root2 = etree.fromstring(styled["group"])

    srgb_after = {}
    for srgb in root2.iter(f"{{{_A}}}srgbClr"):
        v = srgb.get("val", "").upper()
        srgb_after[v] = srgb_after.get(v, 0) + 1
    print(f"\nAfter recolor srgbClr: {dict(sorted(srgb_after.items(), key=lambda x:-x[1])[:10])}")

    # Check what colors remain that are NOT brand colors
    brand_set = {"00AA00", "66CC00", "E6F5E6", "1A3A1A", "FFFFFF"}
    non_brand = {k: v for k, v in srgb_after.items() if k not in brand_set}
    if non_brand:
        print(f"NON-BRAND colors remaining: {dict(sorted(non_brand.items(), key=lambda x:-x[1])[:10])}")

    all_t = [t for t in root2.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
    print(f"After fill texts ({len(all_t)}): {[t.text.strip()[:20] for t in all_t[:8]]}")

lib.close()
