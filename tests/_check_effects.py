"""Check what effects are in broken components."""

from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

lib = ComponentLibrary(find_db_path())

for cat in ["hierarchy", "infographic", "chart"]:
    results = lib.search(type="group", category=cat, limit=1)
    if not results:
        continue
    comp = results[0]
    xml_parts = lib.load_xml(comp["id"])
    root = etree.fromstring(xml_parts["group"])

    print(f"\n{cat} (id={comp['id']}):")

    effect_types = {}
    for ef in root.iter(f"{{{_A}}}effectLst"):
        for child in ef:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            effect_types[tag] = effect_types.get(tag, 0) + 1
    print(f"  Effect types: {effect_types}")

    # Check sp3d
    sp3d_count = 0
    for sp3d in root.iter(f"{{{_A}}}sp3d"):
        sp3d_count += 1
    if sp3d_count:
        print(f"  sp3d (3D): {sp3d_count}")

    # Sample one shape's full XML (first shape with text)
    _P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    for sp in root.iter(f"{{{_P}}}sp"):
        texts = [t.text.strip() for t in sp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
        if texts:
            xml_str = etree.tostring(sp, encoding="unicode", pretty_print=True)
            lines = xml_str.split("\n")
            print(f"  Sample shape (text={texts[0][:15]}): {len(lines)} lines")
            # Print just the spPr section
            spPr = sp.find(f"{{{_P}}}spPr")
            if spPr is not None:
                spPr_str = etree.tostring(spPr, encoding="unicode", pretty_print=True)
                print(f"  spPr:\n{spPr_str[:500]}")
            break

lib.close()
