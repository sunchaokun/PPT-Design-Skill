"""Trace 000000 parent chain in hierarchy component."""

from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

lib = ComponentLibrary(find_db_path())
results = lib.search(type="group", category="hierarchy", limit=1)
xml_parts = lib.load_xml(results[0]["id"])
root = etree.fromstring(xml_parts["group"])

count = 0
for srgb in root.iter(f"{{{_A}}}srgbClr"):
    if srgb.get("val", "").upper() == "000000":
        if count >= 5:
            break
        chain = []
        p = srgb
        while p is not None:
            tag = p.tag.split("}")[-1] if "}" in p.tag else p.tag
            chain.append(tag)
            p = p.getparent()
        chain.reverse()
        print("  chain: " + " > ".join(chain))
        count += 1

lib.close()
