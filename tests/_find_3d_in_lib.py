"""Find components with 3D/pic/images in rebuilt library."""

from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

lib = ComponentLibrary(find_db_path())

for cat in ["hierarchy", "process", "swot", "infographic", "timeline", "chart"]:
    results = lib.search(type="group", category=cat, limit=30)
    for comp in results:
        xml_parts = lib.load_xml(comp["id"])
        img_keys = [k for k in xml_parts if k.startswith("img_")]
        root = etree.fromstring(xml_parts["group"])
        has_3d = len(list(root.iter(f"{{{_A}}}scene3d"))) > 0
        has_pic = len(list(root.iter(f"{{{_P}}}pic"))) > 0
        if has_3d or has_pic or img_keys:
            nc = comp["node_count"]
            cid = comp["id"]
            print(f"{cat} id={cid}: nodes={nc}, 3d={has_3d}, pic={has_pic}, imgs={img_keys}")

lib.close()
