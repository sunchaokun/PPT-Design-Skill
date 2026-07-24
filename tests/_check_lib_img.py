"""Check library component image status for the 3D component with images."""

import os
from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from ppt_pro_max.enterprise.group_extractor import GroupExtractor
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

lib = ComponentLibrary(find_db_path())
ext = GroupExtractor()
def_dir = r"E:\BaiduNetdiskDownload\DEF"

for root_dir, dirs, files in os.walk(def_dir):
    for fname in files:
        if not fname.endswith(".pptx"):
            continue
        fpath = os.path.join(root_dir, fname)
        try:
            results = ext.extract_all(fpath)
            for r in results:
                img_blobs = r.get("image_blobs", {})
                root = etree.fromstring(r["xml_parts"]["group"])
                has_3d = len(list(root.iter(f"{{{_A}}}scene3d"))) > 0
                if has_3d and len(img_blobs) >= 3:
                    node_count = r["node_count"]
                    cat = r["category"]
                    lib_results = lib.search(type="group", category=cat, limit=50)
                    for lr in lib_results:
                        if lr["node_count"] == node_count:
                            xml_parts = lib.load_xml(lr["id"])
                            img_keys = [k for k in xml_parts if k.startswith("img_")]
                            lroot = etree.fromstring(xml_parts["group"])
                            lpic = len(list(lroot.iter(f"{{{_P}}}pic")))
                            lblip = len(list(lroot.iter(f"{{{_A}}}blip")))
                            print(f"Extracted: nodes={node_count}, cat={cat}, imgs={list(img_blobs.keys())}")
                            print(f"Library: id={lr['id']}, img_keys={img_keys}, pic={lpic}, blip={lblip}")
                            print(f"Library XML dir: {lr.get('xml_path', '?')}")
                            raise SystemExit(0)
        except SystemExit:
            raise
        except Exception:
            pass
    break

lib.close()
