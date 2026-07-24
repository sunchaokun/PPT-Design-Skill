"""Find all 3D components in DEF and check image_blobs."""

import os
from ppt_pro_max.enterprise.group_extractor import GroupExtractor
from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

ext = GroupExtractor()
lib = ComponentLibrary(find_db_path())
def_dir = r"E:\BaiduNetdiskDownload\DEF"

found_3d = 0
for root_dir, dirs, files in os.walk(def_dir):
    for fname in files:
        if not fname.endswith(".pptx"):
            continue
        fpath = os.path.join(root_dir, fname)
        try:
            results = ext.extract_all(fpath)
            for r in results:
                root = etree.fromstring(r["xml_parts"]["group"])
                has_3d = len(list(root.iter(f"{{{_A}}}scene3d"))) > 0
                if has_3d:
                    ib = r.get("image_blobs", {})
                    pic_count = len(list(root.iter(f"{{{_P}}}pic")))
                    nc = r["node_count"]
                    cat = r["category"]

                    # Find in library
                    lib_results = lib.search(type="group", category=cat, limit=100)
                    lib_match = None
                    for lr in lib_results:
                        if lr["node_count"] == nc:
                            lib_match = lr
                            break

                    if lib_match:
                        xml_parts = lib.load_xml(lib_match["id"])
                        img_keys = [k for k in xml_parts if k.startswith("img_")]
                        lroot = etree.fromstring(xml_parts["group"])
                        lpic = len(list(lroot.iter(f"{{{_P}}}pic")))
                        lblip = len(list(lroot.iter(f"{{{_A}}}blip")))
                        l3d = len(list(lroot.iter(f"{{{_A}}}scene3d"))) > 0
                        print(f"EXTRACT: nodes={nc} cat={cat} imgs={len(ib)} pic={pic_count}")
                        print(f"LIBRARY: id={lib_match['id']} img_keys={img_keys} pic={lpic} blip={lblip} 3d={l3d}")
                        print(f"  DIFF: pic_lost={pic_count - lpic}, img_lost={len(ib) - len(img_keys)}, 3d_lost={has_3d and not l3d}")
                        print()

                    found_3d += 1
                    if found_3d >= 10:
                        raise SystemExit(0)
        except SystemExit:
            raise
        except Exception:
            pass
    if found_3d >= 10:
        break

print(f"Total 3D components checked: {found_3d}")
lib.close()
