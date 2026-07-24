"""Find 3D components with images in DEF and check library match."""

import os
from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from ppt_pro_max.enterprise.group_extractor import GroupExtractor
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

lib = ComponentLibrary(find_db_path())
ext = GroupExtractor()
def_dir = r"E:\BaiduNetdiskDownload\DEF"

count = 0
for root_dir, dirs, files in os.walk(def_dir):
    for fname in files:
        if not fname.endswith(".pptx"):
            continue
        fpath = os.path.join(root_dir, fname)
        try:
            results = ext.extract_all(fpath)
            for r in results:
                ib = len(r.get("image_blobs", {}))
                root = etree.fromstring(r["xml_parts"]["group"])
                has_3d = len(list(root.iter(f"{{{_A}}}scene3d"))) > 0
                pic_count = len(list(root.iter(f"{{{_P}}}pic")))
                if has_3d and pic_count > 0:
                    nc = r["node_count"]
                    cat = r["category"]
                    # Find in library
                    lib_results = lib.search(type="group", category=cat, limit=100)
                    for lr in lib_results:
                        if lr["node_count"] == nc:
                            xml_parts = lib.load_xml(lr["id"])
                            img_keys = [k for k in xml_parts if k.startswith("img_")]
                            lroot = etree.fromstring(xml_parts["group"])
                            lpic = len(list(lroot.iter(f"{{{_P}}}pic")))
                            lblip = len(list(lroot.iter(f"{{{_A}}}blip")))
                            l3d = len(list(lroot.iter(f"{{{_A}}}scene3d"))) > 0
                            print(f"EXTRACTED: nodes={nc} cat={cat} imgs={ib} pic={pic_count}")
                            print(f"LIBRARY:   id={lr['id']} img_keys={img_keys} pic={lpic} blip={lblip} 3d={l3d}")
                            count += 1
                            break
                    if count >= 5:
                        raise SystemExit(0)
        except SystemExit:
            raise
        except Exception:
            pass
    if count >= 5:
        break

print(f"Total 3D+pic components matched: {count}")
lib.close()
