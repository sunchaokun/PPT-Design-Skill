"""Trace image_blobs through bulk_import to find where they're lost."""

import os
from ppt_pro_max.enterprise.group_extractor import GroupExtractor
from ppt_pro_max.enterprise.component_library import ComponentLibrary
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

def_dir = r"E:\BaiduNetdiskDownload\DEF"
ext = GroupExtractor()

# Find the specific PPT with 3D+pic
for root_dir, dirs, files in os.walk(def_dir):
    for fname in files:
        if not fname.endswith(".pptx"):
            continue
        fpath = os.path.join(root_dir, fname)
        try:
            results = ext.extract_all(fpath)
            for r in results:
                ib = r.get("image_blobs", {})
                root = etree.fromstring(r["xml_parts"]["group"])
                has_3d = len(list(root.iter(f"{{{_A}}}scene3d"))) > 0
                pic_count = len(list(root.iter(f"{{{_P}}}pic")))
                if has_3d and pic_count > 0 and len(ib) > 0:
                    nc = r["node_count"]
                    cat = r["category"]
                    print(f"Component: nodes={nc}, cat={cat}")
                    print(f"  image_blobs keys: {list(ib.keys())}")
                    print(f"  image_blobs sizes: {[len(v) for v in ib.values()]}")
                    print(f"  xml_parts keys: {list(r['xml_parts'].keys())}")

                    # Check if pic/blip survived in xml_parts['group']
                    grp_xml = r["xml_parts"]["group"]
                    grp_root = etree.fromstring(grp_xml)
                    pic_after = len(list(grp_root.iter(f"{{{_P}}}pic")))
                    blip_after = len(list(grp_root.iter(f"{{{_A}}}blip")))
                    scene3d_after = len(list(grp_root.iter(f"{{{_A}}}scene3d")))
                    print(f"  In xml_parts['group']: pic={pic_after}, blip={blip_after}, scene3d={scene3d_after}")

                    raise SystemExit(0)
        except SystemExit:
            raise
        except Exception:
            pass
    break
