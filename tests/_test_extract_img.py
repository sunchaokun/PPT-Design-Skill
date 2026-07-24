"""Test extraction: check if images are extracted correctly."""

import os
from ppt_pro_max.enterprise.group_extractor import GroupExtractor
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

def_dir = r"E:\BaiduNetdiskDownload\DEF"
ext = GroupExtractor()

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
                pic_count = len(list(root.iter(f"{{{_P}}}pic")))
                if has_3d and pic_count > 0:
                    print(f"3D+pic: nodes={r['node_count']}, cat={r['category']}, imgs={list(img_blobs.keys())}, pic_in_xml={pic_count}, img_sizes={[len(v) for v in img_blobs.values()]}")
                    if len(img_blobs) > 0:
                        raise SystemExit(0)
            break
        except SystemExit:
            raise
        except Exception as e:
            print(f"Error: {e}")
        break
