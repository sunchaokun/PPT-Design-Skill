"""Test add+load cycle with images for 3D+pic component."""

import os
from ppt_pro_max.enterprise.group_extractor import GroupExtractor
from ppt_pro_max.enterprise.component_library import ComponentLibrary
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

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
                ib = r.get("image_blobs", {})
                root = etree.fromstring(r["xml_parts"]["group"])
                has_3d = len(list(root.iter(f"{{{_A}}}scene3d"))) > 0
                pic_count = len(list(root.iter(f"{{{_P}}}pic")))
                if has_3d and pic_count > 0:
                    grp_xml = r["xml_parts"]["group"]
                    grp_root = etree.fromstring(grp_xml)
                    pic_in_xml = len(list(grp_root.iter(f"{{{_P}}}pic")))
                    blip_in_xml = len(list(grp_root.iter(f"{{{_A}}}blip")))
                    scene3d_in_xml = len(list(grp_root.iter(f"{{{_A}}}scene3d")))
                    print(f"Before add(): pic={pic_in_xml}, blip={blip_in_xml}, scene3d={scene3d_in_xml}, img_blobs={len(ib)}")

                    lib = ComponentLibrary(r"E:\PPT-Design-Skill\component_library\index.db")
                    cid = lib.add(
                        type=r["type"],
                        category=r.get("category", "process"),
                        variant=r.get("variant", ""),
                        node_count=r.get("node_count", 0),
                        level_count=r.get("level_count", 0),
                        xml_parts=r.get("xml_parts", {}),
                        image_blobs=ib,
                    )
                    print(f"Added with id={cid}")

                    xml_parts2 = lib.load_xml(cid)
                    img_keys2 = [k for k in xml_parts2 if k.startswith("img_")]
                    root2 = etree.fromstring(xml_parts2["group"])
                    pic2 = len(list(root2.iter(f"{{{_P}}}pic")))
                    blip2 = len(list(root2.iter(f"{{{_A}}}blip")))
                    scene3d2 = len(list(root2.iter(f"{{{_A}}}scene3d")))
                    print(f"After load_xml(): pic={pic2}, blip={blip2}, scene3d={scene3d2}, img_keys={img_keys2}")

                    lib.close()
                    raise SystemExit(0)
        except SystemExit:
            raise
        except Exception as e:
            print(f"Error: {e}")
    break
