"""Re-import the 3D+pic component into library with correct images."""

import os
from ppt_pro_max.enterprise.group_extractor import GroupExtractor
from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
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
                ib = r.get("image_blobs", {})
                root = etree.fromstring(r["xml_parts"]["group"])
                has_3d = len(list(root.iter(f"{{{_A}}}scene3d"))) > 0
                pic_count = len(list(root.iter(f"{{{_P}}}pic")))
                if has_3d and pic_count > 0:
                    nc = r["node_count"]
                    cat = r["category"]

                    # Find existing component with same node_count + category
                    existing = lib.search(type="group", category=cat, limit=100)
                    old_id = None
                    for e in existing:
                        if e["node_count"] == nc:
                            old_id = e["id"]
                            break

                    if old_id:
                        # Delete old entry so we can re-add with images
                        print(f"Deleting old id={old_id} (cat={cat}, nodes={nc})")
                        lib._db.execute("DELETE FROM components WHERE id = ?", (old_id,))

                    cid = lib.add(
                        type=r["type"],
                        category=cat,
                        variant=r.get("variant", ""),
                        node_count=nc,
                        level_count=r.get("level_count", 0),
                        xml_parts=r.get("xml_parts", {}),
                        image_blobs=ib,
                    )

                    # Verify
                    xml_parts2 = lib.load_xml(cid)
                    img_keys = [k for k in xml_parts2 if k.startswith("img_")]
                    root2 = etree.fromstring(xml_parts2["group"])
                    pic2 = len(list(root2.iter(f"{{{_P}}}pic")))
                    blip2 = len(list(root2.iter(f"{{{_A}}}blip")))
                    scene3d2 = len(list(root2.iter(f"{{{_A}}}scene3d")))

                    print(f"New id={cid}: pic={pic2}, blip={blip2}, scene3d={scene3d2}, img_keys={img_keys}")
                    lib._db.commit()
                    raise SystemExit(0)
        except SystemExit:
            raise
        except Exception as e:
            print(f"Error: {e}")

lib.close()
