"""Debug: check why images are not extracted for 3D components."""

import os
import zipfile
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_rel = "http://schemas.openxmlformats.org/package/2006/relationships"

def_dir = r"E:\BaiduNetdiskDownload\DEF"

# Find the 3D hierarchy PPT
for root_dir, dirs, files in os.walk(def_dir):
    for fname in files:
        if not fname.endswith(".pptx"):
            continue
        fpath = os.path.join(root_dir, fname)
        try:
            with zipfile.ZipFile(fpath) as z:
                for sf in [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml") and "rels" not in n]:
                    slide_root = etree.fromstring(z.read(sf))
                    sp_tree = slide_root.find(f".//{{{_P}}}spTree")
                    if sp_tree is None:
                        continue

                    # Read slide rels
                    rels_path = sf.replace("slides/", "slides/_rels/") + ".rels"
                    slide_rels = {}
                    if rels_path in z.namelist():
                        rels_root = etree.fromstring(z.read(rels_path))
                        for rel in rels_root.findall(f"{{{_rel}}}Relationship"):
                            slide_rels[rel.get("Id", "")] = rel.get("Target", "")

                    for grp in sp_tree.iter(f"{{{_P}}}grpSp"):
                        t_count = len([t for t in grp.iter(f"{{{_A}}}t") if t.text and t.text.strip()])
                        has_3d = len(list(grp.iter(f"{{{_A}}}scene3d"))) > 0
                        if t_count >= 10 and has_3d:
                            # Check pic elements and their blip refs
                            pic_count = len(list(grp.iter(f"{{{_P}}}pic")))
                            blip_embeds = []
                            for blip in grp.iter(f"{{{_A}}}blip"):
                                embed = blip.get(f"{{{_R}}}embed")
                                blip_embeds.append(embed)

                            print(f"=== {fname}/{sf} ===")
                            print(f"  texts={t_count}, 3d=True, pic={pic_count}, blip_embeds={blip_embeds[:5]}")
                            print(f"  slide_rels keys: {list(slide_rels.keys())[:10]}")

                            # Check which embeds are in slide_rels
                            for embed in blip_embeds[:5]:
                                if embed:
                                    target = slide_rels.get(embed, "NOT FOUND")
                                    print(f"  r:embed={embed} -> {target}")
                                    if target != "NOT FOUND":
                                        abs_target = os.path.normpath(os.path.join(os.path.dirname(sf), target)).replace("\\", "/")
                                        exists = abs_target in z.namelist()
                                        print(f"    abs={abs_target}, exists_in_zip={exists}")

                            raise SystemExit(0)
        except SystemExit:
            raise
        except Exception:
            pass
