"""Find 3D components in DEF and compare original vs extracted XML."""

import os
import zipfile
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

def_dir = r"E:\BaiduNetdiskDownload\DEF"

# Find a 3D component in DEF
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
                    for grp in sp_tree.iter(f"{{{_P}}}grpSp"):
                        t_count = len([t for t in grp.iter(f"{{{_A}}}t") if t.text and t.text.strip()])
                        has_3d = len(list(grp.iter(f"{{{_A}}}scene3d"))) > 0
                        if t_count >= 10 and has_3d:
                            # Print the FULL grpSp XML (first 5000 chars)
                            grp_str = etree.tostring(grp, encoding="unicode", pretty_print=True)
                            print(f"=== ORIGINAL grpSp from {fname}/{sf} ===")
                            print(f"Total XML length: {len(grp_str)} chars")
                            print(f"Texts: {t_count}, Has 3D: True")
                            print(grp_str[:5000])
                            raise SystemExit(0)
        except SystemExit:
            raise
        except Exception:
            pass

print("No 3D component found")
