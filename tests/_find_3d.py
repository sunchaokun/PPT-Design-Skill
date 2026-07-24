"""Extract a hierarchy component directly from source PPT and compare with library version."""

import os
import zipfile
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

def_dir = r"E:\BaiduNetdiskDownload\DEF"

# Find a PPT with hierarchy grpSp that has scene3d
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
                        has_3d = any(grp.iter(f"{{{_A}}}scene3d"))
                        t_count = len([t for t in grp.iter(f"{{{_A}}}t") if t.text and t.text.strip()])
                        if has_3d and t_count >= 20:
                            print(f"FOUND 3D component: {fpath}")
                            print(f"  slide: {sf}, texts: {t_count}")

                            # Print first shape with text
                            for sp in grp.iter(f"{{{_P}}}sp"):
                                texts = [t.text.strip() for t in sp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
                                if texts:
                                    sp_str = etree.tostring(sp, encoding="unicode", pretty_print=True)
                                    print(f"\n=== ORIGINAL shape (text={texts[0][:10]}): ===")
                                    print(sp_str[:2000])
                                    raise SystemExit(0)
        except SystemExit:
            raise
        except Exception:
            pass

print("No 3D component found")
