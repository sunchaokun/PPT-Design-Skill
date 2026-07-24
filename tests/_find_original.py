"""Find the original PPT that contains the hierarchy component with 102 'A' text nodes."""

import os
import zipfile
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

def_dir = r"E:\BaiduNetdiskDownload\DEF"

found = 0
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
                        if t_count == 102:
                            # Check if texts are all 'A'
                            all_texts = [t.text.strip() for t in grp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
                            if all(t == "A" for t in all_texts[:5]):
                                print(f"FOUND: {fpath}")
                                # Print first shape
                                for sp in grp.iter(f"{{{_P}}}sp"):
                                    texts = [t.text.strip() for t in sp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
                                    if texts:
                                        sp_str = etree.tostring(sp, encoding="unicode", pretty_print=True)
                                        print(sp_str[:2000])
                                        found += 1
                                        if found >= 1:
                                            raise SystemExit(0)
        except SystemExit:
            raise
        except Exception:
            pass

print(f"Found: {found}")
