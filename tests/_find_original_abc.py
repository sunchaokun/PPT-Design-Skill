"""Find the original 3D hierarchy component in ABC directory and compare."""

import os
import zipfile
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

abc_dir = r"E:\BaiduNetdiskDownload\ABC\960页7套PPT高端高级架构图表"

# Search for a grpSp with 102 text nodes all "A" and scene3d
for fname in os.listdir(abc_dir):
    if not fname.endswith(".pptx"):
        continue
    fpath = os.path.join(abc_dir, fname)
    print(f"Checking: {fname} ({os.path.getsize(fpath)/1024/1024:.0f}MB)...")
    try:
        with zipfile.ZipFile(fpath) as z:
            slide_files = [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml") and "rels" not in n]
            for sf in slide_files[:10]:
                slide_root = etree.fromstring(z.read(sf))
                sp_tree = slide_root.find(f".//{{{_P}}}spTree")
                if sp_tree is None:
                    continue
                for grp in sp_tree.iter(f"{{{_P}}}grpSp"):
                    t_count = len([t for t in grp.iter(f"{{{_A}}}t") if t.text and t.text.strip()])
                    has_3d = len(list(grp.iter(f"{{{_A}}}scene3d"))) > 0
                    if t_count >= 20 and has_3d:
                        all_texts = [t.text.strip() for t in grp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
                        print(f"  Found: {sf}, texts={t_count}, 3d=True, sample={all_texts[:3]}")
                        # Print first shape
                        for sp in grp.iter(f"{{{_P}}}sp"):
                            texts = [t.text.strip() for t in sp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
                            if texts:
                                sp_str = etree.tostring(sp, encoding="unicode", pretty_print=True)
                                print(f"\n=== ORIGINAL shape ===")
                                print(sp_str[:3000])
                                raise SystemExit(0)
    except SystemExit:
        raise
    except Exception as e:
        print(f"  Error: {e}")
