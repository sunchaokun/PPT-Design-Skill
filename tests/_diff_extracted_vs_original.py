"""Compare original shape XML vs library-extracted XML to find what's missing."""

import os
import zipfile
from lxml import etree
from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

lib = ComponentLibrary(find_db_path())

# Get extracted hierarchy component
results = lib.search(type="group", category="hierarchy", limit=1)
comp = results[0]
xml_parts = lib.load_xml(comp["id"])
extracted_root = etree.fromstring(xml_parts["group"])

# Find first shape with text in extracted
print("=== EXTRACTED (from library) ===")
for sp in extracted_root.iter(f"{{{_P}}}sp"):
    texts = [t.text.strip() for t in sp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
    if texts:
        sp_str = etree.tostring(sp, encoding="unicode", pretty_print=True)
        print(sp_str)
        break

# Now find the SAME component in the original PPT
# The hierarchy with 102 nodes all "A" and scene3d
def_dir = r"E:\BaiduNetdiskDownload\DEF"
found = False
for root_dir, dirs, files in os.walk(def_dir):
    for fname in files:
        if not fname.endswith(".pptx") or found:
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
                        all_texts = [t.text.strip() for t in grp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
                        has_3d = len(list(grp.iter(f"{{{_A}}}scene3d"))) > 0
                        if t_count == 102 and has_3d and all(t == "A" for t in all_texts[:5]):
                            print(f"\n=== ORIGINAL (from {fname}) ===")
                            for sp in grp.iter(f"{{{_P}}}sp"):
                                texts = [t.text.strip() for t in sp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
                                if texts:
                                    sp_str = etree.tostring(sp, encoding="unicode", pretty_print=True)
                                    print(sp_str)
                                    found = True
                                    break
                            break
                    if found:
                        break
            if found:
                break
        except Exception:
            pass
    if found:
        break

if not found:
    print("Original not found in DEF, trying ABC...")
    abc_dir = r"E:\BaiduNetdiskDownload\ABC\960页7套PPT高端高级架构图表"
    for root_dir, dirs, files in os.walk(abc_dir):
        for fname in files:
            if not fname.endswith(".pptx") or found:
                continue
            fpath = os.path.join(root_dir, fname)
            try:
                with zipfile.ZipFile(fpath) as z:
                    for sf in [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml") and "rels" not in n][:5]:
                        slide_root = etree.fromstring(z.read(sf))
                        sp_tree = slide_root.find(f".//{{{_P}}}spTree")
                        if sp_tree is None:
                            continue
                        for grp in sp_tree.iter(f"{{{_P}}}grpSp"):
                            t_count = len([t for t in grp.iter(f"{{{_A}}}t") if t.text and t.text.strip()])
                            has_3d = len(list(grp.iter(f"{{{_A}}}scene3d"))) > 0
                            if t_count >= 20 and has_3d:
                                print(f"\n=== ORIGINAL 3D (from {fname}, t={t_count}) ===")
                                for sp in grp.iter(f"{{{_P}}}sp"):
                                    texts = [t.text.strip() for t in sp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
                                    if texts:
                                        sp_str = etree.tostring(sp, encoding="unicode", pretty_print=True)
                                        print(sp_str)
                                        found = True
                                        break
                                break
                        if found:
                            break
                if found:
                    break
            except Exception:
                pass
        if found:
            break

lib.close()
