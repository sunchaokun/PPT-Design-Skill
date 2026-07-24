"""Compare original PPT XML vs extracted XML for a broken component."""

import os
import zipfile
from lxml import etree
from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

lib = ComponentLibrary(find_db_path())

# Get the hierarchy component's source info
results = lib.search(type="group", category="hierarchy", limit=1)
comp = results[0]
print(f"Component: id={comp['id']}, category={comp['category']}, nodes={comp['node_count']}")

# Get extracted XML
xml_parts = lib.load_xml(comp["id"])
extracted_root = etree.fromstring(xml_parts["group"])

# Find first shape with text
for sp in extracted_root.iter(f"{{{_P}}}sp"):
    texts = [t.text.strip() for t in sp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
    if texts:
        sp_str = etree.tostring(sp, encoding="unicode", pretty_print=True)
        print(f"\n=== EXTRACTED shape (text={texts[0][:10]}): ===")
        print(sp_str[:1500])
        break

# Now find the same component in the original PPT
# We need to find which PPT file this came from
# Let's search DEF directory
def_dir = r"E:\BaiduNetdiskDownload\DEF"
pptx_files = [f for f in os.listdir(def_dir) if f.endswith(".pptx")]

# Search for a PPT that has a hierarchy-like grpSp with 102 text nodes
for pf in pptx_files[:5]:
    pptx_path = os.path.join(def_dir, pf)
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = z.namelist()
            slide_files = [n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml") and "rels" not in n]
            for sf in slide_files[:3]:
                slide_xml = z.read(sf)
                slide_root = etree.fromstring(slide_xml)
                sp_tree = slide_root.find(f".//{{{_P}}}spTree")
                if sp_tree is None:
                    continue
                for grp in sp_tree.iter(f"{{{_P}}}grpSp"):
                    t_count = len([t for t in grp.iter(f"{{{_A}}}t") if t.text and t.text.strip()])
                    if t_count == 102:
                        # Found it!
                        for sp in grp.iter(f"{{{_P}}}sp"):
                            texts = [t.text.strip() for t in sp.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
                            if texts:
                                sp_str = etree.tostring(sp, encoding="unicode", pretty_print=True)
                                print(f"\n=== ORIGINAL shape from {pf}/{sf} (text={texts[0][:10]}): ===")
                                print(sp_str[:1500])
                                break
                        print(f"\nSource: {pf}")
                        break
                else:
                    continue
                break
            else:
                continue
            break
    except Exception as e:
        pass

lib.close()
