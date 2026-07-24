"""Debug _extract_from_xml_element to find where pic/scene3d are lost."""

import os
import zipfile
import copy
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_rel = "http://schemas.openxmlformats.org/package/2006/relationships"

def_dir = r"E:\BaiduNetdiskDownload\DEF"

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
                        pic_count = len(list(grp.iter(f"{{{_P}}}pic")))
                        if t_count >= 10 and has_3d and pic_count > 0:
                            print(f"FOUND: {fname}/{sf}, texts={t_count}")

                            # Step 1: Before deepcopy
                            print(f"  Before deepcopy: pic={len(list(grp.iter(f'{{{_P}}}pic')))}, scene3d={len(list(grp.iter(f'{{{_A}}}scene3d')))}, blip={len(list(grp.iter(f'{{{_A}}}blip')))}")

                            # Step 2: After deepcopy
                            elem = copy.deepcopy(grp)
                            print(f"  After deepcopy: pic={len(list(elem.iter(f'{{{_P}}}pic')))}, scene3d={len(list(elem.iter(f'{{{_A}}}scene3d')))}, blip={len(list(elem.iter(f'{{{_A}}}blip')))}")

                            # Step 3: After resolve scheme colors
                            theme_colors = {"bg1": "FFFFFF", "accent1": "4472C4", "tx1": "000000"}
                            for scheme in list(elem.iter(f"{{{_A}}}schemeClr")):
                                parent = scheme.getparent()
                                if parent is None:
                                    continue
                                val = scheme.get("val", "")
                                if val in theme_colors:
                                    idx = list(parent).index(scheme)
                                    srgb = etree.SubElement(parent, f"{{{_A}}}srgbClr")
                                    srgb.set("val", theme_colors[val])
                                    for child in scheme:
                                        srgb.append(copy.deepcopy(child))
                                    parent.remove(srgb)
                                    parent.insert(idx, srgb)
                            print(f"  After resolve: pic={len(list(elem.iter(f'{{{_P}}}pic')))}, scene3d={len(list(elem.iter(f'{{{_A}}}scene3d')))}, blip={len(list(elem.iter(f'{{{_A}}}blip')))}")

                            # Step 4: After tostring
                            grp_xml = etree.tostring(elem, xml_declaration=False, encoding="unicode")
                            print(f"  After tostring: len={len(grp_xml)}, contains_pic={'p:pic' in grp_xml}, contains_scene3d={'scene3d' in grp_xml}, contains_blip={'blip' in grp_xml}")

                            # Step 5: Re-parse
                            root2 = etree.fromstring(grp_xml.encode("utf-8"))
                            print(f"  After reparse: pic={len(list(root2.iter(f'{{{_P}}}pic')))}, scene3d={len(list(root2.iter(f'{{{_A}}}scene3d')))}, blip={len(list(root2.iter(f'{{{_A}}}blip')))}")

                            raise SystemExit(0)
        except SystemExit:
            raise
        except Exception:
            pass
