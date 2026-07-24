"""Analyze the 3 broken components."""

from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

lib = ComponentLibrary(find_db_path())

for cid, cat in [(4864, "hierarchy"), (4092, "timeline"), (4439, "chart")]:
    xml_parts = lib.load_xml(cid)
    root = etree.fromstring(xml_parts["group"])
    img_keys = [k for k in xml_parts if k.startswith("img_")]

    sp_count = len(list(root.iter(f"{{{_P}}}sp")))
    grp_count = len(list(root.iter(f"{{{_P}}}grpSp")))
    pic_count = len(list(root.iter(f"{{{_P}}}pic")))
    cxn_count = len(list(root.iter(f"{{{_P}}}cxnSp")))

    # Check fill types
    noFill = 0
    solidFill = 0
    gradFill = 0
    for sp in root.iter(f"{{{_P}}}sp"):
        spPr = sp.find(f".//{{{_P}}}spPr")
        if spPr is not None:
            if spPr.find(f"{{{_A}}}noFill") is not None:
                noFill += 1
            if spPr.find(f"{{{_A}}}solidFill") is not None:
                solidFill += 1
            if spPr.find(f"{{{_A}}}gradFill") is not None:
                gradFill += 1

    # Check for custGeom
    custGeom = len(list(root.iter(f"{{{_A}}}custGeom")))

    texts = [t.text.strip()[:15] for t in root.iter(f"{{{_A}}}t") if t.text and t.text.strip()]

    print(f"\n{cat} (id={cid}):")
    print(f"  sp={sp_count}, grpSp={grp_count}, pic={pic_count}, cxnSp={cxn_count}")
    print(f"  noFill={noFill}, solidFill={solidFill}, gradFill={gradFill}")
    print(f"  custGeom={custGeom}, img_keys={img_keys}")
    print(f"  texts ({len(texts)}): {texts[:6]}")

    # Sample first shape with noFill
    for sp in root.iter(f"{{{_P}}}sp"):
        spPr = sp.find(f".//{{{_P}}}spPr")
        if spPr is not None and spPr.find(f"{{{_A}}}noFill") is not None:
            sp_str = etree.tostring(sp, encoding="unicode", pretty_print=True)
            print(f"  Sample noFill shape ({len(sp_str)} chars):")
            print(f"  {sp_str[:300]}")
            break

lib.close()
