"""Compare XML structure of working vs broken components."""

from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

lib = ComponentLibrary(find_db_path())

# P1=hierarchy(bad), P2=process(ok), P3=swot(ok), P4=infographic(bad), P5=timeline(ok), P6=chart(bad)
categories = [
    ("hierarchy", "BAD"),
    ("process", "OK"),
    ("swot", "OK"),
    ("infographic", "BAD"),
    ("timeline", "OK"),
    ("chart", "BAD"),
]

for cat, status in categories:
    results = lib.search(type="group", category=cat, limit=1)
    if not results:
        continue
    comp = results[0]
    xml_parts = lib.load_xml(comp["id"])
    root = etree.fromstring(xml_parts["group"])

    print(f"\n{'='*60}")
    print(f"{cat} [{status}] (id={comp['id']}, nodes={comp['node_count']})")
    print(f"{'='*60}")

    # 1. Root tag
    print(f"Root tag: {root.tag.split('}')[-1]}")

    # 2. Direct children tags
    children = list(root)
    child_tags = []
    for ch in children:
        tag = ch.tag.split("}")[-1] if "}" in ch.tag else ch.tag
        child_tags.append(tag)
    from collections import Counter
    print(f"Direct children: {dict(Counter(child_tags))}")

    # 3. Nested grpSp depth
    max_depth = 0
    for gs in root.iter(f"{{{_P}}}grpSp"):
        depth = 0
        p = gs.getparent()
        while p is not None:
            if p.tag == gs.tag:
                depth += 1
            p = p.getparent()
        max_depth = max(max_depth, depth)
    print(f"Max grpSp nesting depth: {max_depth}")

    # 4. Total shapes and groups
    sp_count = len(list(root.iter(f"{{{_P}}}sp")))
    grp_count = len(list(root.iter(f"{{{_P}}}grpSp")))
    cxnSp_count = len(list(root.iter(f"{{{_P}}}cxnSp")))
    print(f"sp={sp_count}, grpSp={grp_count}, cxnSp={cxnSp_count}")

    # 5. Check for image references
    blip_count = 0
    for blip in root.iter(f"{{{_A}}}blip"):
        blip_count += 1
    print(f"blip (image) refs: {blip_count}")

    # 6. Check grpSpPr xfrm
    grpSpPr = root.find(f".//{{{_P}}}grpSpPr")
    if grpSpPr is not None:
        xfrm = grpSpPr.find(f".//{{{_A}}}xfrm")
        if xfrm is not None:
            off = xfrm.find(f"{{{_A}}}off")
            ext = xfrm.find(f"{{{_A}}}ext")
            chOff = xfrm.find(f"{{{_A}}}chOff")
            chExt = xfrm.find(f"{{{_A}}}chExt")
            if off is not None:
                print(f"Root xfrm: off=({off.get('x')},{off.get('y')}), ext=({ext.get('cx')},{ext.get('cy')})")
            if chOff is not None:
                print(f"  chOff=({chOff.get('x')},{chOff.get('y')}), chExt=({chExt.get('cx')},{chExt.get('cy')})")

    # 7. Check for missing nvGrpSpPr
    nvGrpSpPr = root.find(f".//{{{_P}}}nvGrpSpPr")
    print(f"Has nvGrpSpPr: {nvGrpSpPr is not None}")

    # 8. Check for empty/missing spPr in child shapes
    missing_spPr = 0
    for sp in root.iter(f"{{{_P}}}sp"):
        spPr = sp.find(f"{{{_P}}}spPr")
        if spPr is None:
            missing_spPr += 1
    if missing_spPr:
        print(f"Shapes missing spPr: {missing_spPr}")

    # 9. Check for custom geometry
    custGeom = 0
    for cg in root.iter(f"{{{_A}}}custGeom"):
        custGeom += 1
    if custGeom:
        print(f"Custom geometries: {custGeom}")

    # 10. Check for effect styles
    effectLst = 0
    for ef in root.iter(f"{{{_A}}}effectLst"):
        effectLst += 1
    if effectLst:
        print(f"Effect lists: {effectLst}")

lib.close()
