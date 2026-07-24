"""Find best 2D components for each category (with fills, custGeom, many shapes)."""

from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

lib = ComponentLibrary(find_db_path())

for cat in ["hierarchy", "process", "swot", "infographic", "timeline", "chart"]:
    results = lib.search(type="group", category=cat, limit=50)
    scored = []
    for comp in results:
        xml_parts = lib.load_xml(comp["id"])
        root = etree.fromstring(xml_parts["group"])

        sp_count = len(list(root.iter(f"{{{_P}}}sp")))
        grp_count = len(list(root.iter(f"{{{_P}}}grpSp")))
        cxn_count = len(list(root.iter(f"{{{_P}}}cxnSp")))
        custGeom = len(list(root.iter(f"{{{_A}}}custGeom")))

        solidFill = 0
        gradFill = 0
        noFill = 0
        for spPr in root.iter(f"{{{_A}}}spPr"):
            parent_tag = spPr.getparent().tag.split("}")[-1] if spPr.getparent() is not None else ""
            if parent_tag == "sp":
                if spPr.find(f"{{{_A}}}solidFill") is not None:
                    solidFill += 1
                elif spPr.find(f"{{{_A}}}gradFill") is not None:
                    gradFill += 1
                elif spPr.find(f"{{{_A}}}noFill") is not None:
                    noFill += 1

        fill_score = solidFill * 2 + gradFill * 3 - noFill * 1 + custGeom * 2 + cxn_count * 2
        img_keys = [k for k in xml_parts if k.startswith("img_")]

        scored.append((fill_score, comp["id"], comp["node_count"], sp_count, grp_count, solidFill, gradFill, noFill, custGeom, len(img_keys)))

    scored.sort(key=lambda x: -x[0])
    print(f"\n{cat} (top 5):")
    for s, cid, nc, sp, grp, sf, gf, nf, cg, imgs in scored[:5]:
        print(f"  id={cid} score={s} nodes={nc} sp={sp} grp={grp} solidFill={sf} gradFill={gf} noFill={nf} custGeom={cg} imgs={imgs}")

lib.close()
