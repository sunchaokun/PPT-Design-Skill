"""Test: does etree.tostring lose p:pic elements due to namespace issues?"""

import copy
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Minimal grpSp with pic
xml = f'''<p:grpSp xmlns:p="{_P}" xmlns:a="{_A}" xmlns:r="{_R}">
  <p:nvGrpSpPr><p:cNvPr id="1" name="G1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:grpSpPr>
  <p:pic>
    <p:nvPicPr><p:cNvPr id="2" name="Pic1"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
    <p:blipFill><a:blip r:embed="rId2" cstate="print"/></p:blipFill>
    <p:spPr>
      <a:xfrm rot="1164247" flipV="1"><a:off x="10" y="10"/><a:ext cx="50" cy="50"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:noFill/><a:ln><a:noFill/></a:ln>
      <a:scene3d><a:camera prst="isometricLeftDown"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>
    </p:spPr>
  </p:pic>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="3" name="Box1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="10" y="10"/><a:ext cx="50" cy="50"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:noFill/>
      <a:scene3d><a:camera prst="orthographicFront"><a:rot lat="19500000" lon="18900000" rev="3600000"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>
    </p:spPr>
    <p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>A</a:t></a:r></a:p></p:txBody>
  </p:sp>
</p:grpSp>'''

root = etree.fromstring(xml)
print(f"BEFORE: pic={len(list(root.iter(f'{{{_P}}}pic')))}, scene3d={len(list(root.iter(f'{{{_A}}}scene3d')))}, blip={len(list(root.iter(f'{{{_A}}}blip')))}")

# Simulate what _extract_from_xml_element does
theme_colors = {"bg1": "FFFFFF", "accent1": "4472C4"}

# Deep copy + resolve scheme colors
elem = copy.deepcopy(root)
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
        parent.remove(scheme)
        parent.insert(idx, srgb)

# Serialize
grp_xml = etree.tostring(elem, xml_declaration=False, encoding="unicode")
print(f"\nSerialized XML length: {len(grp_xml)}")

# Re-parse
root2 = etree.fromstring(grp_xml.encode("utf-8"))
print(f"AFTER roundtrip: pic={len(list(root2.iter(f'{{{_P}}}pic')))}, scene3d={len(list(root2.iter(f'{{{_A}}}scene3d')))}, blip={len(list(root2.iter(f'{{{_A}}}blip')))}")
print(f"Contains 'p:pic': {'p:pic' in grp_xml}")
print(f"Contains 'scene3d': {'scene3d' in grp_xml}")
print(f"Contains 'r:embed': {'r:embed' in grp_xml}")
