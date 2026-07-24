"""Test: does _resolve_scheme_colors remove pic/scene3d elements?"""

import copy
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

# Create a test grpSp with pic + scene3d + schemeClr
test_xml = f'''<p:grpSp xmlns:p="{_P}" xmlns:a="{_A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvGrpSpPr><p:cNvPr id="1" name="Group 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:grpSpPr>
  <p:pic>
    <p:nvPicPr><p:cNvPr id="2" name="Pic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
    <p:blipFill><a:blip r:embed="rId2"/></p:blipFill>
    <p:spPr>
      <a:xfrm><a:off x="10" y="10"/><a:ext cx="50" cy="50"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:noFill/>
      <a:scene3d><a:camera prst="isometricLeftDown"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>
    </p:spPr>
  </p:pic>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="3" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="10" y="10"/><a:ext cx="50" cy="50"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
      <a:scene3d><a:camera prst="isometricLeftDown"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>
    </p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>Test</a:t></a:r></a:p></p:txBody>
  </p:sp>
</p:grpSp>'''

root = etree.fromstring(test_xml)

# Before resolve
pic_before = len(list(root.iter(f"{{{_P}}}pic")))
scene3d_before = len(list(root.iter(f"{{{_A}}}scene3d")))
scheme_before = len(list(root.iter(f"{{{_A}}}schemeClr")))
blip_before = len(list(root.iter(f"{{{_A}}}blip")))
print(f"BEFORE: pic={pic_before}, scene3d={scene3d_before}, schemeClr={scheme_before}, blip={blip_before}")

# Simulate _resolve_scheme_colors
from ppt_pro_max.enterprise.group_extractor import GroupExtractor
ext = GroupExtractor()
theme_colors = {"accent1": "4472C4", "bg1": "FFFFFF"}

resolved = ext._resolve_scheme_colors(root, theme_colors)

pic_after = len(list(resolved.iter(f"{{{_P}}}pic")))
scene3d_after = len(list(resolved.iter(f"{{{_A}}}scene3d")))
scheme_after = len(list(resolved.iter(f"{{{_A}}}schemeClr")))
blip_after = len(list(resolved.iter(f"{{{_A}}}blip")))
print(f"AFTER:  pic={pic_after}, scene3d={scene3d_after}, schemeClr={scheme_after}, blip={blip_after}")

# Check serialized XML
xml_str = etree.tostring(resolved, encoding="unicode")
print(f"\nXML length: {len(xml_str)}")
print(f"Contains <p:pic: {'p:pic' in xml_str}")
print(f"Contains scene3d: {'scene3d' in xml_str}")
print(f"Contains r:embed: {'r:embed' in xml_str}")
