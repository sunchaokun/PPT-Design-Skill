"""Check raw stored XML for 3D elements."""

import gzip
from lxml import etree

_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

path = r"E:\PPT-Design-Skill\component_library\storage\group_process__24_8eea540f\group.xml.gz"
with gzip.open(path, "rb") as f:
    data = f.read()

print(f"scene3d in raw: {b'scene3d' in data}")
print(f"p:pic in raw: {b'p:pic' in data}")
print(f"blip in raw: {b'blip' in data}")

root = etree.fromstring(data)
tags = set()
for elem in root.iter():
    tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
    tags.add(tag)
print(f"All tags: {sorted(tags)}")

# Check spPr children (what fills are used)
fill_types = set()
for spPr in root.iter(f"{{{_P}}}spPr"):
    for child in spPr:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        fill_types.add(tag)
print(f"spPr child tags: {sorted(fill_types)}")
