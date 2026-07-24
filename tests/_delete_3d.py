"""Delete all 3D components (with scene3d) from the library."""

from ppt_pro_max.enterprise.component_library import ComponentLibrary, find_db_path
from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

lib = ComponentLibrary(find_db_path())

deleted = 0
kept = 0
total = 0

cur = lib._db.execute("SELECT id, type, category, node_count FROM components WHERE type = 'group'")
rows = cur.fetchall()

for row in rows:
    cid = row["id"]
    cat = row["category"]
    nc = row["node_count"]
    xml_parts = lib.load_xml(cid)
    if not xml_parts or "group" not in xml_parts:
        continue
    root = etree.fromstring(xml_parts["group"])
    has_3d = len(list(root.iter(f"{{{_A}}}scene3d"))) > 0
    total += 1
    if has_3d:
        lib._db.execute("DELETE FROM components WHERE id = ?", (cid,))
        deleted += 1
    else:
        kept += 1

lib._db.commit()

stats = lib.stats()
print(f"Total checked: {total}")
print(f"Deleted (3D): {deleted}")
print(f"Kept (2D): {kept}")
print(f"Library now: {stats}")

lib.close()
