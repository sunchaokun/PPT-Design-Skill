"""Check component source file in DB."""

import sqlite3
import json

db_path = r"E:\PPT-Design-Skill\component_library\index.db"
conn = sqlite3.connect(db_path)
cur = conn.cursor()

cur.execute("SELECT id, category, node_count, meta FROM components WHERE category = 'hierarchy' LIMIT 3")
for row in cur.fetchall():
    meta = json.loads(row[3]) if row[3] else {}
    print(f"id={row[0]}, cat={row[1]}, nodes={row[2]}, source={meta.get('source_file', '?')}")

# Also check the build_library log for source mapping
cur.execute("SELECT id, category, node_count, meta FROM components WHERE id = 5751")
row = cur.fetchone()
if row:
    meta = json.loads(row[3]) if row[3] else {}
    print(f"\nTarget component id=5751:")
    print(f"  source={meta.get('source_file', '?')}")
    print(f"  meta keys={list(meta.keys())}")

conn.close()
