import sqlite3
conn = sqlite3.connect(r"E:\PPT-Design-Skill\component_library\index.db")
cur = conn.cursor()
cur.execute("SELECT id, category, node_count, source FROM components WHERE id = 5751")
row = cur.fetchone()
if row:
    print(f"id={row[0]}, cat={row[1]}, nodes={row[2]}, source={row[3]}")
cur.execute("SELECT id, category, node_count, source FROM components WHERE category = 'hierarchy' LIMIT 5")
for row in cur.fetchall():
    print(f"id={row[0]}, cat={row[1]}, nodes={row[2]}, source={row[3][:80]}")
conn.close()
