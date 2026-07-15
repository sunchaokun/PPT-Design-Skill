"""ComponentLibrary — SQLite-indexed SmartArt/Group/OLE template library.

Manages tens of thousands of component templates with:
- SQLite index for fast search by type/category/node_count/tags
- File system storage for XML template parts (gzip compressed)
- Match engine to find best component for extracted data
- Bulk import from PPT files (batch transaction, skip empty groups)
- Checksum-based deduplication

Key constraints from R&D:
- SmartArt stores 4 XML parts: data, layout, colors, quickStyle
- drawing.xml NOT stored (PowerPoint auto-rebuilds)
- colors.xml MUST store original (131 color mappings, can't generate from scratch)
- GroupShape stores full <p:grpSp> XML + image blobs
"""

from __future__ import annotations

import gzip
import hashlib
import json
import os
import sqlite3
from typing import Any


_CREATE_TABLES = """
CREATE TABLE IF NOT EXISTS components (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    type        TEXT NOT NULL,
    category    TEXT NOT NULL,
    variant     TEXT NOT NULL DEFAULT '',
    node_count  INTEGER NOT NULL DEFAULT 0,
    level_count INTEGER NOT NULL DEFAULT 0,
    tags        TEXT,
    xml_path    TEXT NOT NULL,
    thumb_path  TEXT,
    source      TEXT NOT NULL DEFAULT 'builtin',
    checksum    TEXT NOT NULL,
    created_at  TEXT DEFAULT (datetime('now')),
    updated_at  TEXT DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_type_category ON components(type, category);
CREATE INDEX IF NOT EXISTS idx_node_count ON components(node_count);
CREATE INDEX IF NOT EXISTS idx_type_category_nodes ON components(type, category, node_count);
CREATE INDEX IF NOT EXISTS idx_checksum ON components(checksum);
"""


class ComponentLibrary:

    def __init__(self, db_path: str = "component_library/index.db", storage_dir: str | None = None):
        self._db_path = db_path
        self._storage_dir = storage_dir or os.path.join(os.path.dirname(db_path), "storage")
        os.makedirs(self._storage_dir, exist_ok=True)
        os.makedirs(os.path.dirname(db_path), exist_ok=True)

        self._db = sqlite3.connect(db_path)
        self._db.row_factory = sqlite3.Row
        self._db.executescript(_CREATE_TABLES)
        self._catalog_cache: dict[str, Any] | None = None

    def catalog(self) -> dict[str, Any]:
        if self._catalog_cache is not None:
            return self._catalog_cache

        rows = self._db.execute(
            "SELECT type, category, COUNT(*) as cnt, MIN(node_count) as min_nc, MAX(node_count) as max_nc "
            "FROM components GROUP BY type, category ORDER BY type, category"
        ).fetchall()

        result: dict[str, Any] = {}
        for row in rows:
            comp_type = row["type"]
            category = row["category"]
            if comp_type not in result:
                result[comp_type] = {}
            result[comp_type][category] = {
                "count": row["cnt"],
                "min_nodes": row["min_nc"],
                "max_nodes": row["max_nc"],
            }

        self._catalog_cache = result
        return result

    def add(
        self,
        type: str,
        category: str,
        variant: str = "",
        node_count: int = 0,
        level_count: int = 0,
        tags: list[str] | None = None,
        xml_parts: dict[str, bytes] | None = None,
        image_blobs: dict[str, bytes] | None = None,
        thumb: bytes | None = None,
        source: str = "builtin",
        _skip_commit: bool = False,
        **meta,
    ) -> int:
        if xml_parts is None:
            xml_parts = {}
        if image_blobs is None:
            image_blobs = {}

        checksum = self._compute_checksum(type, category, variant, node_count, xml_parts)

        existing = self._db.execute(
            "SELECT id FROM components WHERE checksum = ?", (checksum,)
        ).fetchone()
        if existing:
            return existing["id"]

        xml_dir_name = f"{type}_{category}_{variant}_{node_count}_{checksum[:8]}"
        xml_dir_name = xml_dir_name.replace(" ", "_").replace("/", "_")
        xml_dir = os.path.join(self._storage_dir, xml_dir_name)
        os.makedirs(xml_dir, exist_ok=True)

        for part_name, part_data in xml_parts.items():
            part_path = os.path.join(xml_dir, f"{part_name}.xml.gz")
            raw = part_data if isinstance(part_data, bytes) else part_data.encode("utf-8")
            with gzip.open(part_path, "wb") as f:
                f.write(raw)

        img_dir = os.path.join(xml_dir, "images")
        if image_blobs:
            os.makedirs(img_dir, exist_ok=True)
        for rId, blob in image_blobs.items():
            safe_name = rId.replace("/", "_").replace("\\", "_")
            img_path = os.path.join(img_dir, f"{safe_name}.bin")
            with open(img_path, "wb") as f:
                f.write(blob)

        thumb_path = None
        if thumb:
            thumb_path = os.path.join(xml_dir, "thumb.png")
            with open(thumb_path, "wb") as f:
                f.write(thumb)

        tags_json = json.dumps(tags or [], ensure_ascii=False)

        cursor = self._db.execute(
            """INSERT INTO components (type, category, variant, node_count, level_count, tags, xml_path, thumb_path, source, checksum)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (type, category, variant, node_count, level_count, tags_json, xml_dir, thumb_path, source, checksum),
        )
        if not _skip_commit:
            self._db.commit()
        self._catalog_cache = None
        return cursor.lastrowid

    def get(self, component_id: int) -> dict[str, Any] | None:
        row = self._db.execute(
            "SELECT * FROM components WHERE id = ?", (component_id,)
        ).fetchone()
        if row is None:
            return None
        return self._row_to_dict(row)

    def search(
        self,
        type: str,
        category: str,
        node_count: int | None = None,
        tags: list[str] | None = None,
        limit: int = 10,
    ) -> list[dict[str, Any]]:
        query = "SELECT * FROM components WHERE type = ? AND category = ?"
        params: list[Any] = [type, category]

        if node_count is not None:
            query += " AND node_count = ?"
            params.append(node_count)

        if tags:
            for tag in tags:
                query += " AND tags LIKE ?"
                params.append(f"%{tag}%")

        query += " ORDER BY node_count DESC, variant LIMIT ?"
        params.append(limit)

        rows = self._db.execute(query, params).fetchall()
        return [self._row_to_dict(r) for r in rows]

    def match(self, extracted: dict) -> dict[str, Any] | None:
        comp_type = extracted.get("type", "")
        category = extracted.get("category", "")
        node_count = extracted.get("node_count", 0)

        if not comp_type or not category:
            return None

        results = self.search(type=comp_type, category=category, node_count=node_count, limit=5)
        if results:
            return results[0]

        results = self.search(type=comp_type, category=category, limit=5)
        if results:
            closest = min(results, key=lambda r: abs(r["node_count"] - node_count))
            return closest

        return None

    def load_xml(self, component_id: int) -> dict[str, bytes] | None:
        comp = self.get(component_id)
        if comp is None:
            return None

        xml_dir = comp["xml_path"]
        if not os.path.isdir(xml_dir):
            return None

        parts: dict[str, bytes] = {}

        for part_name in ("data", "layout", "colors", "quickStyle", "group"):
            gz_path = os.path.join(xml_dir, f"{part_name}.xml.gz")
            raw_path = os.path.join(xml_dir, f"{part_name}.xml")
            if os.path.isfile(gz_path):
                with gzip.open(gz_path, "rb") as f:
                    parts[part_name] = f.read()
            elif os.path.isfile(raw_path):
                with open(raw_path, "rb") as f:
                    parts[part_name] = f.read()

        if comp["type"] == "group":
            for fname in os.listdir(xml_dir):
                if fname.endswith(".xml.gz") and fname[:-3] not in ("group.xml",):
                    with gzip.open(os.path.join(xml_dir, fname), "rb") as f:
                        parts[fname[:-7]] = f.read()
                elif fname.endswith(".xml") and fname not in ("group.xml",):
                    with open(os.path.join(xml_dir, fname), "rb") as f:
                        parts[fname[:-4]] = f.read()

        img_dir = os.path.join(xml_dir, "images")
        if os.path.isdir(img_dir):
            for fname in os.listdir(img_dir):
                if fname.endswith(".bin"):
                    rId = fname[:-4]
                    with open(os.path.join(img_dir, fname), "rb") as f:
                        parts[f"img_{rId}"] = f.read()

        return parts

    def load_thumbnail(self, component_id: int) -> bytes | None:
        comp = self.get(component_id)
        if comp is None or comp.get("thumb_path") is None:
            return None
        thumb_path = comp["thumb_path"]
        if not os.path.isfile(thumb_path):
            return None
        with open(thumb_path, "rb") as f:
            return f.read()

    def remove(self, component_id: int) -> bool:
        comp = self.get(component_id)
        if comp is None:
            return False

        xml_dir = comp["xml_path"]
        if os.path.isdir(xml_dir):
            import shutil
            shutil.rmtree(xml_dir, ignore_errors=True)

        self._db.execute("DELETE FROM components WHERE id = ?", (component_id,))
        self._db.commit()
        self._catalog_cache = None
        return True

    def bulk_import(self, pptx_paths: list[str], min_node_count: int = 1) -> dict[str, int]:
        from ppt_pro_max.enterprise.smartart_extractor import SmartArtExtractor
        from ppt_pro_max.enterprise.group_extractor import GroupExtractor

        added = 0
        skipped = 0
        errors = 0

        sa_ext = SmartArtExtractor()
        grp_ext = GroupExtractor()

        self._db.execute("BEGIN")
        try:
            for pptx_path in pptx_paths:
                try:
                    sa_results = sa_ext.extract_all(pptx_path)
                    for r in sa_results:
                        try:
                            sa_nc = len(r.get("nodes", []))
                            if sa_nc < min_node_count:
                                skipped += 1
                                continue
                            cid = self.add(
                                type=r["type"],
                                category=r.get("category", "process"),
                                variant=r.get("variant", ""),
                                node_count=sa_nc,
                                xml_parts=r.get("xml_parts", {}),
                                _skip_commit=True,
                            )
                            if cid:
                                added += 1
                            else:
                                skipped += 1
                        except Exception:
                            skipped += 1

                    grp_results = grp_ext.extract_all(pptx_path)
                    for r in grp_results:
                        try:
                            if r.get("node_count", 0) < min_node_count:
                                skipped += 1
                                continue
                            cid = self.add(
                                type=r["type"],
                                category=r.get("category", "infographic"),
                                variant=r.get("variant", ""),
                                node_count=r.get("node_count", 0),
                                level_count=r.get("level_count", 0),
                                xml_parts=r.get("xml_parts", {}),
                                image_blobs=r.get("image_blobs", {}),
                                _skip_commit=True,
                            )
                            if cid:
                                added += 1
                            else:
                                skipped += 1
                        except Exception:
                            skipped += 1
                except Exception:
                    errors += 1

            self._db.commit()
        except Exception:
            self._db.rollback()
            raise

        return {"added": added, "skipped": skipped, "errors": errors}

    def stats(self) -> dict[str, int]:
        rows = self._db.execute(
            "SELECT type, COUNT(*) as cnt FROM components GROUP BY type"
        ).fetchall()
        result = {"total": 0}
        for row in rows:
            result[row["type"]] = row["cnt"]
            result["total"] += row["cnt"]
        return result

    def coverage(self, category: str) -> dict[str, int]:
        rows = self._db.execute(
            "SELECT node_count, COUNT(*) as cnt FROM components WHERE category = ? GROUP BY node_count ORDER BY node_count",
            (category,),
        ).fetchall()
        return {str(row["node_count"]): row["cnt"] for row in rows}

    def _compute_checksum(self, type: str, category: str, variant: str, node_count: int, xml_parts: dict) -> str:
        h = hashlib.md5()
        h.update(f"{type}:{category}:{variant}:{node_count}".encode())
        for key in sorted(xml_parts.keys()):
            val = xml_parts[key]
            if isinstance(val, bytes):
                h.update(val[:4096])
            else:
                h.update(val[:4096].encode("utf-8"))
        return h.hexdigest()

    def _row_to_dict(self, row) -> dict[str, Any]:
        d = dict(row)
        if d.get("tags"):
            try:
                d["tags"] = json.loads(d["tags"])
            except (json.JSONDecodeError, TypeError):
                d["tags"] = []
        return d

    def close(self):
        self._db.close()

    def __del__(self):
        try:
            self._db.close()
        except Exception:
            pass


def find_db_path(component_library: str | None = None, project_dir: str | None = None) -> str | None:
    if component_library and os.path.exists(component_library):
        return component_library
    if project_dir:
        p = os.path.join(project_dir, "component_library", "index.db")
        if os.path.exists(p):
            return p
    pkg_dir = os.path.dirname(os.path.abspath(__file__))
    for parent in (pkg_dir, os.path.dirname(pkg_dir), os.path.dirname(os.path.dirname(pkg_dir))):
        p = os.path.join(parent, "component_library", "index.db")
        if os.path.exists(p):
            return p
    return None
