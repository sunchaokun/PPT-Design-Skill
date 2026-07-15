"""OLEExtractor — extract OLE/embedded object metadata from .pptx.

OLE objects (embedded Excel, Word, etc.) are stored in the .pptx ZIP
as oleObjectN.xml parts. This extractor reads metadata like progId
and checks for associated chart data.
"""

from __future__ import annotations

import os
import zipfile
from typing import Any

from lxml import etree


_NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


class OLEExtractor:

    def extract(self, pptx_path: str, slide_index: int, shape_rId: str) -> dict[str, Any] | None:
        try:
            with zipfile.ZipFile(pptx_path) as z:
                names = z.namelist()
                rels_path = f"ppt/slides/_rels/slide{slide_index + 1}.xml.rels"
                if rels_path not in names:
                    return None

                rels_xml = z.read(rels_path)
                rels_root = etree.fromstring(rels_xml)
                rel_ns = _NS["rel"]

                for rel in rels_root.findall(f"{{{rel_ns}}}Relationship"):
                    rid = rel.get("Id", "")
                    if rid != shape_rId:
                        continue
                    target = rel.get("Target", "")
                    rel_type = rel.get("Type", "")

                    if "oleObject" in target or "embed" in rel_type.lower():
                        abs_dir = os.path.normpath(f"ppt/slides/{os.path.dirname(target)}").replace("\\", "/")
                        fname = os.path.basename(target)
                        abs_path = f"{abs_dir}/{fname}"

                        prog_id = ""
                        if abs_path in names:
                            ole_xml = z.read(abs_path)
                            ole_root = etree.fromstring(ole_xml)
                            prog_id = ole_root.get("progId", ole_root.get("ProgID", ""))

                        has_chart = False
                        chart_data = None
                        for chart_name in names:
                            if f"slide{slide_index + 1}" in chart_name and "chart" in chart_name.lower():
                                has_chart = True
                                break

                        return {
                            "type": "ole",
                            "prog_id": prog_id,
                            "has_chart": has_chart,
                            "chart_data": chart_data,
                        }
        except Exception:
            pass

        return None

    def extract_all(self, pptx_path: str) -> list[dict[str, Any]]:
        results: list[dict[str, Any]] = []
        try:
            with zipfile.ZipFile(pptx_path) as z:
                names = z.namelist()
                ole_files = [n for n in names if "oleObject" in n]
                embed_files = [n for n in names if "embed" in n.lower() and n.endswith(".xml")]

                for ole_path in ole_files + embed_files:
                    try:
                        ole_xml = z.read(ole_path)
                        ole_root = etree.fromstring(ole_xml)
                        prog_id = ole_root.get("progId", ole_root.get("ProgID", ""))

                        results.append({
                            "type": "ole",
                            "prog_id": prog_id,
                            "has_chart": "Excel" in prog_id or "Chart" in prog_id,
                            "chart_data": None,
                            "source_path": ole_path,
                        })
                    except Exception:
                        continue
        except Exception:
            pass

        return results
