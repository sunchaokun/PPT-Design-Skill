"""Shared slide utilities for python-pptx."""

from __future__ import annotations

from pptx import Presentation


def remove_slide(prs: Presentation, index: int) -> None:
    sldIdLst = prs.slides._sldIdLst
    sldId_entry = sldIdLst[index]
    rId = sldId_entry.get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    )
    prs.part.drop_rel(rId)
    sldIdLst.remove(sldId_entry)
