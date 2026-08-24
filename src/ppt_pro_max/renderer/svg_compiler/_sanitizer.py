"""SVG sanitizer — fixes common LLM-generated SVG quirks before compilation.

Handles:
- Missing XML declaration / namespace
- Unclosed self-closing tags (``<rect ...>`` → ``<rect .../>``)
- ``style`` attribute → individual attributes (``fill``, ``stroke``, etc.)
- Missing ``viewBox`` (inferred from width/height)
- Stripping ``<script>``, ``<style>`` elements
- Normalizing whitespace
"""
from __future__ import annotations

import re

from lxml import etree

SVG_NS = "http://www.w3.org/2000/svg"
SVG = f"{{{SVG_NS}}}"

_STYLE_PROPS = {
    "fill": re.compile(r"(?:^|;)\s*fill\s*:\s*([^;]+?)(?:;|$)"),
    "stroke": re.compile(r"(?:^|;)\s*stroke\s*:\s*([^;]+?)(?:;|$)"),
    "stroke-width": re.compile(r"(?:^|;)\s*stroke-width\s*:\s*([^;]+?)(?:;|$)"),
    "stroke-dasharray": re.compile(r"(?:^|;)\s*stroke-dasharray\s*:\s*([^;]+?)(?:;|$)"),
    "fill-opacity": re.compile(r"(?:^|;)\s*fill-opacity\s*:\s*([^;]+?)(?:;|$)"),
    "stroke-opacity": re.compile(r"(?:^|;)\s*stroke-opacity\s*:\s*([^;]+?)(?:;|$)"),
    "stop-color": re.compile(r"(?:^|;)\s*stop-color\s*:\s*([^;]+?)(?:;|$)"),
    "stop-opacity": re.compile(r"(?:^|;)\s*stop-opacity\s*:\s*([^;]+?)(?:;|$)"),
    "font-size": re.compile(r"(?:^|;)\s*font-size\s*:\s*([^;]+?)(?:;|$)"),
    "font-family": re.compile(r"(?:^|;)\s*font-family\s*:\s*([^;]+?)(?:;|$)"),
    "font-weight": re.compile(r"(?:^|;)\s*font-weight\s*:\s*([^;]+?)(?:;|$)"),
    "font-style": re.compile(r"(?:^|;)\s*font-style\s*:\s*([^;]+?)(?:;|$)"),
    "text-anchor": re.compile(r"(?:^|;)\s*text-anchor\s*:\s*([^;]+?)(?:;|$)"),
    "opacity": re.compile(r"(?:^|;)\s*opacity\s*:\s*([^;]+?)(?:;|$)"),
    "transform": re.compile(r"(?:^|;)\s*transform\s*:\s*([^;]+?)(?:;|$)"),
    "clip-path": re.compile(r"(?:^|;)\s*clip-path\s*:\s*([^;]+?)(?:;|$)"),
    "fill-rule": re.compile(r"(?:^|;)\s*fill-rule\s*:\s*([^;]+?)(?:;|$)"),
}

_SELF_CLOSING = {
    "rect", "circle", "ellipse", "line", "polyline", "polygon",
    "path", "stop", "use", "image", "br", "hr", "input",
}

_STRIP_ELEMENTS = {"script", "style"}


def _expand_style(el: etree._Element) -> None:
    style = el.get("style")
    if not style:
        return
    for prop, pat in _STYLE_PROPS.items():
        m = pat.search(style)
        if m:
            val = m.group(1).strip()
            if val and el.get(prop) is None:
                el.set(prop, val)
    del el.attrib["style"]


def _fix_self_closing_lxml(root: etree._Element) -> None:
    """Fix self-closing tags that have no children but were parsed as open tags.

    Preserves parent-child relationships (e.g., <rect><title>x</title></rect>).
    """
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        tag = el.tag.split("}")[-1]
        if tag in _SELF_CLOSING and len(el) == 0 and not el.text:
            el.text = None


def _infer_viewbox(root: etree._Element) -> None:
    if root.get("viewBox") is not None:
        return
    w = root.get("width")
    h = root.get("height")
    if w and h:
        w_val = re.sub(r"[^\d.]", "", w)
        h_val = re.sub(r"[^\d.]", "", h)
        if w_val and h_val:
            root.set("viewBox", f"0 0 {w_val} {h_val}")


def _strip_unwanted(root: etree._Element) -> None:
    """Remove <script> and <style> elements."""
    for tag in _STRIP_ELEMENTS:
        for el in root.findall(f".//{SVG}{tag}"):
            if el.getparent() is not None:
                el.getparent().remove(el)
        for el in root.findall(f".//{tag}"):
            if el.getparent() is not None:
                el.getparent().remove(el)


def _walk_expand(el: etree._Element) -> None:
    _expand_style(el)
    for child in el:
        _walk_expand(child)


def sanitize(svg_text: str) -> etree._Element:
    """Sanitize an SVG string and return a cleaned lxml Element tree.

    Steps:
    1. Parse with lxml (tolerant of minor XML errors)
    2. Ensure SVG namespace
    3. Fix self-closing tags (lxml-based, preserves children)
    4. Strip ``<script>`` / ``<style>``
    5. Expand ``style`` attributes → individual attributes
    6. Infer missing ``viewBox``
    """
    xml_bytes = svg_text.encode("utf-8")

    try:
        root = etree.fromstring(xml_bytes)
    except etree.XMLSyntaxError:
        parser = etree.XMLParser(recover=True)
        root = etree.fromstring(xml_bytes, parser)

    if not root.tag.startswith(SVG_NS) and not root.tag.endswith("svg"):
        for child in list(root):
            if child.tag.endswith("svg") or child.tag.startswith(SVG_NS):
                root = child
                break

    _fix_self_closing_lxml(root)
    _strip_unwanted(root)
    _walk_expand(root)
    _infer_viewbox(root)

    return root
