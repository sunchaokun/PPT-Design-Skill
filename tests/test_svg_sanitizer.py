"""Tests for svg_compiler._sanitizer — LLM SVG quirks."""

from ppt_pro_max.renderer.svg_compiler._sanitizer import SVG, sanitize


class TestSanitizeBasic:
    def test_valid_svg(self):
        root = sanitize('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300"><rect x="0" y="0" width="100" height="100" fill="red"/></svg>')
        assert root is not None

    def test_missing_namespace(self):
        root = sanitize('<svg viewBox="0 0 400 300"><rect x="0" y="0" width="100" height="100"/></svg>')
        assert root is not None


class TestSanitizeStyleExpansion:
    def test_style_to_attrs(self):
        root = sanitize('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300"><rect x="0" y="0" width="100" height="100" style="fill:#FF0000;stroke:#000000;stroke-width:2"/></svg>')
        rects = root.findall(f".//{SVG}rect")
        if not rects:
            rects = root.findall(".//rect")
        assert len(rects) >= 1
        r = rects[0]
        assert r.get("fill") is not None or r.get("style") is None

    def test_style_does_not_override_existing_attr(self):
        root = sanitize('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300"><rect x="0" y="0" width="100" height="100" fill="blue" style="fill:red"/></svg>')
        rects = root.findall(f".//{SVG}rect")
        if not rects:
            rects = root.findall(".//rect")
        r = rects[0]
        assert r.get("fill") == "blue"


class TestSanitizeViewBox:
    def test_missing_viewbox_inferred(self):
        root = sanitize('<svg xmlns="http://www.w3.org/2000/svg" width="400" height="300"><rect x="0" y="0" width="100" height="100"/></svg>')
        assert root.get("viewBox") == "0 0 400 300"

    def test_existing_viewbox_preserved(self):
        root = sanitize('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 600"><rect x="0" y="0" width="100" height="100"/></svg>')
        assert root.get("viewBox") == "0 0 800 600"


class TestSanitizeStripUnwanted:
    def test_script_removed(self):
        root = sanitize('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300"><script>alert(1)</script><rect x="0" y="0" width="100" height="100"/></svg>')
        scripts = root.findall(f".//{SVG}script")
        if not scripts:
            scripts = root.findall(".//script")
        assert len(scripts) == 0

    def test_style_element_removed(self):
        root = sanitize('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300"><style>.cls{fill:red}</style><rect x="0" y="0" width="100" height="100"/></svg>')
        styles = root.findall(f".//{SVG}style")
        if not styles:
            styles = root.findall(".//style")
        assert len(styles) == 0


class TestSanitizeSelfClosing:
    def test_unclosed_rect(self):
        svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300"><rect x="0" y="0" width="100" height="100"></svg>'
        root = sanitize(svg)
        assert root is not None


class TestSanitizeBrokenXML:
    def test_recover_from_malformed(self):
        svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300"><rect x="0" y="0" width="100" height="100" fill="red"><circle cx="50" cy="50" r="20"/></svg>'
        root = sanitize(svg)
        assert root is not None
