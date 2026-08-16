"""Regression tests for _resolve_svg_color return contract.

Bug: `_resolve_svg_color` was annotated `-> str` but returned `None` for
"none" / empty input. Downstream callers rely on `None` being falsy
(e.g. `_compiler.py: fill_hex or "#FFFFFF"`, `_add_freeform`'s
`fill is None` / `stroke=None`). The fix keeps the runtime behaviour —
it returns `None` for the no-color case (never raises) — and corrects the
type annotation to `str | None` so tooling/readers know None is expected.

These tests LOCK that behaviour so nobody "fixes" it into a `"none"` string,
which would silently turn fills/strokes black (truthy "none").
"""
import pytest

from ppt_pro_max.renderer.svg_compiler import SVGCompileError
from ppt_pro_max.renderer.svg_compiler._compiler import _resolve_svg_color


class TestResolveSvgColorNoneContract:
    def test_none_input_returns_none(self):
        assert _resolve_svg_color(None, {}, "#000000") is None

    def test_empty_string_returns_none(self):
        assert _resolve_svg_color("", {}, "#000000") is None

    def test_none_keyword_returns_none(self):
        assert _resolve_svg_color("none", {}, "#000000") is None

    def test_whitespace_only_returns_none(self):
        assert _resolve_svg_color("   ", {}, "#000000") is None

    def test_none_is_falsy_for_downstream(self):
        """Callers rely on None being falsy (fill_hex or '#FFFFFF')."""
        assert not _resolve_svg_color("none", {}, "#FFFFFF")


class TestResolveSvgColorStillResolves:
    def test_hex(self):
        assert _resolve_svg_color("#FF0000", {}, "#000000") == "#FF0000"

    def test_var_token(self):
        assert _resolve_svg_color("var(--primary)", {"primary": "#1D78FA"}, "#000000") == "#1D78FA"

    def test_named_color(self):
        assert _resolve_svg_color("red", {}, "#000000") == "#FF0000"

    def test_unresolved_var_raises(self):
        with pytest.raises(SVGCompileError, match="unresolved"):
            _resolve_svg_color("var(--nope)", {}, "#000000")


class TestAnnotationContract:
    def test_none_return_is_documented(self):
        import inspect
        sig = inspect.signature(_resolve_svg_color)
        # return annotation must admit None
        ann = sig.return_annotation
        assert None in getattr(ann, "__args__", (None,)), (
            f"return annotation {ann!r} must be Optional/str|None"
        )
