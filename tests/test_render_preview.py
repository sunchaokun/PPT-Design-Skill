"""Tests for render_preview — pptx → PNG preview tooling.

Backend engines are mocked so tests run on any machine (CI included).
"""
from __future__ import annotations

from unittest.mock import patch

import pytest

from ppt_pro_max.render_preview import (
    _build_html,
    _find_libreoffice,
    detect_engine,
    render_preview,
)


# ── fixture: a tiny real pptx via python-pptx ──────────────────────────────
@pytest.fixture
def sample_pptx(tmp_path):
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    for _ in range(3):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        tb.text_frame.text = "preview test slide"
    path = tmp_path / "deck.pptx"
    prs.save(str(path))
    return path


def _make_pngs(d, n):
    paths = []
    for i in range(1, n + 1):
        p = d / f"slide{i}.png"
        p.write_bytes(b"PNG")
        paths.append(p)
    return paths


# ── detect_engine ──────────────────────────────────────────────────────────
def test_detect_engine_none_when_no_backend():
    with patch("ppt_pro_max.render_preview._powerpoint_available", return_value=False), \
         patch("ppt_pro_max.render_preview._find_libreoffice", return_value=None):
        assert detect_engine() == "none"


def test_detect_engine_powerpoint_when_available():
    with patch("ppt_pro_max.render_preview._powerpoint_available", return_value=True):
        assert detect_engine() == "powerpoint"


def test_detect_engine_libreoffice_when_available():
    with patch("ppt_pro_max.render_preview._powerpoint_available", return_value=False), \
         patch("ppt_pro_max.render_preview._find_libreoffice", return_value="/usr/bin/soffice"):
        assert detect_engine() == "libreoffice"


def test_find_libreoffice_none_when_absent():
    with patch("ppt_pro_max.render_preview.shutil.which", return_value=None), \
         patch("ppt_pro_max.render_preview.os.path.isfile", return_value=False):
        assert _find_libreoffice() is None


# ── render_preview ─────────────────────────────────────────────────────────


def test_render_preview_powerpoint(sample_pptx, tmp_path):
    out = tmp_path / "preview"
    with patch("ppt_pro_max.render_preview.detect_engine", return_value="powerpoint"), \
         patch("ppt_pro_max.render_preview._render_powerpoint",
               side_effect=lambda p, d, w, h: _make_pngs(d, 3)):
        result = render_preview(str(sample_pptx), out_dir=str(out))
    assert result["engine"] == "powerpoint"
    assert len(result["pngs"]) == 3
    assert all(p.exists() for p in result["pngs"])
    assert result["html"].exists()


def test_render_preview_default_out_dir(sample_pptx):
    expected = sample_pptx.parent / "preview" / sample_pptx.stem
    with patch("ppt_pro_max.render_preview.detect_engine", return_value="powerpoint"), \
         patch("ppt_pro_max.render_preview._render_powerpoint",
               side_effect=lambda p, d, w, h: [d / "slide1.png"]):
        result = render_preview(str(sample_pptx))
    assert result["pngs"][0].parent == expected


def test_render_preview_missing_file(tmp_path):
    with pytest.raises(FileNotFoundError):
        render_preview(str(tmp_path / "nope.pptx"))


def test_render_preview_no_engine(sample_pptx):
    with patch("ppt_pro_max.render_preview.detect_engine", return_value="none"), \
         pytest.raises(RuntimeError, match="No rendering engine"):
        render_preview(str(sample_pptx))


def test_render_preview_engine_forced_unknown(sample_pptx, tmp_path):
    with pytest.raises(RuntimeError):
        render_preview(str(sample_pptx), out_dir=str(tmp_path / "x"), engine="bogus")


def test_render_preview_empty_pngs(sample_pptx, tmp_path):
    with patch("ppt_pro_max.render_preview.detect_engine", return_value="powerpoint"), \
         patch("ppt_pro_max.render_preview._render_powerpoint", return_value=[]), \
         pytest.raises(RuntimeError, match="no slide images"):
        render_preview(str(sample_pptx), out_dir=str(tmp_path / "y"))


def test_render_preview_relative_out_dir_resolved_to_pptx_dir(sample_pptx):
    with patch("ppt_pro_max.render_preview.detect_engine", return_value="powerpoint"), \
         patch("ppt_pro_max.render_preview._render_powerpoint",
               side_effect=lambda p, d, w, h: _make_pngs(d, 2)):
        result = render_preview(str(sample_pptx), out_dir="preview_rel")
    expected = sample_pptx.parent / "preview_rel"
    assert result["pngs"][0].parent == expected
    assert result["pngs"][0].exists()


def test_render_preview_relative_pptx_resolved(sample_pptx, monkeypatch):
    # chdir to the pptx's dir so a bare filename works
    monkeypatch.chdir(sample_pptx.parent)
    with patch("ppt_pro_max.render_preview.detect_engine", return_value="powerpoint"), \
         patch("ppt_pro_max.render_preview._render_powerpoint",
               side_effect=lambda p, d, w, h: _make_pngs(d, 1)):
        result = render_preview(sample_pptx.name)
    assert result["pngs"][0].exists()


# ── _build_html ────────────────────────────────────────────────────────────
def test_build_html_embeds_every_png(tmp_path):
    d = tmp_path / "preview"
    d.mkdir()
    pngs = [d / "slide1.png", d / "slide2.png"]
    for p in pngs:
        p.write_bytes(b"PNG")
    html = _build_html(pngs, d.parent / "deck.pptx", title="My Deck")
    assert html.exists()
    text = html.read_text(encoding="utf-8")
    assert "slide1.png" in text
    assert "slide2.png" in text
    assert "My Deck" in text
    assert "2 slides" in text
