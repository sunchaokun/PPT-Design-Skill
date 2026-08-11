"""Tests for installer render-dependency auto-install (LibreOffice + poppler)."""
from __future__ import annotations

from unittest.mock import patch

import pytest

from installer.install import (
    _find_pdftoppm_exe,
    _find_soffice_bin,
    _winget_install,
    detect_render_deps,
    ensure_render_deps,
)


# ── detection ─────────────────────────────────────────────────────────────
def test_detect_render_deps_both_present(tmp_path):
    soffice = tmp_path / "Program Files" / "LibreOffice" / "program" / "soffice.bin"
    soffice.parent.mkdir(parents=True)
    soffice.write_bytes(b"bin")

    poppler = tmp_path / "codex" / "poppler" / "Library" / "bin" / "pdftoppm.exe"
    poppler.parent.mkdir(parents=True)
    poppler.write_bytes(b"exe")

    with patch("installer.install._find_soffice_bin", return_value=str(soffice)), \
         patch("installer.install._find_pdftoppm_exe", return_value=str(poppler)):
        deps = detect_render_deps()
    assert deps == {"libreoffice": True, "poppler": True}


def test_detect_render_deps_none_present():
    with patch("installer.install._find_soffice_bin", return_value=None), \
         patch("installer.install._find_pdftoppm_exe", return_value=None):
        deps = detect_render_deps()
    assert deps == {"libreoffice": False, "poppler": False}


def test_find_soffice_bin_prefers_bin():
    with patch("installer.install.os.path.isfile",
               side_effect=lambda p: p.endswith("soffice.bin")), \
         patch("installer.install.shutil.which", return_value=None):
        found = _find_soffice_bin()
    assert found is not None and found.endswith("soffice.bin")


def test_find_pdftoppm_exe_skips_cmd_shim():
    """Only a .CMD shim on PATH must be rejected (no real exe)."""
    with patch("installer.install.os.path.expanduser", return_value="C:/nohome"), \
         patch("installer.install.os.path.isdir", return_value=False), \
         patch("installer.install.shutil.which", return_value="C:/bin/pdftoppm.cmd"):
        assert _find_pdftoppm_exe() is None


def test_find_pdftoppm_exe_accepts_real_exe():
    with patch("installer.install.os.path.expanduser", return_value="C:/nohome"), \
         patch("installer.install.os.path.isdir", return_value=False), \
         patch("installer.install.shutil.which", return_value="C:/bin/pdftoppm.exe"):
        assert _find_pdftoppm_exe() == "C:/bin/pdftoppm.exe"


# ── winget install ────────────────────────────────────────────────────────
def test_winget_install_success():
    with patch("installer.install.shutil.which", return_value="winget"), \
         patch("installer.install.subprocess.run",
               return_value=__import__("subprocess").CompletedProcess([], 0)):
        assert _winget_install("some.pkg", "Test") is True


def test_winget_install_failure_returns_false():
    import subprocess

    with patch("installer.install.shutil.which", return_value="winget"), \
         patch("installer.install.subprocess.run",
               return_value=subprocess.CompletedProcess([], 3, stderr="boom")):
        assert _winget_install("some.pkg", "Test") is False


def test_winget_install_no_winget():
    with patch("installer.install.shutil.which", return_value=None):
        assert _winget_install("some.pkg", "Test") is False


# ── ensure_render_deps ────────────────────────────────────────────────────
def test_ensure_render_deps_skips_when_present():
    with patch("installer.install.detect_render_deps",
               return_value={"libreoffice": True, "poppler": True}), \
         patch("installer.install._winget_install") as mock_install:
        status = ensure_render_deps()
    assert status == {"libreoffice": "ok", "poppler": "ok"}
    mock_install.assert_not_called()


def test_ensure_render_deps_installs_missing():
    import subprocess

    with patch("installer.install.detect_render_deps",
               return_value={"libreoffice": False, "poppler": False}), \
         patch("installer.install._winget_install",
               side_effect=lambda pid, label: True), \
         patch("installer.install._find_soffice_bin", return_value="C:/soffice.bin"), \
         patch("installer.install._find_pdftoppm_exe", return_value="C:/pdftoppm.exe"):
        status = ensure_render_deps()
    assert status == {"libreoffice": "installed", "poppler": "installed"}
