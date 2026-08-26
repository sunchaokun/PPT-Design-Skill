"""Locate optional desktop render dependencies on supported platforms."""

from __future__ import annotations

import os
import platform
import shutil
from pathlib import Path


def _windows_candidates(command: str) -> list[Path]:
    if platform.system() != "Windows":
        return []
    roots = [
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
        os.environ.get("LOCALAPPDATA", ""),
    ]
    candidates = []
    if command == "soffice":
        for root in roots:
            if root:
                candidates.append(Path(root) / "LibreOffice" / "program" / "soffice.exe")
    elif command == "pdftoppm":
        for root in roots:
            if root:
                candidates.extend(Path(root).glob("Microsoft\\WinGet\\Packages\\*Poppler*\\*\\Library\\bin\\pdftoppm.exe"))
    return candidates


def _registry_candidates(command: str) -> list[Path]:
    if platform.system() != "Windows" or command != "soffice":
        return []
    try:
        import winreg
    except ImportError:
        return []
    found: list[Path] = []
    locations = (
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\LibreOffice\UNO\InstallPath"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\LibreOffice\UNO\InstallPath"),
    )
    for hive, key_name in locations:
        try:
            with winreg.OpenKey(hive, key_name) as key:
                value, _ = winreg.QueryValueEx(key, None)
                found.append(Path(value) / "soffice.exe")
        except (FileNotFoundError, OSError):
            continue
    return found


def resolve_executable(command: str) -> str | None:
    """Return an executable path from PATH, standard Windows paths, or registry."""
    from_path = shutil.which(command)
    if from_path:
        return from_path
    for candidate in [*_windows_candidates(command), *_registry_candidates(command)]:
        if candidate.is_file():
            return str(candidate)
    return None
