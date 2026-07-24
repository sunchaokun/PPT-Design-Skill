#!/usr/bin/env python3
"""Backward-compatible entry point — delegates to installer/."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "installer"))
from install import main

if __name__ == "__main__":
    main()
