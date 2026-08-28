"""Regression tests for isolated Skill installation and runtime reporting."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from installer.install import SOURCE, copy_bundle, installed_package_version


class InstallerBundleTests(unittest.TestCase):
    def test_installed_package_version_is_read_only(self) -> None:
        self.assertIsNotNone(installed_package_version("pptx-designer"))
        self.assertIsNone(installed_package_version("package-that-does-not-exist"))

    def test_force_replaces_legacy_bundle_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            destination = Path(temp_dir) / "ppt-design-skill"
            legacy = destination / "src" / "ppt_pro_max" / "__init__.py"
            legacy.parent.mkdir(parents=True)
            legacy.write_text("legacy", encoding="utf-8")
            old_script = destination / "scripts" / "generate_ppt.py"
            old_script.parent.mkdir(parents=True)
            old_script.write_text("legacy", encoding="utf-8")

            self.assertTrue(copy_bundle(destination, force=True))
            self.assertFalse(legacy.exists())
            self.assertFalse(old_script.exists())
            self.assertTrue((destination / "SKILL.md").exists())
            self.assertEqual(
                (destination / "SKILL.md").read_text(encoding="utf-8"),
                (SOURCE / "SKILL.md").read_text(encoding="utf-8"),
            )

    def test_existing_bundle_is_not_touched_without_force(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            destination = Path(temp_dir) / "ppt-design-skill"
            destination.mkdir()
            marker = destination / "user-marker.txt"
            marker.write_text("keep", encoding="utf-8")

            self.assertFalse(copy_bundle(destination, force=False))
            self.assertEqual(marker.read_text(encoding="utf-8"), "keep")


if __name__ == "__main__":
    unittest.main()
