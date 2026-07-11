"""Tests for CLI enterprise args and EnterprisePipeline skeleton."""

from __future__ import annotations

import json
import os
from pathlib import Path

import pytest


# ============================================================
# Phase A-5: CLI enterprise arguments
# ============================================================

class TestCLIEnterpriseArgs:

    def test_cli_accepts_project_arg(self):
        from ppt_pro_max.cli import main
        import sys
        old_argv = sys.argv
        try:
            sys.argv = ["ppt-design", "test", "--project", ".", "--dry-run"]
            main()
        except SystemExit as e:
            assert e.code == 0
        finally:
            sys.argv = old_argv

    def test_cli_accepts_business_mode(self):
        from ppt_pro_max.cli import main
        import sys
        old_argv = sys.argv
        try:
            sys.argv = ["ppt-design", "test", "--project", ".", "--business-mode", "pitch", "--dry-run"]
            main()
        except SystemExit as e:
            assert e.code == 0
        finally:
            sys.argv = old_argv

    def test_cli_accepts_review_flag(self):
        from ppt_pro_max.cli import main
        import sys
        old_argv = sys.argv
        try:
            sys.argv = ["ppt-design", "test", "--project", ".", "--review", "--dry-run"]
            main()
        except SystemExit as e:
            assert e.code == 0
        finally:
            sys.argv = old_argv

    def test_cli_accepts_pages_arg(self):
        from ppt_pro_max.cli import main
        import sys
        old_argv = sys.argv
        try:
            sys.argv = ["ppt-design", "test", "--project", ".", "--from-version", "1", "--pages", "3,5", "--dry-run"]
            main()
        except SystemExit as e:
            assert e.code == 0
        finally:
            sys.argv = old_argv

    def test_cli_history_mode_query_optional(self):
        """--history should make query optional."""
        from ppt_pro_max.cli import main
        import sys
        old_argv = sys.argv
        try:
            sys.argv = ["ppt-design", "--project", ".", "--history"]
            main()
        except SystemExit as e:
            pass
        finally:
            sys.argv = old_argv

    def test_cli_output_version_and_from_version_mutually_exclusive(self):
        """--output-version and --from-version should be mutually exclusive."""
        from ppt_pro_max.cli import main
        import sys
        old_argv = sys.argv
        try:
            sys.argv = ["ppt-design", "test", "--project", ".", "--output-version", "2", "--from-version", "1"]
            main()
            assert False, "Should have exited with error"
        except SystemExit as e:
            assert e.code != 0
        finally:
            sys.argv = old_argv


# ============================================================
# Phase A-6: EnterprisePipeline skeleton
# ============================================================

class TestEnterprisePipeline:

    def test_pipeline_run_dry_run(self, tmp_path):
        """EnterprisePipeline.run() in dry_run mode should return plan."""
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="test query",
            project_dir=str(tmp_path),
            dry_run=True,
        )
        assert result.get("dry_run") is True
        assert "brand_spec" in result

    def test_pipeline_with_brand_json(self, tmp_path):
        """Pipeline should read brand.json and merge with template."""
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        (tmp_path / "brand.json").write_text(
            json.dumps({"colors": {"primary": "#1A3C6E"}, "fonts": {"heading": "Arial"}}),
            encoding="utf-8",
        )
        pipeline = EnterprisePipeline()
        result = pipeline.run(query="test", project_dir=str(tmp_path), dry_run=True)
        assert result["brand_spec"]["colors"]["primary"] == "#1A3C6E"

    def test_pipeline_with_content_json(self, tmp_path):
        """Pipeline should read content.json and include in result."""
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        (tmp_path / "content.json").write_text(
            json.dumps({"meta": {"title": "Test"}, "slides": [{"goal": "hook", "title": "Hello"}]}),
            encoding="utf-8",
        )
        pipeline = EnterprisePipeline()
        result = pipeline.run(query="test", project_dir=str(tmp_path), dry_run=True)
        assert result["assets"]["content_json"] is True

    def test_pipeline_history_mode(self, tmp_path):
        """Pipeline in history mode should return version list."""
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        v1_dir = output_dir / "v1"
        v1_dir.mkdir()
        (v1_dir / "presentation.pptx").write_bytes(b"PK\x03\x04" + b"\x00" * 20)
        (v1_dir / "meta.json").write_text(
            json.dumps({"version": 1, "query": "test", "created_at": "2026-07-11T10:00:00"}),
            encoding="utf-8",
        )
        pipeline = EnterprisePipeline()
        result = pipeline.run(query="", project_dir=str(tmp_path), history=True)
        assert result.get("history") is True
        assert len(result.get("versions", [])) >= 1


# ============================================================
# Version management
# ============================================================

class TestVersionManager:

    def test_next_version_empty(self, tmp_path):
        from ppt_pro_max.enterprise.version_manager import next_version
        assert next_version(str(tmp_path)) == 1

    def test_next_version_with_existing(self, tmp_path):
        from ppt_pro_max.enterprise.version_manager import next_version
        (tmp_path / "v1").mkdir()
        (tmp_path / "v2").mkdir()
        assert next_version(str(tmp_path)) == 3

    def test_next_version_with_gaps(self, tmp_path):
        from ppt_pro_max.enterprise.version_manager import next_version
        (tmp_path / "v1").mkdir()
        (tmp_path / "v3").mkdir()
        assert next_version(str(tmp_path)) == 4

    def test_next_version_ignores_latest(self, tmp_path):
        from ppt_pro_max.enterprise.version_manager import next_version
        (tmp_path / "v1").mkdir()
        (tmp_path / "v2").mkdir()
        (tmp_path / "latest").mkdir()
        assert next_version(str(tmp_path)) == 3

    def test_write_meta(self, tmp_path):
        from ppt_pro_max.enterprise.version_manager import write_meta
        meta = {"version": 1, "query": "test"}
        write_meta(str(tmp_path), meta)
        loaded = json.loads((tmp_path / "meta.json").read_text(encoding="utf-8"))
        assert loaded["version"] == 1

    def test_read_meta(self, tmp_path):
        from ppt_pro_max.enterprise.version_manager import read_meta
        meta = {"version": 2, "query": "hello"}
        (tmp_path / "meta.json").write_text(json.dumps(meta), encoding="utf-8")
        loaded = read_meta(str(tmp_path))
        assert loaded["version"] == 2

    def test_read_meta_missing(self, tmp_path):
        from ppt_pro_max.enterprise.version_manager import read_meta
        loaded = read_meta(str(tmp_path / "nonexistent"))
        assert loaded is None
