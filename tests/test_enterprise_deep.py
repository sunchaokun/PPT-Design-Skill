"""Deep code review tests — boundary conditions and edge cases."""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

import pytest


# ============================================================
# Scanner edge cases
# ============================================================

class TestScannerEdgeCases:

    def test_nonexistent_project_dir(self, tmp_path):
        """Non-existent directory should return empty ProjectAsset."""
        from ppt_pro_max.enterprise.scanner import ProjectScanner
        scanner = ProjectScanner()
        asset = scanner.scan(str(tmp_path / "nonexistent"))
        assert asset.template_path is None
        assert asset.image_pool == []

    def test_multiple_pptx_uses_first(self, tmp_path):
        """Multiple .pptx in root should pick one (design doc says '唯一的', but need graceful handling)."""
        from ppt_pro_max.enterprise.scanner import ProjectScanner
        (tmp_path / "template.pptx").write_bytes(b"PK\x03\x04" + b"\x00" * 20)
        (tmp_path / "backup.pptx").write_bytes(b"PK\x03\x04" + b"\x00" * 20)
        scanner = ProjectScanner()
        asset = scanner.scan(str(tmp_path))
        assert asset.template_path is not None

    def test_logo_not_in_image_pool(self, tmp_path):
        """logo.png should be detected as LOGO, NOT added to image_pool."""
        from ppt_pro_max.enterprise.scanner import ProjectScanner
        (tmp_path / "logo.png").write_bytes(b"\x89PNG")
        (tmp_path / "hero.png").write_bytes(b"\x89PNG")
        scanner = ProjectScanner()
        asset = scanner.scan(str(tmp_path))
        assert asset.logo_path is not None
        logo_in_pool = any("logo" in Path(p).stem for p in asset.image_pool)
        assert not logo_in_pool, "logo should not appear in image_pool"

    def test_template_not_in_image_pool(self, tmp_path):
        """.pptx template should NOT appear in image_pool."""
        from ppt_pro_max.enterprise.scanner import ProjectScanner
        (tmp_path / "template.pptx").write_bytes(b"PK\x03\x04" + b"\x00" * 20)
        (tmp_path / "hero.png").write_bytes(b"\x89PNG")
        scanner = ProjectScanner()
        asset = scanner.scan(str(tmp_path))
        assert asset.template_path is not None
        pptx_in_pool = any(p.endswith(".pptx") for p in asset.image_pool)
        assert not pptx_in_pool, ".pptx should not appear in image_pool"

    def test_corrupted_brand_json(self, tmp_path):
        """Corrupted brand.json should not crash, just return None."""
        from ppt_pro_max.enterprise.scanner import ProjectScanner
        (tmp_path / "brand.json").write_text("{invalid json}", encoding="utf-8")
        scanner = ProjectScanner()
        asset = scanner.scan(str(tmp_path))
        assert asset.brand_raw is None

    def test_case_insensitive_logo(self, tmp_path):
        """LOGO detection should be case-insensitive: Logo.png, LOGO.png."""
        from ppt_pro_max.enterprise.scanner import ProjectScanner
        scanner = ProjectScanner()
        (tmp_path / "Logo.png").write_bytes(b"\x89PNG")
        asset = scanner.scan(str(tmp_path))
        assert asset.logo_path is not None

    def test_webp_image_in_pool(self, tmp_path):
        """.webp files should be included in image_pool."""
        from ppt_pro_max.enterprise.scanner import ProjectScanner
        (tmp_path / "photo.webp").write_bytes(b"RIFF")
        scanner = ProjectScanner()
        asset = scanner.scan(str(tmp_path))
        assert len(asset.image_pool) == 1

    def test_svg_image_in_pool(self, tmp_path):
        """.svg files should be included in image_pool."""
        from ppt_pro_max.enterprise.scanner import ProjectScanner
        (tmp_path / "diagram.svg").write_text("<svg></svg>", encoding="utf-8")
        scanner = ProjectScanner()
        asset = scanner.scan(str(tmp_path))
        assert len(asset.image_pool) == 1


# ============================================================
# TemplateAnalyzer edge cases
# ============================================================

class TestTemplateAnalyzerEdgeCases:

    def test_analyze_invalid_pptx(self, tmp_path):
        """Invalid .pptx should return fallback BrandSpec (not crash)."""
        from ppt_pro_max.enterprise.template_analyzer import TemplateAnalyzer
        bad_path = str(tmp_path / "bad.pptx")
        Path(bad_path).write_bytes(b"not a pptx")
        analyzer = TemplateAnalyzer()
        result = analyzer.analyze(bad_path)
        assert result.source == "template_fallback"

    def test_analyze_extracts_all_color_roles(self):
        """TemplateAnalyzer should extract at least foreground/background/primary/accent."""
        from ppt_pro_max.enterprise.template_analyzer import TemplateAnalyzer
        from pptx import Presentation
        path = str(Path(tempfile.gettempdir()) / "_test_all_colors.pptx")
        prs = Presentation()
        prs.save(path)
        analyzer = TemplateAnalyzer()
        spec = analyzer.analyze(path)
        assert spec.colors is not None
        assert "foreground" in spec.colors, f"Missing 'foreground' in {spec.colors}"
        assert "background" in spec.colors, f"Missing 'background' in {spec.colors}"

    def test_analyze_dark_mode_detection(self):
        """Dark mode should be detected from background color."""
        from ppt_pro_max.enterprise.template_analyzer import TemplateAnalyzer
        analyzer = TemplateAnalyzer()
        assert analyzer._detect_dark_mode({"background": "#000000"}) is True
        assert analyzer._detect_dark_mode({"background": "#FFFFFF"}) is False
        assert analyzer._detect_dark_mode({"background": "#1A1A2E"}) is True
        assert analyzer._detect_dark_mode({}) is False


# ============================================================
# BrandSpec edge cases
# ============================================================

class TestBrandSpecEdgeCases:

    def test_merge_brand_overrides_template_colors(self):
        """brand.json colors should override template colors."""
        from ppt_pro_max.enterprise.brand_spec import BrandSpec
        template = BrandSpec(
            source="template",
            colors={"primary": "#000000", "foreground": "#111111", "background": "#FFFFFF"},
        )
        merged = BrandSpec.merge(template, {"colors": {"primary": "#1A3C6E"}})
        assert merged.colors["primary"] == "#1A3C6E"
        assert merged.colors["foreground"] == "#111111"  # preserved
        assert merged.colors["background"] == "#FFFFFF"  # preserved

    def test_merge_no_template_colors(self):
        """Merge with empty template should still work."""
        from ppt_pro_max.enterprise.brand_spec import BrandSpec
        template = BrandSpec(source="template")
        merged = BrandSpec.merge(template, {"colors": {"primary": "#1A3C6E"}})
        assert merged.colors["primary"] == "#1A3C6E"

    def test_from_brand_json_minimal(self):
        """Minimal brand.json with only one field."""
        from ppt_pro_max.enterprise.brand_spec import BrandSpec
        spec = BrandSpec.from_brand_json({"colors": {"primary": "#1A3C6E"}})
        assert spec.source == "brand_json"
        assert spec.colors["primary"] == "#1A3C6E"
        assert spec.fonts is None
        assert spec.logo is None


# ============================================================
# parse_pages edge cases
# ============================================================

class TestParsePagesEdgeCases:

    def test_empty_string(self):
        from ppt_pro_max.enterprise.page_revision import parse_pages
        ops = parse_pages("")
        assert ops == []

    def test_whitespace_only(self):
        from ppt_pro_max.enterprise.page_revision import parse_pages
        ops = parse_pages("  ,  ,  ")
        assert ops == []

    def test_target_zero_rejected(self):
        from ppt_pro_max.enterprise.page_revision import parse_pages
        with pytest.raises(ValueError):
            parse_pages("3>0", num_slides=10)

    def test_insert_at_page_0_rejected(self):
        from ppt_pro_max.enterprise.page_revision import parse_pages
        with pytest.raises(ValueError):
            parse_pages("+0", num_slides=10)

    def test_delete_page_0_rejected(self):
        from ppt_pro_max.enterprise.page_revision import parse_pages
        with pytest.raises(ValueError):
            parse_pages("-0", num_slides=10)


# ============================================================
# compute_target_sequence edge cases
# ============================================================

class TestComputeTargetSequenceEdgeCases:

    def test_move_same_page_noop(self):
        """Moving page to after itself should be a no-op."""
        from ppt_pro_max.enterprise.page_revision import PageOp, compute_target_sequence
        ops = [PageOp(page=3, action="move", target=3)]
        order, new = compute_target_sequence(5, ops)
        assert order == [0, 1, 2, 3, 4]

    def test_swap_same_page_noop(self):
        """Swapping page with itself should be a no-op."""
        from ppt_pro_max.enterprise.page_revision import PageOp, compute_target_sequence
        ops = [PageOp(page=3, action="swap", target=3)]
        order, new = compute_target_sequence(5, ops)
        assert order == [0, 1, 2, 3, 4]

    def test_delete_already_deleted_page(self):
        """Deleting a page that's already been removed should be safe."""
        from ppt_pro_max.enterprise.page_revision import PageOp, compute_target_sequence
        ops = [PageOp(page=3, action="delete"), PageOp(page=3, action="delete")]
        order, new = compute_target_sequence(5, ops)
        assert order == [0, 1, 3, 4]

    def test_single_page_delete(self):
        """Deleting from a single-page document."""
        from ppt_pro_max.enterprise.page_revision import PageOp, compute_target_sequence
        ops = [PageOp(page=1, action="delete")]
        order, new = compute_target_sequence(1, ops)
        assert order == []

    def test_insert_at_deleted_page_goes_to_end(self):
        """Insert after a page that was deleted should append to end."""
        from ppt_pro_max.enterprise.page_revision import PageOp, compute_target_sequence
        ops = [
            PageOp(page=3, action="delete"),
            PageOp(page=3, action="insert"),
        ]
        order, new = compute_target_sequence(5, ops)
        assert order == [0, 1, 3, 4]
        assert len(new) > 0

    def test_complex_scenario_from_design_doc(self):
        """Verify: delete 3,7 + modify 2,5 + insert after 4.
        From §9.6.8 Scenario 1."""
        from ppt_pro_max.enterprise.page_revision import PageOp, compute_target_sequence
        ops = [
            PageOp(page=2, action="modify"),
            PageOp(page=5, action="modify"),
            PageOp(page=3, action="delete"),
            PageOp(page=7, action="delete"),
            PageOp(page=4, action="insert"),
        ]
        order, new = compute_target_sequence(10, ops)
        assert order == [0, 1, 3, 4, 5, 7, 8, 9]

    def test_move_from_design_doc_scenario2(self):
        """Move page 3 to after page 5. From §9.6.8 Scenario 2."""
        from ppt_pro_max.enterprise.page_revision import PageOp, compute_target_sequence
        ops = [PageOp(page=3, action="move", target=5)]
        order, new = compute_target_sequence(8, ops)
        assert order == [0, 1, 3, 4, 2, 5, 6, 7]

    def test_swap_from_design_doc_scenario3(self):
        """Swap pages 3 and 7. From §9.6.8 Scenario 3."""
        from ppt_pro_max.enterprise.page_revision import PageOp, compute_target_sequence
        ops = [PageOp(page=3, action="swap", target=7)]
        order, new = compute_target_sequence(8, ops)
        assert order == [0, 1, 6, 3, 4, 5, 2, 7]

    def test_multiple_inserts_same_position(self):
        """Two inserts after same page should both be recorded, not overwrite."""
        from ppt_pro_max.enterprise.page_revision import PageOp, compute_target_sequence
        ops = [
            PageOp(page=2, action="insert"),
            PageOp(page=2, action="insert"),
        ]
        order, new = compute_target_sequence(5, ops)
        assert order == [0, 1, 2, 3, 4]
        assert 2 in new
        assert len(new[2]) == 2


# ============================================================
# Integration: generate_ppt dispatch
# ============================================================

class TestGeneratePptIntegration:

    def test_enterprise_with_template(self, tmp_path):
        """Enterprise pipeline with template.pptx should extract brand spec."""
        from ppt_pro_max import generate_ppt
        from pptx import Presentation

        template_path = tmp_path / "template.pptx"
        prs = Presentation()
        prs.save(str(template_path))

        (tmp_path / "brand.json").write_text(
            json.dumps({"colors": {"primary": "#1A3C6E"}, "fonts": {"heading": "Arial"}}),
            encoding="utf-8",
        )

        result = generate_ppt("test query", project=str(tmp_path), dry_run=True)
        assert result["pipeline"] == "enterprise"
        assert result["brand_spec"]["source"] == "merged"
        assert result["brand_spec"]["colors"]["primary"] == "#1A3C6E"

    def test_enterprise_empty_project(self, tmp_path):
        """Enterprise pipeline with empty project should still work."""
        from ppt_pro_max import generate_ppt
        project_dir = tmp_path / "empty_project"
        project_dir.mkdir()
        result = generate_ppt("test query", project=str(project_dir), dry_run=True)
        assert result["pipeline"] == "enterprise"
        assert result["brand_spec"]["source"] == "none"

    def test_freestyle_untouched(self):
        """FreeStyle pipeline should produce identical results to before."""
        from ppt_pro_max import generate_ppt
        result = generate_ppt("AI startup pitch", dry_run=True)
        assert result.get("dry_run") is True
        assert "strategy" in result
        assert result.get("pipeline") is None
