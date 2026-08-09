"""Strict tests for the bundled design database (BM25 search engine).

Covers: data availability, domain detection, search result structure,
design system generation, dial resolution, robustness edge cases.
"""

import sys
from pathlib import Path
from typing import ClassVar

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from ppt_pro_max.adapters import design_search as ds


@pytest.fixture(scope="module")
def data_dir():
    return ds.DATA_DIR


class TestDataAvailability:
    """7 PPT-relevant CSVs must exist and be loadable."""

    REQUIRED_CSVS: ClassVar[list[str]] = [
        "colors.csv", "typography.csv", "styles.csv", "products.csv",
        "landing.csv", "ui-reasoning.csv", "motion.csv",
    ]

    def test_data_dir_exists(self, data_dir):
        assert data_dir.exists(), "Data dir missing"
        assert data_dir.is_dir()

    @pytest.mark.parametrize("csv_name", REQUIRED_CSVS)
    def test_csv_exists(self, data_dir, csv_name):
        fp = data_dir / csv_name
        assert fp.exists(), f"Missing {csv_name}"
        assert fp.stat().st_size > 100, f"{csv_name} too small (possibly truncated)"

    def test_all_configs_point_to_existing_files(self, data_dir):
        for domain, cfg in ds.PPT_CSV_CONFIG.items():
            assert (data_dir / cfg["file"]).exists(), (
                f"domain {domain} references missing {cfg['file']}"
            )

    def test_csv_config_has_required_keys(self):
        for domain, cfg in ds.PPT_CSV_CONFIG.items():
            assert "search_cols" in cfg, f"{domain} missing search_cols"
            assert "output_cols" in cfg, f"{domain} missing output_cols"
            assert len(cfg["search_cols"]) >= 2
            assert len(cfg["output_cols"]) >= 3


class TestBM25:
    def test_fit_empty(self):
        b = ds.BM25()
        b.fit([])
        assert b.N == 0

    def test_ranked_by_relevance(self):
        docs = [
            "fintech banking payments",
            "healthcare hospital clinical",
            "education university learning",
        ]
        b = ds.BM25()
        b.fit(docs)
        ranked = b.score("banking payments")
        # Top result should be the fintech doc (idx 0)
        assert ranked[0][0] == 0
        assert ranked[0][1] > 0

    def test_no_match_returns_zero(self):
        b = ds.BM25()
        b.fit(["fintech banking", "healthcare clinical"])
        ranked = b.score("quantum physics")
        assert all(s == 0 for _, s in ranked)

    def test_tokenize_filters_short(self):
        b = ds.BM25()
        tokens = b.tokenize("A quick test of X")
        assert all(len(t) >= 2 for t in tokens)
        assert "x" not in tokens


class TestSearch:
    def test_product_domain(self):
        results = ds.search("fintech", "product", 3)
        assert len(results) >= 1
        assert "Product Type" in results[0]

    def test_style_domain(self):
        results = ds.search("dark tech", "style", 3)
        assert len(results) >= 1
        assert "Style Category" in results[0]

    def test_color_domain(self):
        results = ds.search("fintech", "color", 2)
        assert len(results) >= 1
        assert "Primary" in results[0]

    def test_typography_domain(self):
        results = ds.search("professional", "typography", 2)
        assert len(results) >= 1
        assert "Heading Font" in results[0]

    def test_landing_domain(self):
        results = ds.search("conversion pricing", "landing", 2)
        assert len(results) >= 1
        assert "Pattern Name" in results[0]

    def test_empty_query_returns_empty(self):
        assert ds.search("", "style", 3) == []

    def test_unknown_domain_returns_empty(self):
        assert ds.search("anything", "nonexistent_domain", 3) == []

    def test_max_results_honored(self):
        results = ds.search("professional modern", "style", 2)
        assert len(results) <= 2

    def test_auto_domain_detection(self):
        results = ds.search("palette hex accent colors")
        assert len(results) >= 1

    def test_results_are_dicts_with_string_values(self):
        results = ds.search("fintech", "product", 2)
        for r in results:
            assert isinstance(r, dict)
            for v in r.values():
                assert isinstance(v, str)


class TestDetectDomain:
    @pytest.mark.parametrize("query,expected", [
        ("palette hex color rgb", "color"),
        ("font pairing serif heading", "typography"),
        ("gsap scrolltrigger parallax", "motion"),
        ("landing page cta hero", "landing"),
        ("saas fintech dashboard", "product"),
        ("glassmorphism brutalism minimalism", "style"),
        ("gibberish qqqzzz", "style"),  # fallback
    ])
    def test_domain_detection(self, query, expected):
        assert ds.detect_domain(query) == expected


class TestDesignSystemGenerator:
    def test_generate_returns_full_dict(self):
        result = ds.get_design_system("AI startup pitch")
        assert isinstance(result, dict)
        for key in ["colors", "typography", "style", "pattern",
                    "key_effects", "anti_patterns", "decision_rules", "dials"]:
            assert key in result, f"missing {key}"

    def test_colors_have_primary(self):
        result = ds.get_design_system("AI startup pitch")
        colors = result["colors"]
        assert "primary" in colors
        assert colors["primary"].startswith("#")
        assert len(colors["primary"]) == 7  # #RRGGBB

    def test_typography_has_heading_body(self):
        result = ds.get_design_system("AI startup pitch")
        typo = result["typography"]
        assert typo["heading"]
        assert typo["body"]

    def test_dials_reflect_input(self):
        result = ds.get_design_system("AI startup pitch", variance=8, motion=2, density=10)
        dials = result["dials"]
        assert dials["variance"] == 8
        assert dials["variance_label"] == "Bold / Asymmetric"
        assert dials["motion"] == 2
        assert dials["motion_label"] == "Subtle"
        assert dials["density"] == 10
        assert dials["density_label"] == "Dense / Dashboard"

    def test_dials_none_when_unset(self):
        result = ds.get_design_system("AI startup pitch")
        dials = result["dials"]
        assert dials["variance"] is None
        assert dials["density"] is None

    def test_motion_snippet_when_motion_set(self):
        result = ds.get_design_system("AI startup pitch", motion=5)
        if result["motion_snippet"]:
            snippet = result["motion_snippet"]
            assert "Intensity Tier" in snippet or "Category" in snippet

    def test_normalized_style_fields(self):
        result = ds.get_design_system("AI startup pitch")
        assert "style_name" in result
        assert "pattern_name" in result

    def test_generate_idempotent(self):
        a = ds.get_design_system("AI startup pitch", variance=5, motion=3, density=5)
        b = ds.get_design_system("AI startup pitch", variance=5, motion=3, density=5)
        assert a["colors"]["primary"] == b["colors"]["primary"]
        assert a["style_name"] == b["style_name"]

    def test_category_detected(self):
        result = ds.get_design_system("SaaS dashboard")
        assert result["category"]  # non-empty

    def test_apply_reasoning_returns_rules(self):
        r = ds.DesignSystemGenerator().apply_reasoning("dashboard")
        assert "pattern" in r
        assert "style_priority" in r
        assert isinstance(r["style_priority"], list)

    def test_apply_reasoning_unknown_category_has_defaults(self):
        r = ds.DesignSystemGenerator().apply_reasoning("nonexistent_category_xyz")
        assert r["pattern"]  # has default pattern


class TestDialResolution:
    def test_variance_low(self):
        assert ds._resolve_dial("variance", 2)["label"] == "Centered / Minimal"

    def test_variance_mid(self):
        assert ds._resolve_dial("variance", 5)["label"] == "Balanced / Modern"

    def test_variance_high(self):
        assert ds._resolve_dial("variance", 9)["label"] == "Bold / Asymmetric"

    def test_clamps_out_of_range(self):
        assert ds._resolve_dial("variance", 999)["value"] == 10
        assert ds._resolve_dial("variance", -5)["value"] == 1

    def test_none_returns_none(self):
        assert ds._resolve_dial("variance", None) is None

    @pytest.mark.parametrize("dial_name", ["variance", "motion", "density"])
    def test_all_dials_resolvable(self, dial_name):
        for v in range(1, 11):
            r = ds._resolve_dial(dial_name, v)
            assert r is not None, f"{dial_name}={v} failed"
            assert r["value"] == v
