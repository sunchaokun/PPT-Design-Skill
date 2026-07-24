"""Phase 3 tests: Proactive component match + expanded infer + count rules + C9 image height."""

import os
import tempfile


def _make_component_lib(td):
    from ppt_pro_max.enterprise.component_library import ComponentLibrary
    db = os.path.join(td, "index.db")
    lib = ComponentLibrary(db_path=db, storage_dir=os.path.join(td, "storage"))
    return lib


def _close_lib(lib):
    try:
        lib._db.close()
    except Exception:
        pass


# ── C1: _proactive_component_match ──────────────────────────────────

class TestProactiveComponentMatch:
    def test_no_match_when_no_bullets_no_cards(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        r = PrecisionRenderer()
        result = r._proactive_component_match([], [], "auto", None)
        assert result == (None, None)

    def test_no_match_when_no_component_lib(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        r = PrecisionRenderer()
        result = r._proactive_component_match(["step 1", "step 2"], [], "auto", None)
        assert result == (None, None)

    def test_match_with_keywords_process(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            lib.add(type="group", category="process", variant="chevron", node_count=3,
                    xml_parts={"group": b"<grpSp/>"})
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                ["流程步骤一", "流程步骤二", "流程步骤三"], [], "auto", lib
            )
            assert ct == "group"
            assert cat == "process"
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_match_with_keywords_swot(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            lib.add(type="group", category="swot", variant="quadrant", node_count=4,
                    xml_parts={"group": b"<grpSp/>"})
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                ["strengths", "weaknesses", "opportunities", "threats"], [], "auto", lib
            )
            assert ct == "group"
            assert cat == "swot"
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_match_by_count_when_no_keywords(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            lib.add(type="group", category="matrix", variant="quadrant", node_count=4,
                    xml_parts={"group": b"<grpSp/>"})
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                ["item A", "item B", "item C", "item D"], [], "auto", lib
            )
            assert ct == "group"
            assert cat == "matrix"
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_match_by_count_2_items_comparison(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            lib.add(type="group", category="comparison", variant="two-col", node_count=2,
                    xml_parts={"group": b"<grpSp/>"})
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                ["option A", "option B"], [], "auto", lib
            )
            assert ct == "group"
            assert cat == "comparison"
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_match_by_count_3_items_pyramid(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            lib.add(type="group", category="pyramid", variant="tiered", node_count=3,
                    xml_parts={"group": b"<grpSp/>"})
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                ["base", "middle", "top"], [], "auto", lib
            )
            assert ct == "group"
            assert cat == "pyramid"
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_match_by_count_5_items_process(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            lib.add(type="group", category="process", variant="chevron", node_count=5,
                    xml_parts={"group": b"<grpSp/>"})
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                ["a", "b", "c", "d", "e"], [], "auto", lib
            )
            assert ct == "group"
            assert cat == "process"
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_match_by_count_6_items_cycle(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            lib.add(type="group", category="cycle", variant="circular", node_count=6,
                    xml_parts={"group": b"<grpSp/>"})
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                ["a", "b", "c", "d", "e", "f"], [], "auto", lib
            )
            assert ct == "group"
            assert cat == "cycle"
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_no_match_when_lib_empty(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                ["流程步骤一", "流程步骤二"], [], "auto", lib
            )
            assert ct is None
            assert cat is None
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_no_match_when_component_has_no_xml(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            lib.add(type="group", category="process", variant="chevron", node_count=3,
                    xml_parts={})
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                ["流程步骤一", "流程步骤二", "流程步骤三"], [], "auto", lib
            )
            assert ct is None
            assert cat is None
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_layout_hint_cards_infers_infographic(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            lib.add(type="group", category="infographic", variant="card-row", node_count=3,
                    xml_parts={"group": b"<grpSp/>"})
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                [], [{"title": "A"}, {"title": "B"}, {"title": "C"}], "cards", lib
            )
            assert ct == "group"
            assert cat == "infographic"
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_existing_component_type_not_overridden(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        r = PrecisionRenderer()
        result = r._proactive_component_match(["a", "b"], [], "auto", "dummy_lib", existing_component_type="group")
        assert result == (None, None)

    def test_mood_fallback_when_keyword_category_not_in_lib(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            lib.add(type="group", category="hierarchy", variant="tree", node_count=4,
                    xml_parts={"group": b"<grpSp/>"})
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                ["item A", "item B", "item C", "item D"], [], "auto", lib,
                mood="mckinsey",
            )
            assert ct == "group"
            assert cat == "hierarchy"
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_keyword_category_takes_priority_over_mood(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        td = tempfile.mkdtemp()
        try:
            lib = _make_component_lib(td)
            lib.add(type="group", category="matrix", variant="quadrant", node_count=4,
                    xml_parts={"group": b"<grpSp/>"})
            lib.add(type="group", category="hierarchy", variant="tree", node_count=4,
                    xml_parts={"group": b"<grpSp/>"})
            r = PrecisionRenderer()
            ct, cat = r._proactive_component_match(
                ["矩阵A", "象限B", "象限C", "象限D"], [], "auto", lib,
                mood="mckinsey",
            )
            assert ct == "group"
            assert cat == "matrix"
            _close_lib(lib)
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)


# ── Expanded infer_component_category ───────────────────────────────

class TestInferComponentCategoryExpanded:
    def test_cycle_keywords(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["循环迭代", "反馈闭环", "持续改进"])
        assert ct == "group"
        assert cat == "cycle"

    def test_pyramid_keywords(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["金字塔底层", "中间层", "顶层"])
        assert ct == "group"
        assert cat == "pyramid"

    def test_matrix_keywords(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["矩阵分析", "象限评估", "四象限模型"])
        assert ct == "group"
        assert cat == "matrix"

    def test_comparison_keywords(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["对比方案A", "对比方案B"])
        assert ct == "group"
        assert cat == "comparison"

    def test_funnel_keywords(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["漏斗转化", "转化率分析", "conversion funnel"])
        assert ct == "group"
        assert cat == "funnel"

    def test_radial_keywords(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["辐射网络", "核心hub", "中心节点"])
        assert ct == "group"
        assert cat == "radial"

    def test_infographic_keywords(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["信息图展示", "数据统计", "关键指标"])
        assert ct == "group"
        assert cat == "infographic"

    def test_chart_keywords(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["图表分析", "柱状对比图", "饼图占比"])
        assert ct == "group"
        assert cat == "chart"

    def test_english_cycle(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["cycle iteration", "feedback loop"])
        assert ct == "group"
        assert cat == "cycle"

    def test_english_funnel(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["funnel stages", "conversion analysis"])
        assert ct == "group"
        assert cat == "funnel"

    def test_english_radial(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["radial diagram", "hub and spoke"])
        assert ct == "group"
        assert cat == "radial"

    def test_count_fallback_2_comparison(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["item A", "item B"])
        assert ct == "group"
        assert cat == "comparison"

    def test_count_fallback_3_pyramid(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["layer 1", "layer 2", "layer 3"])
        assert ct == "group"
        assert cat == "pyramid"

    def test_count_fallback_4_matrix(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["q1", "q2", "q3", "q4"])
        assert ct == "group"
        assert cat == "matrix"

    def test_count_fallback_5_process(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["a", "b", "c", "d", "e"])
        assert ct == "group"
        assert cat == "process"

    def test_count_fallback_6_cycle(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["a", "b", "c", "d", "e", "f"])
        assert ct == "group"
        assert cat == "cycle"

    def test_count_fallback_8_cycle(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["a", "b", "c", "d", "e", "f", "g", "h"])
        assert ct == "group"
        assert cat == "cycle"

    def test_count_fallback_9_infographic(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(list("abcdefghi"))
        assert ct == "group"
        assert cat == "infographic"

    def test_too_few_items_returns_none(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["only one"])
        assert ct is None
        assert cat is None

    def test_too_many_items_returns_none(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category([f"item {i}" for i in range(10)])
        assert ct is None
        assert cat is None

    def test_original_swot_still_works(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["strengths", "weaknesses", "opportunities", "threats"])
        assert ct == "group"
        assert cat == "swot"

    def test_original_hierarchy_still_works(self):
        from ppt_pro_max.enterprise.content_parser import infer_component_category
        ct, cat = infer_component_category(["CEO", "CTO", "CFO", "VP Engineering"])
        assert ct == "group"
        assert cat == "hierarchy"


# ── C9: classify_image_size with height ─────────────────────────────

class TestClassifyImageSizeWithHeight:
    def _classify(self, w, h):
        from ppt_pro_max.enterprise.image_matcher import classify_image_size
        from PIL import Image as PILImage
        td = tempfile.mkdtemp()
        try:
            path = os.path.join(td, "test.png")
            img = PILImage.new("RGB", (w, h), "blue")
            img.save(path)
            img.close()
            result = classify_image_size(path)
            return result
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def test_narrow_tall_image_not_background(self):
        result = self._classify(1600, 200)
        assert result != "background"

    def test_wide_short_image_not_background(self):
        result = self._classify(1600, 200)
        assert result in ("scene", "icon")

    def test_large_square_is_background(self):
        result = self._classify(2000, 2000)
        assert result == "background"

    def test_medium_landscape_is_scene(self):
        result = self._classify(1200, 800)
        assert result == "scene"

    def test_small_square_is_icon(self):
        result = self._classify(400, 400)
        assert result == "icon"

    def test_narrow_tall_800px_is_scene(self):
        result = self._classify(900, 1600)
        assert result == "scene"


# ── PagePlan.mood field ─────────────────────────────────────────────

class TestPagePlanMoodField:
    def test_mood_field_exists(self):
        from ppt_pro_max.enterprise.design_dna_extractor import PagePlan
        plan = PagePlan(page_type="content", mood="mckinsey")
        assert plan.mood == "mckinsey"

    def test_mood_default_none(self):
        from ppt_pro_max.enterprise.design_dna_extractor import PagePlan
        plan = PagePlan(page_type="content")
        assert plan.mood is None

    def test_mood_category_map(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        r = PrecisionRenderer()
        cat = r._mood_to_preferred_category("mckinsey")
        assert cat == "hierarchy"

    def test_mood_category_map_creative(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        r = PrecisionRenderer()
        cat = r._mood_to_preferred_category("creative")
        assert cat == "radial"

    def test_mood_category_map_unknown(self):
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        r = PrecisionRenderer()
        cat = r._mood_to_preferred_category("unknown_mood")
        assert cat is None
