"""Tests for StoryPlanner new strategies and business mode mapping."""

from __future__ import annotations

import pytest


class TestStoryPlannerStrategies:

    def test_education_strategy(self):
        from ppt_pro_max.planner.story_planner import StoryPlanner
        planner = StoryPlanner()
        plan = planner.plan("Python education course", strategy_override="Education Course")
        assert plan.strategy == "Education Course"
        assert plan.total_slides == 12
        assert plan.pages[0].goal == "hook"
        assert plan.pages[2].goal == "concept"

    def test_training_strategy(self):
        from ppt_pro_max.planner.story_planner import StoryPlanner
        planner = StoryPlanner()
        plan = planner.plan("Safety training workshop", strategy_override="Training Workshop")
        assert plan.strategy == "Training Workshop"
        assert plan.total_slides == 10
        assert plan.pages[3].goal == "demo"

    def test_report_strategy(self):
        from ppt_pro_max.planner.story_planner import StoryPlanner
        planner = StoryPlanner()
        plan = planner.plan("Q4 business report", strategy_override="Business Report")
        assert plan.strategy == "Business Report"
        assert plan.total_slides == 8
        assert plan.pages[2].goal == "metrics"

    def test_auto_detect_education(self):
        from ppt_pro_max.planner.story_planner import StoryPlanner
        planner = StoryPlanner()
        plan = planner.plan("Machine learning education course")
        assert plan.strategy == "Education Course"

    def test_auto_detect_training(self):
        from ppt_pro_max.planner.story_planner import StoryPlanner
        planner = StoryPlanner()
        plan = planner.plan("Employee training workshop")
        assert plan.strategy == "Training Workshop"

    def test_auto_detect_report(self):
        from ppt_pro_max.planner.story_planner import StoryPlanner
        planner = StoryPlanner()
        plan = planner.plan("Annual business report")
        assert plan.strategy == "Business Report"

    def test_slide_count_override(self):
        from ppt_pro_max.planner.story_planner import StoryPlanner
        planner = StoryPlanner()
        plan = planner.plan("Education", strategy_override="Education Course", slide_count_override=5)
        assert plan.total_slides == 5


class TestBusinessModeMapping:

    def test_education_mode_maps_to_strategy(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        assert pipeline._map_business_mode_to_strategy("education") == "Education Course"

    def test_training_mode_maps_to_strategy(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        assert pipeline._map_business_mode_to_strategy("training") == "Training Workshop"

    def test_report_mode_maps_to_strategy(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        assert pipeline._map_business_mode_to_strategy("report") == "Business Report"

    def test_pitch_mode_returns_none(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        assert pipeline._map_business_mode_to_strategy("pitch") is None

    def test_unknown_mode_returns_none(self):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        pipeline = EnterprisePipeline()
        assert pipeline._map_business_mode_to_strategy("unknown") is None

    def test_pipeline_uses_business_mode_strategy(self, tmp_path):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        import json
        project = tmp_path / "eduproj"
        project.mkdir()
        (project / "brand.json").write_text("{}", encoding="utf-8")
        pipeline = EnterprisePipeline()
        result = pipeline.run(
            query="Python course",
            project_dir=str(project),
            business_mode="education",
            density=5,
        )
        assert result["pipeline"] == "enterprise"
        assert result["num_slides"] == 7
