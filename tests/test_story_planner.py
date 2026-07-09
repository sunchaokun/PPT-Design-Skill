"""Tests for story planner."""

from ppt_pro_max.planner.story_planner import StoryPlanner


def test_plan_default_strategy():
    planner = StoryPlanner()
    plan = planner.plan("AI产品融资路演")
    assert plan.strategy == "YC Seed Deck"
    assert plan.total_slides == 10
    assert len(plan.pages) == 10
    assert plan.pages[0].goal == "hook"
    assert plan.pages[-1].goal == "cta"


def test_plan_product_demo():
    planner = StoryPlanner()
    plan = planner.plan("产品展示demo")
    assert plan.strategy == "Product Demo"
    assert plan.total_slides == 8


def test_plan_sales_pitch():
    planner = StoryPlanner()
    plan = planner.plan("销售报价方案")
    assert plan.strategy == "Sales Pitch"
    assert plan.total_slides == 7


def test_plan_override_slide_count():
    planner = StoryPlanner()
    plan = planner.plan("融资路演", slide_count_override=5)
    assert plan.total_slides == 5
    assert len(plan.pages) == 5


def test_plan_override_strategy():
    planner = StoryPlanner()
    plan = planner.plan("随便什么", strategy_override="Product Demo")
    assert plan.strategy == "Product Demo"


def test_emotion_arc():
    planner = StoryPlanner()
    plan = planner.plan("融资路演")
    assert "→" in plan.emotion_arc
    emotions = plan.emotion_arc.split("→")
    assert len(emotions) == plan.total_slides
