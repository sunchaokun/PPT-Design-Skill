"""Phase 3: Content Generator — context-aware real content with copy formulas."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Any

from ppt_pro_max.decider.design_decider import PageDesign


@dataclass
class PageContent:
    position: int
    goal: str
    title: str = ""
    subtitle: str | None = None
    bullets: list[str] = field(default_factory=list)
    metrics: list[dict[str, str]] | None = None
    quote: dict[str, str] | None = None
    chart_data: dict[str, Any] | None = None
    image_keywords: str = ""


_GOAL_IMAGE_KEYWORDS: dict[str, str] = {
    "hook": "abstract technology innovation",
    "problem": "frustration stress problem dark",
    "agitation": "urgency alarm warning",
    "solution": "hope light solution bright",
    "features": "product features grid clean",
    "traction": "growth chart upward success",
    "market": "market globe world opportunity",
    "team": "team collaboration people",
    "cta": "action button start begin",
    "product": "product interface app modern",
    "business": "business strategy chart professional",
    "financial": "financial growth revenue chart",
    "demo": "product demo screenshot interface",
    "testimonials": "people happy success team",
    "pricing": "pricing plans comparison value",
    "offer": "special offer deal value",
    "proof": "data evidence proof results",
}


class ContentGenerator:
    def __init__(self, query: str = "", product_type: str = ""):
        self._query = query
        self._product_type = product_type
        self._context = self._build_context(query, product_type)

    def _build_context(self, query: str, product_type: str) -> dict[str, Any]:
        q = query.lower()
        ctx: dict[str, Any] = {
            "is_tech": any(k in q for k in ["tech", "software", "saas", "app", "platform", "sdk", "api", "cloud", "data", "neural", "算法", "技术", "软件", "平台"]),
            "is_finance": any(k in q for k in ["investor", "pitch", "seed", "funding", "路演", "融资", "投资", "估值", "arr", "mrr"]),
            "is_product": any(k in q for k in ["product", "demo", "展示", "产品", "launch", "发布"]),
            "is_sales": any(k in q for k in ["sales", "报价", "销售", "advisory", "premium", "高端"]),
            "is_sustainability": any(k in q for k in ["sustainability", "green", "carbon", "esg", "可持续", "环保", "零碳"]),
            "is_deep_tech": any(k in q for k in ["deep tech", "neural", " ml ", "gpu", "深度", "神经", " ai ", "ai ", " ai,"]),
        }
        if ctx["is_sustainability"]:
            ctx.update({"industry": "Sustainability / CleanTech", "product_name": "Carbon Tracking Platform", "pain1": "Carbon reporting is manual, error-prone, and months late", "pain2": "No real-time visibility into supply chain emissions", "pain3": "Regulatory compliance demands are accelerating globally", "solution_title": "Automated Carbon Intelligence", "solution_sub": "Real-time carbon tracking across your entire value chain", "feat1": "Scope 1-3 auto-calculation", "feat2": "Supply chain mapping", "feat3": "Regulatory report generator", "m1_label": "Carbon Reduced", "m1_value": "45K tons", "m2_label": "Companies", "m2_value": "2,800+", "m3_label": "Countries", "m3_value": "32", "m4_label": "Cost Savings", "m4_value": "$12M", "hook_title": "Every ton of carbon you can't measure, you can't reduce", "hook_sub": "GreenPath gives you real-time visibility and automated reporting across your entire value chain.", "cta_title": "Start tracking for free", "cta_sub": "Free for companies under 100 employees. No setup fee.", "chart_labels": ["2022", "2023", "2024", "2025"], "chart_values": [5, 18, 42, 78]})
        elif ctx["is_deep_tech"]:
            ctx.update({"industry": "Deep Tech / AI", "product_name": "AI Platform", "pain1": "Training takes weeks for production models", "pain2": "GPU cloud costs are exploding 3x year-over-year", "pain3": "Lab models fail silently in production environments", "solution_title": "One-Click Model Deployment", "solution_sub": "From notebook to production in 5 minutes, not 5 weeks", "feat1": "Auto-optimized inference", "feat2": "Built-in model monitoring", "feat3": "Zero-config scaling", "m1_label": "Models Deployed", "m1_value": "500+", "m2_label": "P99 Latency", "m2_value": "<10ms", "m3_label": "Accuracy", "m3_value": "99.7%", "m4_label": "Developers", "m4_value": "12K+", "hook_title": "What if deploying AI was as easy as git push?", "hook_sub": "NeuralForge makes production AI deployment instant, reliable, and cost-effective.", "cta_title": "Start deploying in minutes", "cta_sub": "Free tier available. No credit card required.", "chart_labels": ["Baseline", "v1.0", "v2.0", "v3.0"], "chart_values": [45, 72, 88, 97]})
        elif ctx["is_finance"] and ctx["is_tech"]:
            ctx.update({"industry": "AI / Technology", "product_name": "AI Platform", "pain1": "Manual data analysis takes 80% of analyst time", "pain2": "Legacy systems can't process real-time signals", "pain3": "Talent shortage: 3.5M unfilled AI roles by 2025", "solution_title": "AI-Powered Decision Engine", "solution_sub": "Transform raw data into actionable insights in seconds, not days", "feat1": "Real-time signal processing", "feat2": "Explainable AI models", "feat3": "One-click compliance reports", "m1_label": "Data Processed", "m1_value": "50TB/day", "m2_label": "Latency", "m2_value": "<100ms", "m3_label": "Accuracy", "m3_value": "94.2%", "m4_label": "Enterprise Clients", "m4_value": "200+", "hook_title": "The future of finance is AI-native", "hook_sub": "We built the infrastructure that makes institutional AI not just possible, but profitable.", "cta_title": "Schedule a demo", "cta_sub": "See the platform in action with your own data.", "chart_labels": ["Q1", "Q2", "Q3", "Q4"], "chart_values": [12, 34, 78, 156]})
        elif ctx["is_sales"]:
            ctx.update({"industry": "Financial Services", "product_name": "Premium Advisory Service", "pain1": "Portfolio management is fragmented across 5+ platforms", "pain2": "Risk exposure is invisible until it's too late", "pain3": "Client reporting takes 20+ hours per quarter", "solution_title": "Unified Wealth Intelligence", "solution_sub": "One dashboard for every asset, every risk, every client", "feat1": "Multi-custodian aggregation", "feat2": "Real-time risk analytics", "feat3": "Automated client reporting", "m1_label": "AUM", "m1_value": "$2.4B", "m2_label": "Satisfaction", "m2_value": "99%", "m3_label": "YoY Growth", "m3_value": "28%", "m4_label": "Clients", "m4_value": "400+", "hook_title": "Your clients deserve better than spreadsheets", "hook_sub": "Aurum Partners delivers institutional-grade wealth management with personal attention.", "cta_title": "Book a consultation", "cta_sub": "Complimentary portfolio review for qualified investors.", "chart_labels": ["2021", "2022", "2023", "2024"], "chart_values": [1.2, 1.6, 2.1, 2.4]})
        elif ctx["is_product"]:
            ctx.update({"industry": "SaaS / Cloud", "product_name": "Cloud Platform", "pain1": "Teams waste 30% of time on repetitive manual workflows", "pain2": "Average team uses 12+ disconnected tools daily", "pain3": "Managers have zero visibility into actual progress", "solution_title": "Intelligent Workflow Automation", "solution_sub": "Connect, automate, and optimize your entire operation", "feat1": "Visual workflow builder", "feat2": "500+ native integrations", "feat3": "AI-powered optimization", "m1_label": "Active Users", "m1_value": "50K+", "m2_label": "Retention", "m2_value": "97%", "m3_label": "Growth", "m3_value": "4.2x", "m4_label": "ARR", "m4_value": "$8M", "hook_title": "Your best people are doing your worst work", "hook_sub": "CloudFlow automates the repetitive so your team can focus on what matters.", "cta_title": "Try CloudFlow free", "cta_sub": "14-day free trial. No credit card required.", "chart_labels": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"], "chart_values": [80, 120, 190, 310, 480, 720]})
        else:
            ctx.update({"industry": "Technology", "product_name": "Innovation Platform", "pain1": "Existing solutions are fragmented and slow to adapt", "pain2": "Teams lack the tools to move at market speed", "pain3": "Data-driven decisions are impossible without real-time insights", "solution_title": "Next-Generation Platform", "solution_sub": "One intelligent platform that adapts to how you work", "feat1": "Real-time analytics", "feat2": "Seamless integrations", "feat3": "AI-powered automation", "m1_label": "Users", "m1_value": "10K+", "m2_label": "Retention", "m2_value": "95%", "m3_label": "Growth", "m3_value": "3x", "m4_label": "ARR", "m4_value": "$2M", "hook_title": "Build faster. Ship smarter. Scale effortlessly.", "hook_sub": "The platform that turns your team's potential into performance.", "cta_title": "Get started today", "cta_sub": "Free tier available. Scale as you grow.", "chart_labels": ["Q1", "Q2", "Q3", "Q4"], "chart_values": [10, 25, 45, 80]})
        return ctx

    def generate(self, page_designs: list[PageDesign], content_file: str | None = None) -> list[PageContent]:
        user_content = self._load_user_content(content_file)
        merged = {**self._context, **user_content}
        if user_content.get("product") and not user_content.get("hook_title"):
            merged["hook_title"] = f"Introducing {user_content['product']}"
        if user_content.get("product") and not user_content.get("cta_title"):
            merged["cta_title"] = f"Try {user_content['product']} free"
        if user_content.get("company") and not user_content.get("hook_sub"):
            merged["hook_sub"] = f"{user_content['company']} — Innovation that delivers"

        contents = []
        for design in page_designs:
            content = self._generate_page(design, merged)
            contents.append(content)
        return contents

    def _generate_page(self, design: PageDesign, ctx: dict[str, Any]) -> PageContent:
        goal = design.goal
        gen = _GOAL_CONTENT_GENERATORS.get(goal, _generate_generic)
        result = gen(goal, design, ctx)

        image_keywords = _GOAL_IMAGE_KEYWORDS.get(goal, "abstract")
        if ctx.get("image_keywords"):
            image_keywords = ctx["image_keywords"]

        return PageContent(
            position=design.position,
            goal=goal,
            title=result.get("title") or "",
            subtitle=result.get("subtitle") or "",
            bullets=result.get("bullets") or [],
            metrics=result.get("metrics") or None,
            quote=result.get("quote") or None,
            chart_data=result.get("chart_data") or None,
            image_keywords=image_keywords,
        )

    def _fill_template(self, template: str, goal: str, user_content: dict[str, Any]) -> str:
        if not template:
            return f"[{goal.upper()}]"
        for key, value in user_content.items():
            if isinstance(value, str):
                template = template.replace(f"{{{key}}}", value)
            elif isinstance(value, (int, float)):
                template = template.replace(f"{{{key}}}", str(value))
        dot_pattern = re.compile(r"\{(\w+(?:\.\w+(?:\[\d+\])?)*)\}")
        for match in dot_pattern.finditer(template):
            path = match.group(1)
            resolved = self._resolve_dot_path(path, user_content)
            if resolved is not None:
                template = template.replace(match.group(0), str(resolved), 1)
        return template

    def _resolve_dot_path(self, path: str, data: dict[str, Any]) -> Any:
        parts = re.split(r"\.|\[|\]", path)
        parts = [p for p in parts if p]
        current = data
        for part in parts:
            if isinstance(current, dict):
                current = current.get(part)
            elif isinstance(current, list):
                try:
                    current = current[int(part)]
                except (ValueError, IndexError):
                    return None
            else:
                return None
            if current is None:
                return None
        return current

    def _load_user_content(self, content_file: str | None) -> dict[str, Any]:
        if content_file is None:
            return {}
        try:
            with open(content_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {}


def _generate_hook(goal: str, design: PageDesign, ctx: dict) -> dict:
    company = ctx.get("company", "")
    product = ctx.get("product", ctx.get("product_name", ""))
    hook_title = ctx.get("hook_title", "")
    if hook_title:
        title = hook_title
    elif product:
        title = f"Introducing {product}"
    else:
        title = "Build faster. Ship smarter. Scale effortlessly."
    hook_sub = ctx.get("hook_sub", "")
    if not hook_sub and company:
        hook_sub = f"{company} — Innovation that delivers"
    elif not hook_sub and product:
        hook_sub = f"The platform that turns your team's potential into performance."
    return {"title": title, "subtitle": hook_sub}


def _generate_problem(goal: str, design: PageDesign, ctx: dict) -> dict:
    bullets = []
    if "pain_points" in ctx:
        for p in ctx["pain_points"][:3]:
            if isinstance(p, dict):
                line = p.get("title", "")
                desc = p.get("desc", "")
                if desc:
                    line = f"{line}: {desc}"
                bullets.append(line)
            elif isinstance(p, str):
                bullets.append(p)
    if not bullets:
        for i in range(1, 4):
            pain = ctx.get(f"pain{i}", "")
            if pain:
                bullets.append(pain)
    if not bullets:
        bullets = ["Manual processes consume 80% of productive time", "Existing tools are fragmented and don't communicate", "Teams lack visibility into real-time performance"]
    title = ctx.get("problem_title", "The problem is bigger than you think")
    return {"title": title, "bullets": bullets}


def _generate_agitation(goal: str, design: PageDesign, ctx: dict) -> dict:
    industry = ctx.get("industry", "technology")
    return {
        "title": "This is costing you more than you realize",
        "subtitle": f"Companies in {industry} lose an average of $2.4M annually to operational inefficiency",
        "bullets": [
            "Each day of delay costs $67K in lost revenue opportunity",
            "Competitors who solve this are growing 3x faster",
            "The gap widens every quarter you wait",
        ],
    }


def _generate_solution(goal: str, design: PageDesign, ctx: dict) -> dict:
    title = ctx.get("solution_title", "A fundamentally different approach")
    sub = ctx.get("solution_sub", "")
    bullets = []
    if "pain_points" in ctx:
        for p in ctx["pain_points"][:3]:
            if isinstance(p, dict):
                desc = p.get("desc", p.get("title", ""))
                bullets.append(f"Eliminates: {desc}")
    if not bullets:
        for i in range(1, 4):
            feat = ctx.get(f"feat{i}", "")
            if feat:
                bullets.append(feat)
    return {"title": title, "subtitle": sub, "bullets": bullets}


def _generate_product(goal: str, design: PageDesign, ctx: dict) -> dict:
    product = ctx.get("product", ctx.get("product_name", "Our Platform"))
    return {
        "title": f"How {product} works",
        "subtitle": "Three steps to transform your workflow",
        "bullets": [
            "Connect: Integrate with your existing tools in minutes, not months",
            "Automate: Build intelligent workflows with our visual designer",
            "Optimize: AI continuously improves your processes",
        ],
    }


def _generate_features(goal: str, design: PageDesign, ctx: dict) -> dict:
    bullets = []
    if "features" in ctx:
        for f in ctx["features"][:3]:
            if isinstance(f, dict):
                name = f.get("title", f.get("name", ""))
                desc = f.get("desc", f.get("benefit", ""))
                bullets.append(f"{name}: {desc}" if desc else name)
            elif isinstance(f, str):
                bullets.append(f)
    if not bullets:
        for i in range(1, 4):
            feat = ctx.get(f"feat{i}", "")
            if feat:
                bullets.append(feat)
    if not bullets:
        bullets = ["Lightning-fast performance", "Enterprise-grade security", "Intuitive user experience"]
    return {"title": "Built for how you actually work", "bullets": bullets}


def _generate_traction(goal: str, design: PageDesign, ctx: dict) -> dict:
    metrics = []
    if "metrics" in ctx:
        for k, v in list(ctx["metrics"].items())[:4]:
            label = k.replace("_", " ").title()
            metrics.append({"label": label, "value": str(v)})
    if not metrics:
        for i in range(1, 5):
            label = ctx.get(f"m{i}_label", f"Metric {i}")
            value = ctx.get(f"m{i}_value", "N/A")
            metrics.append({"label": label, "value": value})
    chart_data = None
    if design.chart_type:
        chart_data = _build_chart(ctx, design.chart_type)
    return {"title": "The numbers speak for themselves", "metrics": metrics, "chart_data": chart_data}


def _generate_market(goal: str, design: PageDesign, ctx: dict) -> dict:
    industry = ctx.get("industry", "technology")
    return {
        "title": f"A massive, underserved {industry} market",
        "bullets": [
            "Total addressable market: $47B and growing at 23% CAGR",
            "Only 12% of enterprises have adopted modern solutions",
            "Regulatory tailwinds are accelerating adoption globally",
        ],
    }


def _generate_business(goal: str, design: PageDesign, ctx: dict) -> dict:
    return {
        "title": "A business model built for scale",
        "bullets": [
            "SaaS subscription with 85%+ gross margins",
            "Land-and-expand: free tier → team → enterprise",
            "Net revenue retention: 120%+ (expansion exceeds churn)",
        ],
    }


def _generate_team(goal: str, design: PageDesign, ctx: dict) -> dict:
    return {
        "title": "Built by operators who've been in your shoes",
        "bullets": [
            "CEO: 15yr industry experience, 2x founder (1 exit)",
            "CTO: Former principal engineer at top tech company",
            "VP Sales: Built $0→$50M ARR pipeline at previous startup",
        ],
    }


def _generate_financial(goal: str, design: PageDesign, ctx: dict) -> dict:
    chart_data = None
    if design.chart_type:
        chart_data = _build_chart(ctx, design.chart_type)
    return {
        "title": "Strong unit economics, clear path to profitability",
        "bullets": [
            "Current ARR: $8M, targeting $20M by year-end",
            "CAC payback: 8 months (industry avg: 18 months)",
            "Rule of 40: ARR growth (60%) + FCF margin (-15%) = 45%",
        ],
        "chart_data": chart_data,
    }


def _generate_demo(goal: str, design: PageDesign, ctx: dict) -> dict:
    product = ctx.get("product", ctx.get("product_name", "the platform"))
    return {
        "title": f"See {product} in action",
        "subtitle": "From setup to impact in under 5 minutes",
        "bullets": [
            "Step 1: Connect your data sources with one click",
            "Step 2: Configure intelligent workflows visually",
            "Step 3: Watch AI optimize and deliver results",
        ],
    }


def _generate_testimonials(goal: str, design: PageDesign, ctx: dict) -> dict:
    return {
        "title": "Trusted by industry leaders",
        "quote": {"text": "This platform transformed how we operate. We went from weeks of manual work to minutes.", "author": "VP of Operations, Fortune 500 Company"},
        "bullets": [
            "Reduced processing time by 85% within first month",
            "Saved $1.2M annually in operational costs",
            "Team satisfaction scores increased by 40%",
        ],
    }


def _generate_pricing(goal: str, design: PageDesign, ctx: dict) -> dict:
    return {
        "title": "Simple, transparent pricing",
        "subtitle": "Start free. Scale as you grow.",
        "bullets": [
            "Starter: Free for teams up to 5 — full features",
            "Pro: $29/user/mo — advanced automation + priority support",
            "Enterprise: Custom — SSO, audit logs, dedicated CSM",
        ],
    }


def _generate_proof(goal: str, design: PageDesign, ctx: dict) -> dict:
    return {
        "title": "Proven results across industries",
        "bullets": [
            "3 of the top 10 banks use our platform for compliance",
            "99.99% uptime SLA — backed by service credits",
            "SOC 2 Type II certified, GDPR & CCPA compliant",
        ],
    }


def _generate_offer(goal: str, design: PageDesign, ctx: dict) -> dict:
    return {
        "title": "Limited-time launch offer",
        "subtitle": "Lock in early-adopter pricing before it expires",
        "bullets": [
            "50% off annual plans for the first 12 months",
            "Free onboarding and migration assistance",
            "Dedicated success manager included",
        ],
    }


def _generate_cta(goal: str, design: PageDesign, ctx: dict) -> dict:
    title = ctx.get("cta_title", "Ready to get started?")
    sub = ctx.get("cta_sub", "Join thousands of teams already transforming their work.")
    return {"title": title, "subtitle": sub}


def _generate_generic(goal: str, design: PageDesign, ctx: dict) -> dict:
    product = ctx.get("product", ctx.get("product_name", "our platform"))
    return {"title": f"More about {product}", "bullets": [f"Key insight about {product}", f"Supporting evidence and data", f"Next steps and action items"]}


def _build_chart(ctx: dict, chart_type: str) -> dict[str, Any]:
    labels = ctx.get("chart_labels", ["Q1", "Q2", "Q3", "Q4"])
    values = ctx.get("chart_values", [10, 25, 45, 80])
    return {"type": chart_type, "data": {"labels": labels, "values": values}}


_GOAL_CONTENT_GENERATORS: dict[str, Any] = {
    "hook": _generate_hook,
    "problem": _generate_problem,
    "agitation": _generate_agitation,
    "solution": _generate_solution,
    "product": _generate_product,
    "features": _generate_features,
    "traction": _generate_traction,
    "market": _generate_market,
    "business": _generate_business,
    "team": _generate_team,
    "financial": _generate_financial,
    "demo": _generate_demo,
    "testimonials": _generate_testimonials,
    "pricing": _generate_pricing,
    "proof": _generate_proof,
    "offer": _generate_offer,
    "cta": _generate_cta,
    "content": _generate_generic,
}
