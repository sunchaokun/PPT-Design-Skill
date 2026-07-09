"""Generate showcase examples — 6 high-quality PPTs for GitHub display."""

import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from ppt_pro_max import generate_ppt
from ppt_pro_max.renderer.image_fetcher import ImageFetcher

EXAMPLES_DIR = os.path.join(os.path.dirname(__file__), "..", "examples")
os.makedirs(EXAMPLES_DIR, exist_ok=True)


def _create_test_image(path: str, width: int = 1920, height: int = 1080, color: tuple = (37, 99, 235), text: str = ""):
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (width, height), color=color)
    draw = ImageDraw.Draw(img)
    if text and width > 80 and height > 80:
        margin = min(40, width // 10, height // 10)
        draw.rectangle([margin, margin, width - margin, height - margin], fill=(255, 255, 255), outline=(200, 200, 200), width=max(1, margin // 20))
        if width > 200 and height > 100:
            draw.text((width // 4, height // 2 - 20), text, fill=color)
    img.save(path, "JPEG", quality=90)


def _seed_image_cache(cache_dir: str, keywords: str, emotion: str, goal: str, width: int, height: int, color: tuple, text: str):
    import hashlib
    fetcher = ImageFetcher(mode="search", image_cache_dir=cache_dir)
    cache_key = f"search:{keywords}:{width}x{height}"
    cache_hash = hashlib.md5(cache_key.encode()).hexdigest()
    cached_path = os.path.join(fetcher._cache_dir, f"{cache_hash}.jpg")
    if not os.path.exists(cached_path):
        _create_test_image(cached_path, width, height, color, text)
    return cached_path


def generate_all():
    cache_dir = os.path.join(tempfile.gettempdir(), "ppt-showcase-cache")
    os.makedirs(cache_dir, exist_ok=True)

    # Color palettes for different themes
    colors = {
        "dark-tech": (15, 23, 42),
        "professional": (37, 99, 235),
        "warm-elegant": (146, 64, 14),
        "vibrant-startup": (124, 58, 237),
        "nature-calm": (5, 150, 105),
    }

    examples = []

    # === Example 1: AI Startup Investor Pitch (dark-tech, YC Seed, 10 pages) ===
    print("Generating 1/6: AI Startup Investor Pitch...")
    seed_kw = "artificial intelligence futuristic dark blue technology"
    _seed_image_cache(cache_dir, seed_kw, "curiosity", "hook", 144, 48, colors["dark-tech"], "AI Logo")
    _seed_image_cache(cache_dir, seed_kw, "curiosity", "hook", 768, 720, colors["dark-tech"], "AI Innovation")

    r1 = generate_ppt(
        query="AI Startup Investor Pitch",
        strategy="YC Seed Deck",
        theme="dark-tech",
        image_mode="search",
        image_config={"image_cache_dir": cache_dir},
        output=os.path.join(EXAMPLES_DIR, "01-ai-startup-investor-pitch.pptx"),
    )
    examples.append(("01-ai-startup-investor-pitch.pptx", "AI Startup Investor Pitch", "YC Seed Deck", "Dark Tech", r1))
    print(f"  Done: {r1['page_count']} pages, QA={r1['qa']['passed']}")

    # === Example 2: SaaS Product Demo (professional, Product Demo, 8 pages) ===
    print("Generating 2/6: SaaS Product Demo...")
    seed_kw2 = "software dashboard clean modern professional"
    _seed_image_cache(cache_dir, seed_kw2, "curiosity", "hook", 144, 48, colors["professional"], "SaaS Logo")
    _seed_image_cache(cache_dir, seed_kw2, "trust", "demo", 768, 720, colors["professional"], "Product Screenshot")

    content_file2 = os.path.join(cache_dir, "saas_content.json")
    with open(content_file2, "w", encoding="utf-8") as f:
        json.dump({
            "company": "CloudFlow",
            "product": "CloudFlow Platform",
            "tagline": "Automate your workflow, amplify your team",
            "metrics": {"users": "50K+", "retention": "97%", "growth": "4.2x", "arr": "$8M"},
            "pain_points": [
                {"title": "Manual Workflows", "desc": "Teams waste 30% of time on repetitive tasks"},
                {"title": "Tool Sprawl", "desc": "Average team uses 12+ disconnected tools"},
                {"title": "No Visibility", "desc": "Managers can't see what's actually happening"},
            ],
            "chart_data": {"mrr": {"labels": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"], "values": [80, 120, 190, 310, 480, 720]}},
        }, f, ensure_ascii=False)

    r2 = generate_ppt(
        query="SaaS Product Demo",
        strategy="Product Demo",
        theme="professional",
        content_file=content_file2,
        image_mode="search",
        image_config={"image_cache_dir": cache_dir},
        output=os.path.join(EXAMPLES_DIR, "02-saas-product-demo.pptx"),
    )
    examples.append(("02-saas-product-demo.pptx", "SaaS Product Demo", "Product Demo", "Professional Modern", r2))
    print(f"  Done: {r2['page_count']} pages, QA={r2['qa']['passed']}")

    # === Example 3: Sales Pitch Premium (warm-elegant, Sales Pitch, 7 pages) ===
    print("Generating 3/6: Premium Sales Pitch...")
    seed_kw3 = "luxury premium warm elegant gold"
    _seed_image_cache(cache_dir, seed_kw3, "curiosity", "hook", 144, 48, colors["warm-elegant"], "Brand Logo")

    content_file3 = os.path.join(cache_dir, "sales_content.json")
    with open(content_file3, "w", encoding="utf-8") as f:
        json.dump({
            "company": "Aurum Partners",
            "product": "Premium Advisory Service",
            "tagline": "Your wealth, our wisdom",
            "metrics": {"aum": "$2.4B", "satisfaction": "99%", "growth": "28%", "clients": "400+"},
        }, f, ensure_ascii=False)

    r3 = generate_ppt(
        query="Premium Advisory Sales Pitch",
        strategy="Sales Pitch",
        theme="warm-elegant",
        content_file=content_file3,
        image_mode="search",
        image_config={"image_cache_dir": cache_dir},
        output=os.path.join(EXAMPLES_DIR, "03-premium-sales-pitch.pptx"),
    )
    examples.append(("03-premium-sales-pitch.pptx", "Premium Sales Pitch", "Sales Pitch", "Warm Elegant", r3))
    print(f"  Done: {r3['page_count']} pages, QA={r3['qa']['passed']}")

    # === Example 4: Tech Startup Vibrant (vibrant-startup, 12 pages) ===
    print("Generating 4/6: Tech Startup Launch...")
    seed_kw4 = "startup launch vibrant purple pink bold"
    _seed_image_cache(cache_dir, seed_kw4, "curiosity", "hook", 144, 48, colors["vibrant-startup"], "Launch Logo")
    _seed_image_cache(cache_dir, seed_kw4, "trust", "demo", 768, 720, colors["vibrant-startup"], "App Screenshot")

    content_file4 = os.path.join(cache_dir, "startup_content.json")
    with open(content_file4, "w", encoding="utf-8") as f:
        json.dump({
            "company": "NovaPay",
            "product": "NovaPay App",
            "tagline": "Payments at the speed of now",
            "metrics": {"transactions": "10M+", "users": "2M+", "nps": "82", "markets": "15"},
            "pain_points": [
                {"title": "Slow Payments", "desc": "Cross-border takes 3-5 days"},
                {"title": "High Fees", "desc": "Banks charge 5-8% for international"},
                {"title": "No Transparency", "desc": "Users can't track their money"},
            ],
            "chart_data": {"growth": {"labels": ["Q1", "Q2", "Q3", "Q4"], "values": [100, 340, 890, 2100]}},
        }, f, ensure_ascii=False)

    r4 = generate_ppt(
        query="Tech Startup Launch Presentation",
        strategy="YC Seed Deck",
        slides=12,
        theme="vibrant-startup",
        content_file=content_file4,
        image_mode="search",
        image_config={"image_cache_dir": cache_dir},
        output=os.path.join(EXAMPLES_DIR, "04-tech-startup-launch.pptx"),
    )
    examples.append(("04-tech-startup-launch.pptx", "Tech Startup Launch", "YC Seed Deck (12p)", "Vibrant Startup", r4))
    print(f"  Done: {r4['page_count']} pages, QA={r4['qa']['passed']}")

    # === Example 5: Sustainability Report (nature-calm, YC Seed, 10 pages) ===
    print("Generating 5/6: Sustainability Report...")
    seed_kw5 = "nature green sustainable calm forest"
    _seed_image_cache(cache_dir, seed_kw5, "curiosity", "hook", 144, 48, colors["nature-calm"], "Eco Logo")
    _seed_image_cache(cache_dir, seed_kw5, "confidence", "traction", 768, 720, colors["nature-calm"], "Nature Impact")

    content_file5 = os.path.join(cache_dir, "sustainability_content.json")
    with open(content_file5, "w", encoding="utf-8") as f:
        json.dump({
            "company": "GreenPath",
            "product": "Carbon Tracking Platform",
            "tagline": "Every step towards zero carbon counts",
            "metrics": {"carbon_reduced": "45K tons", "companies": "2,800+", "countries": "32", "savings": "$12M"},
            "chart_data": {"impact": {"labels": ["2022", "2023", "2024", "2025"], "values": [5, 18, 42, 78]}},
        }, f, ensure_ascii=False)

    r5 = generate_ppt(
        query="Sustainability Impact Report",
        theme="nature-calm",
        content_file=content_file5,
        image_mode="search",
        image_config={"image_cache_dir": cache_dir},
        output=os.path.join(EXAMPLES_DIR, "05-sustainability-report.pptx"),
    )
    examples.append(("05-sustainability-report.pptx", "Sustainability Report", "YC Seed Deck", "Nature Calm", r5))
    print(f"  Done: {r5['page_count']} pages, QA={r5['qa']['passed']}")

    # === Example 6: Dark Tech Deep Dive (dark-tech, Product Demo, 8 pages with persist) ===
    print("Generating 6/6: Deep Tech Product Deep Dive...")
    seed_kw6 = "deep tech neural network dark futuristic"
    _seed_image_cache(cache_dir, seed_kw6, "curiosity", "hook", 144, 48, (99, 102, 241), "Neural Logo")
    _seed_image_cache(cache_dir, seed_kw6, "trust", "demo", 768, 720, (99, 102, 241), "Neural Interface")

    content_file6 = os.path.join(cache_dir, "deeptech_content.json")
    with open(content_file6, "w", encoding="utf-8") as f:
        json.dump({
            "company": "NeuralForge",
            "product": "NeuralForge SDK",
            "tagline": "Build AI that thinks with you",
            "metrics": {"models": "500+", "latency": "< 10ms", "accuracy": "99.7%", "developers": "12K"},
            "pain_points": [
                {"title": "Slow Training", "desc": "Weeks to train production models"},
                {"title": "GPU Costs", "desc": "Cloud training costs are exploding"},
                {"title": "Deployment Gap", "desc": "Lab models don't work in production"},
            ],
            "chart_data": {"perf": {"labels": ["Baseline", "v1.0", "v2.0", "v3.0"], "values": [45, 72, 88, 97]}},
        }, f, ensure_ascii=False)

    r6 = generate_ppt(
        query="Deep Tech Product Deep Dive",
        strategy="Product Demo",
        theme="dark-tech",
        content_file=content_file6,
        image_mode="search",
        image_config={"image_cache_dir": cache_dir},
        persist=True,
        output=os.path.join(EXAMPLES_DIR, "06-deep-tech-product-dive.pptx"),
    )
    examples.append(("06-deep-tech-product-dive.pptx", "Deep Tech Product Dive", "Product Demo", "Dark Tech", r6))
    print(f"  Done: {r6['page_count']} pages, QA={r6['qa']['passed']}")

    # === Summary ===
    print("\n" + "=" * 70)
    print("SHOWCASE GENERATION COMPLETE")
    print("=" * 70)
    print(f"{'File':<40} {'Strategy':<20} {'Theme':<20} {'Pages':>5} {'Size':>8}")
    print("-" * 70)
    for filename, title, strategy, theme, result in examples:
        size = os.path.getsize(result["output_path"]) / 1024
        print(f"{filename:<40} {strategy:<20} {theme:<20} {result['page_count']:>5} {size:>7.1f}KB")
    print("=" * 70)


if __name__ == "__main__":
    generate_all()
