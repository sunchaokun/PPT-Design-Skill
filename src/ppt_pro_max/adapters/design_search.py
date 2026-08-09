"""Integrated design knowledge search engine for PPT generation.

BM25-based search across 7 PPT-relevant CSV datasets:
  colors, typography, styles, products, landing, ui-reasoning, motion

This module replaces the external ui-ux-pro-max dependency.
Data is loaded from ppt_pro_max/data/ bundled with the package.
"""

from __future__ import annotations

import csv
import re
from collections import defaultdict
from math import log
from pathlib import Path
from typing import Any

DATA_DIR = Path(__file__).parent.parent / "data"

PPT_CSV_CONFIG: dict[str, dict[str, Any]] = {
    "style": {
        "file": "styles.csv",
        "search_cols": ["Style Category", "Keywords", "Best For", "Type", "AI Prompt Keywords"],
        "output_cols": [
            "Style Category", "Type", "Keywords", "Primary Colors",
            "Effects & Animation", "Best For", "Light Mode ✓", "Dark Mode ✓",
            "Performance", "Accessibility", "Complexity", "AI Prompt Keywords",
        ],
    },
    "color": {
        "file": "colors.csv",
        "search_cols": ["Product Type", "Notes"],
        "output_cols": [
            "Product Type", "Primary", "On Primary", "Secondary", "On Secondary",
            "Accent", "On Accent", "Background", "Foreground", "Card",
            "Card Foreground", "Muted", "Muted Foreground", "Border",
            "Destructive", "On Destructive", "Ring", "Notes",
        ],
    },
    "landing": {
        "file": "landing.csv",
        "search_cols": ["Pattern Name", "Keywords", "Conversion Optimization", "Section Order"],
        "output_cols": [
            "Pattern Name", "Keywords", "Section Order",
            "Primary CTA Placement", "Color Strategy", "Conversion Optimization",
        ],
    },
    "product": {
        "file": "products.csv",
        "search_cols": ["Product Type", "Keywords", "Primary Style Recommendation", "Key Considerations"],
        "output_cols": [
            "Product Type", "Keywords", "Primary Style Recommendation",
            "Secondary Styles", "Landing Page Pattern",
            "Dashboard Style (if applicable)", "Color Palette Focus",
        ],
    },
    "typography": {
        "file": "typography.csv",
        "search_cols": ["Font Pairing Name", "Category", "Mood/Style Keywords", "Best For", "Heading Font", "Body Font"],
        "output_cols": [
            "Font Pairing Name", "Category", "Heading Font", "Body Font",
            "Mood/Style Keywords", "Best For", "Google Fonts URL", "CSS Import", "Notes",
        ],
    },
    "motion": {
        "file": "motion.csv",
        "search_cols": ["Category", "Intensity Tier", "Keywords", "Trigger"],
        "output_cols": [
            "Category", "Intensity Tier", "Trigger", "Duration", "Easing",
            "GSAP Snippet", "Framework Notes", "Do", "Don't", "Performance Notes",
        ],
    },
}

_PPT_DOMAIN_KEYWORDS: dict[str, list[str]] = {
    "color": ["color", "palette", "hex", "#", "rgb", "token", "semantic", "accent", "destructive", "muted", "foreground"],
    "landing": ["landing", "page", "cta", "conversion", "hero", "testimonial", "pricing", "section"],
    "product": [
        "saas", "ecommerce", "e-commerce", "fintech", "healthcare", "gaming",
        "portfolio", "crypto", "dashboard", "fitness", "restaurant", "hotel",
        "travel", "music", "education", "learning", "legal", "insurance",
        "medical", "beauty", "pharmacy", "dental", "pet", "dating", "wedding",
        "recipe", "delivery", "ride", "booking", "calendar", "timer", "tracker",
        "diary", "note", "chat", "messenger", "crm", "invoice", "parking",
        "transit", "vpn", "alarm", "weather", "sleep", "meditation", "fasting",
        "habit", "grocery", "meme", "wardrobe", "plant care", "reading",
        "flashcard", "puzzle", "trivia", "arcade", "photography", "streaming",
        "podcast", "newsletter", "marketplace", "freelancer", "coworking",
        "airline", "museum", "theater", "church", "non-profit", "charity",
        "kindergarten", "daycare", "senior care", "veterinary", "florist",
        "bakery", "brewery", "construction", "automotive", "real estate",
        "logistics", "agriculture", "coding bootcamp",
    ],
    "style": [
        "style", "design", "ui", "minimalism", "glassmorphism", "neumorphism",
        "brutalism", "dark mode", "flat", "aurora", "prompt", "css",
        "implementation", "variable", "checklist", "tailwind",
    ],
    "typography": ["font pairing", "typography pairing", "heading font", "body font"],
    "motion": ["gsap", "quickto", "scrolltrigger", "stagger", "parallax", "page transition", "scroll reveal", "animation"],
}

_WEB_ONLY_RULES: set[str] = {
    "cursor-pointer", "touch target", "44px", "hover state",
    "responsive", "375px", "768px", "1024px", "1440px",
    "prefers-reduced-motion", "aria-", "focus state", "keyboard nav",
    "scrollbar", "viewport", "breakpoint", "media query",
    "skeleton loader", "lazy load", "bundle", "rerender",
    "server component", "dynamic import", "code splitting",
}


class BM25:
    def __init__(self, k1: float = 1.5, b: float = 0.75):
        self.k1 = k1
        self.b = b
        self.corpus: list[list[str]] = []
        self.doc_lengths: list[int] = []
        self.avgdl: float = 0.0
        self.idf: dict[str, float] = {}
        self.doc_freqs: dict[str, int] = defaultdict(int)
        self.N: int = 0

    def tokenize(self, text: str) -> list[str]:
        text = re.sub(r"[^\w\s]", " ", str(text).lower())
        return [w for w in text.split() if len(w) >= 2]

    def fit(self, documents: list[str]) -> None:
        self.corpus = [self.tokenize(doc) for doc in documents]
        self.N = len(self.corpus)
        if self.N == 0:
            return
        self.doc_lengths = [len(doc) for doc in self.corpus]
        self.avgdl = sum(self.doc_lengths) / self.N

        for doc in self.corpus:
            seen: set[str] = set()
            for word in doc:
                if word not in seen:
                    self.doc_freqs[word] += 1
                    seen.add(word)

        for word, freq in self.doc_freqs.items():
            self.idf[word] = log((self.N - freq + 0.5) / (freq + 0.5) + 1)

    def score(self, query: str) -> list[tuple[int, float]]:
        query_tokens = self.tokenize(query)
        scores: list[tuple[int, float]] = []

        for idx, doc in enumerate(self.corpus):
            score = 0.0
            doc_len = self.doc_lengths[idx]
            term_freqs: dict[str, int] = defaultdict(int)
            for word in doc:
                term_freqs[word] += 1

            for token in query_tokens:
                if token in self.idf:
                    tf = term_freqs[token]
                    idf = self.idf[token]
                    numerator = tf * (self.k1 + 1)
                    denominator = tf + self.k1 * (1 - self.b + self.b * doc_len / self.avgdl)
                    score += idf * numerator / denominator

            scores.append((idx, score))

        return sorted(scores, key=lambda x: x[1], reverse=True)


def _load_csv(filepath: Path) -> list[dict[str, str]]:
    if not filepath.exists():
        return []
    with open(filepath, "r", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def _search_csv(
    filepath: Path,
    search_cols: list[str],
    output_cols: list[str],
    query: str,
    max_results: int,
) -> list[dict[str, str]]:
    data = _load_csv(filepath)
    if not data:
        return []

    documents = [" ".join(str(row.get(col, "")) for col in search_cols) for row in data]
    bm25 = BM25()
    bm25.fit(documents)
    ranked = bm25.score(query)

    results: list[dict[str, str]] = []
    for idx, score in ranked[:max_results]:
        if score > 0:
            row = data[idx]
            results.append({col: row.get(col, "") for col in output_cols if col in row})

    return results


def detect_domain(query: str) -> str:
    query_lower = query.lower()
    scores: dict[str, int] = {}
    for domain, keywords in _PPT_DOMAIN_KEYWORDS.items():
        scores[domain] = sum(1 for kw in keywords if re.search(r"\b" + re.escape(kw) + r"\b", query_lower))
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "style"


def search(query: str, domain: str | None = None, max_results: int = 3) -> list[dict[str, str]]:
    if domain is None:
        domain = detect_domain(query)

    config = PPT_CSV_CONFIG.get(domain)
    if config is None:
        return []

    filepath = DATA_DIR / config["file"]
    return _search_csv(filepath, config["search_cols"], config["output_cols"], query, max_results)


DIAL_TIERS: dict[str, list[tuple[int, int, dict[str, Any]]]] = {
    "variance": [
        (1, 3, {"label": "Centered / Minimal", "style_keywords": ["Minimalism", "Exaggerated Minimalism", "centered", "symmetric", "grid-based"]}),
        (4, 7, {"label": "Balanced / Modern", "style_keywords": ["modern", "structured", "balanced"]}),
        (8, 10, {"label": "Bold / Asymmetric", "style_keywords": ["Brutalism", "Bento Grids", "asymmetric", "experimental"]}),
    ],
    "motion": [
        (1, 3, {"label": "Subtle", "tier": "Subtle"}),
        (4, 7, {"label": "Standard", "tier": "Standard"}),
        (8, 10, {"label": "Complex", "tier": "Complex"}),
    ],
    "density": [
        (1, 3, {"label": "Spacious", "spacing": {"xs": "4px", "sm": "8px", "md": "24px", "lg": "32px", "xl": "48px", "2xl": "64px", "3xl": "96px"}}),
        (4, 7, {"label": "Standard", "spacing": {"xs": "4px", "sm": "8px", "md": "16px", "lg": "24px", "xl": "32px", "2xl": "48px", "3xl": "64px"}}),
        (8, 10, {"label": "Dense / Dashboard", "spacing": {"xs": "2px", "sm": "4px", "md": "8px", "lg": "12px", "xl": "16px", "2xl": "24px", "3xl": "32px"}}),
    ],
}


def _resolve_dial(dial_name: str, value: int | None) -> dict[str, Any] | None:
    if value is None:
        return None
    value = max(1, min(10, int(value)))
    for lo, hi, info in DIAL_TIERS[dial_name]:
        if lo <= value <= hi:
            return {**info, "value": value}
    return None


class DesignSystemGenerator:
    def __init__(self) -> None:
        self.reasoning_data = self._load_reasoning()

    def _load_reasoning(self) -> list[dict[str, str]]:
        filepath = DATA_DIR / "ui-reasoning.csv"
        if not filepath.exists():
            return []
        with open(filepath, "r", encoding="utf-8") as f:
            return list(csv.DictReader(f))

    def _find_reasoning_rule(self, category: str) -> dict[str, str]:
        category_lower = category.lower()

        for rule in self.reasoning_data:
            if rule.get("UI_Category", "").lower() == category_lower:
                return rule

        for rule in self.reasoning_data:
            ui_cat = rule.get("UI_Category", "").lower()
            if ui_cat in category_lower or category_lower in ui_cat:
                return rule

        for rule in self.reasoning_data:
            ui_cat = rule.get("UI_Category", "").lower()
            keywords = ui_cat.replace("/", " ").replace("-", " ").split()
            if any(kw in category_lower for kw in keywords):
                return rule

        return {}

    def apply_reasoning(self, category: str) -> dict[str, Any]:
        rule = self._find_reasoning_rule(category)

        if not rule:
            return {
                "pattern": "Hero + Features + CTA",
                "style_priority": ["Minimalism", "Flat Design"],
                "color_mood": "Professional",
                "typography_mood": "Clean",
                "key_effects": "Subtle hover transitions",
                "anti_patterns": "",
                "decision_rules": {},
                "severity": "MEDIUM",
            }

        import json

        decision_rules: dict[str, Any] = {}
        try:
            decision_rules = json.loads(rule.get("Decision_Rules", "{}"))
        except json.JSONDecodeError:
            pass

        return {
            "pattern": rule.get("Recommended_Pattern", ""),
            "style_priority": [s.strip() for s in rule.get("Style_Priority", "").split("+")],
            "color_mood": rule.get("Color_Mood", ""),
            "typography_mood": rule.get("Typography_Mood", ""),
            "key_effects": rule.get("Key_Effects", ""),
            "anti_patterns": rule.get("Anti_Patterns", ""),
            "decision_rules": decision_rules,
            "severity": rule.get("Severity", "MEDIUM"),
        }

    def _select_best_match(self, results: list[dict[str, str]], priority_keywords: list[str]) -> dict[str, str]:
        if not results:
            return {}
        if not priority_keywords:
            return results[0]

        for priority in priority_keywords:
            priority_lower = priority.lower().strip()
            for result in results:
                style_name = result.get("Style Category", "").lower()
                if priority_lower in style_name or style_name in priority_lower:
                    return result

        scored: list[tuple[int, dict[str, str]]] = []
        for result in results:
            result_str = str(result).lower()
            score = 0
            for kw in priority_keywords:
                kw_lower = kw.lower().strip()
                if kw_lower in result.get("Style Category", "").lower():
                    score += 10
                elif kw_lower in result.get("Keywords", "").lower():
                    score += 3
                elif kw_lower in result_str:
                    score += 1
            scored.append((score, result))

        scored.sort(key=lambda x: x[0], reverse=True)
        return scored[0][1] if scored and scored[0][0] > 0 else results[0]

    def generate(
        self,
        query: str,
        variance: int | None = None,
        motion: int | None = None,
        density: int | None = None,
    ) -> dict[str, Any]:
        variance_info = _resolve_dial("variance", variance)
        motion_info = _resolve_dial("motion", motion)
        density_info = _resolve_dial("density", density)

        product_results = search(query, "product", 1)
        category = "General"
        if product_results:
            category = product_results[0].get("Product Type", "General")

        reasoning = self.apply_reasoning(category)
        style_priority = reasoning.get("style_priority", [])

        effective_style_priority = style_priority
        if variance_info:
            effective_style_priority = variance_info["style_keywords"] + style_priority

        style_results = search(" ".join([query] + effective_style_priority[:2]), "style", 3)
        color_results = search(query, "color", 2)
        typography_results = search(query, "typography", 2)
        landing_results = search(query, "landing", 2)

        best_style = self._select_best_match(style_results, effective_style_priority)
        best_color = color_results[0] if color_results else {}
        best_typography = typography_results[0] if typography_results else {}
        best_landing = landing_results[0] if landing_results else {}

        motion_snippet: dict[str, str] = {}
        if motion_info:
            motion_results = search(f"{query} {motion_info['tier']}", "motion", 5)
            tiered = [m for m in motion_results if m.get("Intensity Tier") == motion_info["tier"]]
            if tiered:
                motion_snippet = tiered[0]
            elif motion_results:
                motion_snippet = motion_results[0]

        style_effects = best_style.get("Effects & Animation", "")
        reasoning_effects = reasoning.get("key_effects", "")
        combined_effects = style_effects if style_effects else reasoning_effects

        return {
            "project_name": query.upper(),
            "category": category,
            "pattern": {
                "name": best_landing.get("Pattern Name", reasoning.get("pattern", "Hero + Features + CTA")),
                "sections": best_landing.get("Section Order", "Hero > Features > CTA"),
                "cta_placement": best_landing.get("Primary CTA Placement", "Above fold"),
                "color_strategy": best_landing.get("Color Strategy", ""),
                "conversion": best_landing.get("Conversion Optimization", ""),
            },
            "style": {
                "name": best_style.get("Style Category", "Minimalism"),
                "type": best_style.get("Type", "General"),
                "effects": style_effects,
                "keywords": best_style.get("Keywords", ""),
                "best_for": best_style.get("Best For", ""),
                "performance": best_style.get("Performance", ""),
                "accessibility": best_style.get("Accessibility", ""),
                "light_mode": best_style.get("Light Mode ✓", ""),
                "dark_mode": best_style.get("Dark Mode ✓", ""),
            },
            "colors": {
                "primary": best_color.get("Primary", "#2563EB"),
                "on_primary": best_color.get("On Primary", ""),
                "secondary": best_color.get("Secondary", "#3B82F6"),
                "accent": best_color.get("Accent", "#F97316"),
                "background": best_color.get("Background", "#F8FAFC"),
                "foreground": best_color.get("Foreground", "#1E293B"),
                "muted": best_color.get("Muted", ""),
                "border": best_color.get("Border", ""),
                "destructive": best_color.get("Destructive", ""),
                "ring": best_color.get("Ring", ""),
                "notes": best_color.get("Notes", ""),
                "cta": best_color.get("Accent", "#F97316"),
                "text": best_color.get("Foreground", "#1E293B"),
            },
            "typography": {
                "heading": best_typography.get("Heading Font", "Inter"),
                "body": best_typography.get("Body Font", "Inter"),
                "mood": best_typography.get("Mood/Style Keywords", reasoning.get("typography_mood", "")),
                "best_for": best_typography.get("Best For", ""),
                "google_fonts_url": best_typography.get("Google Fonts URL", ""),
                "css_import": best_typography.get("CSS Import", ""),
            },
            "key_effects": combined_effects,
            "anti_patterns": reasoning.get("anti_patterns", ""),
            "decision_rules": reasoning.get("decision_rules", {}),
            "severity": reasoning.get("severity", "MEDIUM"),
            "dials": {
                "variance": variance_info["value"] if variance_info else None,
                "variance_label": variance_info["label"] if variance_info else None,
                "motion": motion_info["value"] if motion_info else None,
                "motion_label": motion_info["label"] if motion_info else None,
                "density": density_info["value"] if density_info else None,
                "density_label": density_info["label"] if density_info else None,
            },
            "motion_snippet": motion_snippet,
            "spacing_scale": density_info["spacing"] if density_info else None,
        }


_gen = DesignSystemGenerator()


def get_design_system(query: str, **kwargs: Any) -> dict[str, Any]:
    ds = _gen.generate(
        query,
        variance=kwargs.get("variance"),
        motion=kwargs.get("motion"),
        density=kwargs.get("density"),
    )
    return _normalize_design_system(ds)


def _normalize_design_system(ds: dict[str, Any]) -> dict[str, Any]:
    colors = ds.get("colors", {})
    normalized_colors: dict[str, Any] = {}
    for key, val in colors.items():
        nk = key.replace(" ", "-").replace("/", "-").lower()
        normalized_colors[nk] = val

    if "primary" not in normalized_colors and "primary" in colors:
        normalized_colors["primary"] = colors["primary"]

    ds["colors"] = normalized_colors

    style = ds.get("style", {})
    if style:
        ds["style_name"] = style.get("name", "")
        ds["style_effects"] = style.get("effects", "")
        ds["style_keywords"] = style.get("keywords", "")
        ds["style_best_for"] = style.get("best_for", "")
        ds["style_dark_mode"] = style.get("dark_mode", "")
        ds["style_light_mode"] = style.get("light_mode", "")

    pattern = ds.get("pattern", {})
    if pattern:
        ds["pattern_name"] = pattern.get("name", "")
        ds["pattern_sections"] = pattern.get("sections", "")
        ds["pattern_cta_placement"] = pattern.get("cta_placement", "")
        ds["pattern_color_strategy"] = pattern.get("color_strategy", "")
        ds["pattern_conversion"] = pattern.get("conversion", "")

    return ds


def is_available() -> bool:
    return DATA_DIR.exists() and (DATA_DIR / "styles.csv").exists()
