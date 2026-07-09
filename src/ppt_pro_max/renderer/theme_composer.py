"""Theme Composer — infinite design combinations from composable atoms.

Design atoms:
  - ColorPalette: 25+ curated palettes (mood + industry based)
  - FontPair: 20+ heading+body font combinations
  - DecorationStyle: 10+ visual decoration patterns
  - LayoutVariant: 8+ structural layout modifications
  - Mood: emotional tone that biases all atom selections

Usage:
  composer = ThemeComposer()
  theme = composer.compose("minimalist fintech pitch with warm tones")
  theme = composer.compose(style="dark-tech")  # preset still works
  theme = composer.compose(palette="ocean-depth", fonts="serif-editorial", decoration="accent-bar", layout="sidebar-nav")
"""

from __future__ import annotations

import re
import random
from dataclasses import dataclass, field
from typing import Any


# ============================================================
# COLOR PALETTES — 25 curated palettes
# ============================================================

COLOR_PALETTES: dict[str, dict[str, str]] = {
    "ocean-blue": {
        "primary": "#1E5AB8", "on-primary": "#FFFFFF", "secondary": "#64748B",
        "accent": "#0096C7", "background": "#FFFFFF", "foreground": "#0A1E3D",
        "muted": "#F0F4F8", "muted-foreground": "#6B7B8D", "border": "#DEE5EF",
        "destructive": "#EF4444",
    },
    "midnight-navy": {
        "primary": "#1E3A5F", "on-primary": "#FFFFFF", "secondary": "#8A9BB5",
        "accent": "#E8A838", "background": "#0A1E3D", "foreground": "#F0F4F8",
        "muted": "#1A2E4A", "muted-foreground": "#8A9BB5", "border": "#2A3E5A",
        "destructive": "#FF6B6B",
    },
    "cyber-neon": {
        "primary": "#6366F1", "on-primary": "#FFFFFF", "secondary": "#94A3B8",
        "accent": "#22D3EE", "background": "#060B18", "foreground": "#F8FAFC",
        "muted": "#0D152A", "muted-foreground": "#64748B", "border": "#1A2A4A",
        "destructive": "#FF0080",
    },
    "neon-gradient": {
        "primary": "#8B5CF6", "on-primary": "#FFFFFF", "secondary": "#A09CB0",
        "accent": "#FF2D87", "background": "#120C1E", "foreground": "#F8FAFC",
        "muted": "#1E1435", "muted-foreground": "#A09CB0", "border": "#2A1E4A",
        "destructive": "#FF4444",
    },
    "golden-luxury": {
        "primary": "#C99A4E", "on-primary": "#FFFFFF", "secondary": "#9A8C7E",
        "accent": "#D4A853", "background": "#FAF3E8", "foreground": "#2C2C2C",
        "muted": "#FEF3C7", "muted-foreground": "#9A8C7E", "border": "#E7E5E4",
        "destructive": "#DC2626",
    },
    "rose-gold": {
        "primary": "#B76E79", "on-primary": "#FFFFFF", "secondary": "#8B7E74",
        "accent": "#E8B4B8", "background": "#FFF5F5", "foreground": "#3D2C2C",
        "muted": "#FDE8E8", "muted-foreground": "#A89898", "border": "#E8D5D5",
        "destructive": "#DC2626",
    },
    "forest-green": {
        "primary": "#1B5E20", "on-primary": "#FFFFFF", "secondary": "#6B8E6B",
        "accent": "#4CAF50", "background": "#F5F7F0", "foreground": "#0D3B11",
        "muted": "#E8F0E5", "muted-foreground": "#6B8E6B", "border": "#C8DCC5",
        "destructive": "#DC2626",
    },
    "sage-calm": {
        "primary": "#5B7553", "on-primary": "#FFFFFF", "secondary": "#8B9E84",
        "accent": "#8BC38A", "background": "#F4F7F1", "foreground": "#2D3B28",
        "muted": "#E5EDE2", "muted-foreground": "#8B9E84", "border": "#D0DCCE",
        "destructive": "#DC2626",
    },
    "sunset-warm": {
        "primary": "#D97706", "on-primary": "#FFFFFF", "secondary": "#92400E",
        "accent": "#F59E0B", "background": "#FFFBEB", "foreground": "#1C1917",
        "muted": "#FEF3C7", "muted-foreground": "#A8A29E", "border": "#E7E5E4",
        "destructive": "#DC2626",
    },
    "terracotta": {
        "primary": "#C4704B", "on-primary": "#FFFFFF", "secondary": "#8B6B5E",
        "accent": "#E8926C", "background": "#FBF5F0", "foreground": "#3D2B1F",
        "muted": "#F0E0D5", "muted-foreground": "#9B8B7E", "border": "#DDD0C5",
        "destructive": "#DC2626",
    },
    "cherry-red": {
        "primary": "#DC2626", "on-primary": "#FFFFFF", "secondary": "#991B1B",
        "accent": "#F87171", "background": "#FEF2F2", "foreground": "#1C1917",
        "muted": "#FEE2E2", "muted-foreground": "#B91C1C", "border": "#FECACA",
        "destructive": "#991B1B",
    },
    "royal-purple": {
        "primary": "#7C3AED", "on-primary": "#FFFFFF", "secondary": "#6D28D9",
        "accent": "#A78BFA", "background": "#F5F3FF", "foreground": "#1E1B4B",
        "muted": "#EDE9FE", "muted-foreground": "#8B5CF6", "border": "#DDD6FE",
        "destructive": "#EF4444",
    },
    "arctic-frost": {
        "primary": "#0EA5E9", "on-primary": "#FFFFFF", "secondary": "#0284C7",
        "accent": "#38BDF8", "background": "#F0F9FF", "foreground": "#0C4A6E",
        "muted": "#E0F2FE", "muted-foreground": "#0369A1", "border": "#BAE6FD",
        "destructive": "#EF4444",
    },
    "slate-minimal": {
        "primary": "#475569", "on-primary": "#FFFFFF", "secondary": "#64748B",
        "accent": "#94A3B8", "background": "#F8FAFC", "foreground": "#0F172A",
        "muted": "#F1F5F9", "muted-foreground": "#94A3B8", "border": "#E2E8F0",
        "destructive": "#EF4444",
    },
    "charcoal-bold": {
        "primary": "#1F2937", "on-primary": "#FFFFFF", "secondary": "#4B5563",
        "accent": "#F97316", "background": "#111827", "foreground": "#F9FAFB",
        "muted": "#1F2937", "muted-foreground": "#6B7280", "border": "#374151",
        "destructive": "#EF4444",
    },
    "coral-energy": {
        "primary": "#F97316", "on-primary": "#FFFFFF", "secondary": "#EA580C",
        "accent": "#FB923C", "background": "#FFF7ED", "foreground": "#431407",
        "muted": "#FFEDD5", "muted-foreground": "#C2410C", "border": "#FED7AA",
        "destructive": "#DC2626",
    },
    "teal-fresh": {
        "primary": "#0D9488", "on-primary": "#FFFFFF", "secondary": "#0F766E",
        "accent": "#2DD4BF", "background": "#F0FDFA", "foreground": "#134E4A",
        "muted": "#CCFBF1", "muted-foreground": "#14B8A6", "border": "#99F6E4",
        "destructive": "#EF4444",
    },
    "indigo-deep": {
        "primary": "#4338CA", "on-primary": "#FFFFFF", "secondary": "#3730A3",
        "accent": "#6366F1", "background": "#EEF2FF", "foreground": "#1E1B4B",
        "muted": "#E0E7FF", "muted-foreground": "#4F46E5", "border": "#C7D2FE",
        "destructive": "#EF4444",
    },
    "copper-industrial": {
        "primary": "#B87333", "on-primary": "#FFFFFF", "secondary": "#8B5E3C",
        "accent": "#D4956A", "background": "#2D2D2D", "foreground": "#E8E0D8",
        "muted": "#3D3D3D", "muted-foreground": "#9B8B7E", "border": "#4D4D4D",
        "destructive": "#FF6B6B",
    },
    "monochrome": {
        "primary": "#18181B", "on-primary": "#FFFFFF", "secondary": "#3F3F46",
        "accent": "#A1A1AA", "background": "#FAFAFA", "foreground": "#09090B",
        "muted": "#F4F4F5", "muted-foreground": "#71717A", "border": "#E4E4E7",
        "destructive": "#DC2626",
    },
    "monochrome-dark": {
        "primary": "#D4D4D8", "on-primary": "#18181B", "secondary": "#A1A1AA",
        "accent": "#F4F4F5", "background": "#09090B", "foreground": "#FAFAFA",
        "muted": "#18181B", "muted-foreground": "#71717A", "border": "#27272A",
        "destructive": "#EF4444",
    },
    "lavender-dream": {
        "primary": "#8B5CF6", "on-primary": "#FFFFFF", "secondary": "#7C3AED",
        "accent": "#C4B5FD", "background": "#FAF5FF", "foreground": "#3B0764",
        "muted": "#F3E8FF", "muted-foreground": "#9333EA", "border": "#E9D5FF",
        "destructive": "#EF4444",
    },
    "mint-fresh": {
        "primary": "#10B981", "on-primary": "#FFFFFF", "secondary": "#059669",
        "accent": "#34D399", "background": "#ECFDF5", "foreground": "#064E3B",
        "muted": "#D1FAE5", "muted-foreground": "#6B7280", "border": "#A7F3D0",
        "destructive": "#DC2626",
    },
    "wine-burgundy": {
        "primary": "#7F1D1D", "on-primary": "#FFFFFF", "secondary": "#991B1B",
        "accent": "#B91C1C", "background": "#1C1010", "foreground": "#FEF2F2",
        "muted": "#2D1A1A", "muted-foreground": "#9B7A7A", "border": "#3D2A2A",
        "destructive": "#EF4444",
    },
    "sky-bright": {
        "primary": "#0284C7", "on-primary": "#FFFFFF", "secondary": "#0369A1",
        "accent": "#38BDF8", "background": "#F0F9FF", "foreground": "#0C4A6E",
        "muted": "#E0F2FE", "muted-foreground": "#0EA5E9", "border": "#BAE6FD",
        "destructive": "#EF4444",
    },
}

# ============================================================
# FONT PAIRS — 20 curated heading+body combinations
# ============================================================

FONT_PAIRS: dict[str, dict[str, str]] = {
    "modern-sans": {"heading": "Inter", "body": "Inter"},
    "geometric-sans": {"heading": "Space Grotesk", "body": "Inter"},
    "bold-sans": {"heading": "Poppins", "body": "Inter"},
    "clean-corporate": {"heading": "Calibri", "body": "Calibri"},
    "serif-editorial": {"heading": "Georgia", "body": "Georgia"},
    "elegant-serif": {"heading": "Playfair Display", "body": "Inter"},
    "literary-serif": {"heading": "Lora", "body": "Inter"},
    "tech-mono": {"heading": "Consolas", "body": "Consolas"},
    "mono-clean": {"heading": "Consolas", "body": "Inter"},
    "swiss-style": {"heading": "Arial", "body": "Arial"},
    "humanist-sans": {"heading": "Segoe UI", "body": "Segoe UI"},
    "friendly-round": {"heading": "Nunito", "body": "Inter"},
    "sharp-modern": {"heading": "Montserrat", "body": "Inter"},
    "classic-formal": {"heading": "Times New Roman", "body": "Times New Roman"},
    "contrast-mix": {"heading": "Playfair Display", "body": "Calibri"},
    "tech-contrast": {"heading": "Space Grotesk", "body": "Consolas"},
    "warm-mix": {"heading": "Georgia", "body": "Calibri"},
    "startup-mix": {"heading": "Poppins", "body": "Segoe UI"},
    "minimal-mix": {"heading": "Inter", "body": "Calibri"},
    "editorial-mix": {"heading": "Georgia", "body": "Segoe UI"},
}

# ============================================================
# DECORATION STYLES — 10 visual decoration patterns
# ============================================================

DECORATION_STYLES: dict[str, dict[str, Any]] = {
    "accent-bar": {
        "name": "Accent Bar",
        "left_accent": True, "title_underline": True, "card_top_bar": True,
        "description": "Clean accent bars and underlines",
    },
    "neon-lines": {
        "name": "Neon Lines",
        "left_accent": True, "title_underline": True, "card_top_bar": True,
        "top_line": True, "bottom_line": True,
        "description": "Glowing neon accent lines",
    },
    "gold-trim": {
        "name": "Gold Trim",
        "left_accent": True, "title_underline": True, "card_top_bar": True,
        "top_line": True, "bottom_line": True,
        "description": "Elegant gold decorative lines",
    },
    "minimal-dots": {
        "name": "Minimal Dots",
        "left_accent": False, "title_underline": False, "card_top_bar": False,
        "bullet_style": "circle",
        "description": "Subtle dot bullets, no lines",
    },
    "diamond-bullets": {
        "name": "Diamond Bullets",
        "left_accent": False, "title_underline": True, "card_top_bar": True,
        "bullet_style": "diamond",
        "description": "Diamond-shaped bullet points",
    },
    "gradient-bar": {
        "name": "Gradient Bar",
        "left_accent": True, "title_underline": True, "card_top_bar": True,
        "gradient_accent": True,
        "description": "Gradient-colored accent bars",
    },
    "circle-accent": {
        "name": "Circle Accent",
        "left_accent": False, "title_underline": False, "card_top_bar": False,
        "circle_decoration": True,
        "description": "Circle decorative elements",
    },
    "sidebar-nav": {
        "name": "Sidebar Navigation",
        "left_accent": True, "title_underline": True, "card_top_bar": True,
        "sidebar": True,
        "description": "Left sidebar with section info",
    },
    "no-decoration": {
        "name": "No Decoration",
        "left_accent": False, "title_underline": False, "card_top_bar": False,
        "description": "Clean, no decorative elements",
    },
    "full-bleed-overlay": {
        "name": "Full Bleed Overlay",
        "left_accent": False, "title_underline": False, "card_top_bar": False,
        "dark_overlay": True,
        "description": "Full-bleed image with dark overlay",
    },
}

# ============================================================
# LAYOUT VARIANTS — 8 structural modifications
# ============================================================

LAYOUT_VARIANTS: dict[str, dict[str, Any]] = {
    "standard": {
        "name": "Standard",
        "content_margin_left": 0.9, "content_margin_right": 0.9,
        "title_alignment": "left", "card_style": "rounded",
        "description": "Standard left-aligned layout",
    },
    "centered": {
        "name": "Centered",
        "content_margin_left": 2.0, "content_margin_right": 2.0,
        "title_alignment": "center", "card_style": "rounded",
        "description": "Centered editorial layout",
    },
    "sidebar-left": {
        "name": "Left Sidebar",
        "sidebar_width": 3.5, "sidebar_side": "left",
        "content_margin_left": 4.2, "content_margin_right": 0.9,
        "title_alignment": "left", "card_style": "rounded",
        "description": "Left sidebar with section info",
    },
    "sidebar-right": {
        "name": "Right Sidebar",
        "sidebar_width": 3.5, "sidebar_side": "right",
        "content_margin_left": 0.9, "content_margin_right": 4.5,
        "title_alignment": "left", "card_style": "rounded",
        "description": "Right sidebar for stats/quotes",
    },
    "wide-cards": {
        "name": "Wide Cards",
        "content_margin_left": 0.6, "content_margin_right": 0.6,
        "title_alignment": "left", "card_style": "wide",
        "card_gap": 0.3,
        "description": "Wide card-based layout",
    },
    "grid-2x2": {
        "name": "2x2 Grid",
        "content_margin_left": 0.8, "content_margin_right": 0.8,
        "title_alignment": "left", "card_style": "grid",
        "grid_rows": 2, "grid_cols": 2,
        "description": "2x2 grid of metric cards",
    },
    "asymmetric": {
        "name": "Asymmetric",
        "content_margin_left": 0.8, "content_margin_right": 0.8,
        "title_alignment": "left", "card_style": "staggered",
        "description": "Asymmetric staggered layout",
    },
    "full-width": {
        "name": "Full Width",
        "content_margin_left": 0.5, "content_margin_right": 0.5,
        "title_alignment": "left", "card_style": "flat",
        "description": "Full-width edge-to-edge layout",
    },
}

# ============================================================
# MOOD KEYWORDS — natural language → atom selection hints
# ============================================================

_MOOD_PALETTE_MAP: dict[str, list[str]] = {
    "professional": ["ocean-blue", "midnight-navy", "slate-minimal", "arctic-frost"],
    "corporate": ["ocean-blue", "midnight-navy", "indigo-deep", "charcoal-bold"],
    "tech": ["cyber-neon", "neon-gradient", "charcoal-bold", "monochrome-dark"],
    "dark": ["cyber-neon", "neon-gradient", "charcoal-bold", "monochrome-dark", "copper-industrial", "wine-burgundy"],
    "warm": ["golden-luxury", "sunset-warm", "terracotta", "coral-energy", "rose-gold"],
    "elegant": ["golden-luxury", "rose-gold", "lavender-dream", "wine-burgundy"],
    "luxury": ["golden-luxury", "rose-gold", "wine-burgundy", "copper-industrial"],
    "vibrant": ["neon-gradient", "coral-energy", "cherry-red", "royal-purple"],
    "startup": ["neon-gradient", "royal-purple", "coral-energy", "arctic-frost"],
    "nature": ["forest-green", "sage-calm", "mint-fresh", "teal-fresh"],
    "calm": ["sage-calm", "arctic-frost", "sky-bright", "lavender-dream"],
    "minimal": ["slate-minimal", "monochrome", "arctic-frost", "ocean-blue"],
    "bold": ["cherry-red", "coral-energy", "charcoal-bold", "cyber-neon"],
    "fresh": ["mint-fresh", "teal-fresh", "arctic-frost", "sky-bright"],
    "industrial": ["copper-industrial", "charcoal-bold", "monochrome-dark", "terracotta"],
    "fintech": ["ocean-blue", "midnight-navy", "forest-green", "indigo-deep"],
    "health": ["mint-fresh", "teal-fresh", "sage-calm", "sky-bright"],
    "education": ["ocean-blue", "indigo-deep", "teal-fresh", "slate-minimal"],
    "creative": ["neon-gradient", "royal-purple", "lavender-dream", "coral-energy"],
    "sustainability": ["forest-green", "sage-calm", "mint-fresh", "teal-fresh"],
}

_MOOD_FONT_MAP: dict[str, list[str]] = {
    "professional": ["modern-sans", "clean-corporate", "swiss-style"],
    "corporate": ["clean-corporate", "modern-sans", "swiss-style"],
    "tech": ["tech-mono", "mono-clean", "geometric-sans"],
    "dark": ["tech-mono", "mono-clean", "geometric-sans"],
    "warm": ["serif-editorial", "warm-mix", "elegant-serif"],
    "elegant": ["elegant-serif", "serif-editorial", "literary-serif"],
    "luxury": ["elegant-serif", "serif-editorial", "contrast-mix"],
    "vibrant": ["bold-sans", "startup-mix", "sharp-modern"],
    "startup": ["bold-sans", "startup-mix", "humanist-sans"],
    "nature": ["literary-serif", "warm-mix", "humanist-sans"],
    "calm": ["humanist-sans", "minimal-mix", "literary-serif"],
    "minimal": ["modern-sans", "minimal-mix", "swiss-style"],
    "bold": ["bold-sans", "sharp-modern", "geometric-sans"],
    "fresh": ["humanist-sans", "friendly-round", "modern-sans"],
    "industrial": ["tech-contrast", "mono-clean", "swiss-style"],
    "fintech": ["clean-corporate", "modern-sans", "swiss-style"],
    "health": ["humanist-sans", "friendly-round", "modern-sans"],
    "education": ["clean-corporate", "editorial-mix", "modern-sans"],
    "creative": ["contrast-mix", "sharp-modern", "friendly-round"],
    "sustainability": ["humanist-sans", "warm-mix", "literary-serif"],
}

_MOOD_DECORATION_MAP: dict[str, list[str]] = {
    "professional": ["accent-bar", "sidebar-nav"],
    "corporate": ["accent-bar", "sidebar-nav", "gradient-bar"],
    "tech": ["neon-lines", "no-decoration"],
    "dark": ["neon-lines", "full-bleed-overlay", "no-decoration"],
    "warm": ["gold-trim", "diamond-bullets"],
    "elegant": ["gold-trim", "diamond-bullets", "circle-accent"],
    "luxury": ["gold-trim", "circle-accent"],
    "vibrant": ["gradient-bar", "neon-lines"],
    "startup": ["gradient-bar", "accent-bar"],
    "nature": ["circle-accent", "minimal-dots"],
    "calm": ["minimal-dots", "no-decoration", "circle-accent"],
    "minimal": ["no-decoration", "minimal-dots"],
    "bold": ["accent-bar", "gradient-bar"],
    "fresh": ["circle-accent", "accent-bar"],
    "industrial": ["accent-bar", "no-decoration"],
    "fintech": ["accent-bar", "sidebar-nav"],
    "health": ["circle-accent", "accent-bar"],
    "education": ["accent-bar", "minimal-dots"],
    "creative": ["gradient-bar", "circle-accent"],
    "sustainability": ["circle-accent", "minimal-dots"],
}

_MOOD_LAYOUT_MAP: dict[str, list[str]] = {
    "professional": ["sidebar-left", "grid-2x2", "standard"],
    "corporate": ["sidebar-left", "standard", "grid-2x2"],
    "tech": ["wide-cards", "standard", "full-width"],
    "dark": ["wide-cards", "full-width", "standard"],
    "warm": ["centered", "standard", "sidebar-right"],
    "elegant": ["centered", "standard", "sidebar-right"],
    "luxury": ["centered", "sidebar-right"],
    "vibrant": ["wide-cards", "grid-2x2", "asymmetric"],
    "startup": ["grid-2x2", "wide-cards", "asymmetric"],
    "nature": ["standard", "sidebar-left", "grid-2x2"],
    "calm": ["standard", "centered", "sidebar-left"],
    "minimal": ["standard", "centered", "no-decoration"],
    "bold": ["full-width", "asymmetric", "wide-cards"],
    "fresh": ["grid-2x2", "standard", "wide-cards"],
    "industrial": ["full-width", "standard", "sidebar-left"],
    "fintech": ["sidebar-left", "grid-2x2", "standard"],
    "health": ["grid-2x2", "standard", "sidebar-left"],
    "education": ["standard", "sidebar-left", "grid-2x2"],
    "creative": ["asymmetric", "wide-cards", "centered"],
    "sustainability": ["grid-2x2", "sidebar-left", "standard"],
}

# Preset theme → atom mapping (backward compatible)
_PRESET_ATOM_MAP: dict[str, dict[str, str]] = {
    "professional": {"palette": "midnight-navy", "fonts": "clean-corporate", "decoration": "accent-bar", "layout": "sidebar-left"},
    "dark-tech": {"palette": "cyber-neon", "fonts": "tech-mono", "decoration": "neon-lines", "layout": "wide-cards"},
    "warm-elegant": {"palette": "golden-luxury", "fonts": "serif-editorial", "decoration": "gold-trim", "layout": "centered"},
    "vibrant-startup": {"palette": "neon-gradient", "fonts": "bold-sans", "decoration": "gradient-bar", "layout": "grid-2x2"},
    "nature-calm": {"palette": "forest-green", "fonts": "humanist-sans", "decoration": "circle-accent", "layout": "sidebar-left"},
}


class ThemeComposer:
    """Compose infinite theme combinations from design atoms."""

    def compose(
        self,
        style: str | None = None,
        palette: str | None = None,
        fonts: str | None = None,
        decoration: str | None = None,
        layout: str | None = None,
        mood: str | None = None,
        seed: int | None = None,
    ) -> dict[str, Any]:
        if style and style in _PRESET_ATOM_MAP:
            atoms = _PRESET_ATOM_MAP[style]
            palette = palette or atoms.get("palette")
            fonts = fonts or atoms.get("fonts")
            decoration = decoration or atoms.get("decoration")
            layout = layout or atoms.get("layout")

        detected_moods = self._detect_moods(style or "") if style else []
        if mood:
            detected_moods = [mood] + detected_moods
        if not detected_moods:
            detected_moods = ["professional"]

        rng = random.Random(seed) if seed is not None else random.Random()

        p = palette or self._pick_from_mood(detected_moods, _MOOD_PALETTE_MAP, rng)
        f = fonts or self._pick_from_mood(detected_moods, _MOOD_FONT_MAP, rng)
        d = decoration or self._pick_from_mood(detected_moods, _MOOD_DECORATION_MAP, rng)
        l = layout or self._pick_from_mood(detected_moods, _MOOD_LAYOUT_MAP, rng)

        colors = dict(COLOR_PALETTES.get(p, COLOR_PALETTES["ocean-blue"]))
        typo = dict(FONT_PAIRS.get(f, FONT_PAIRS["modern-sans"]))
        deco = dict(DECORATION_STYLES.get(d, DECORATION_STYLES["accent-bar"]))
        lay = dict(LAYOUT_VARIANTS.get(l, LAYOUT_VARIANTS["standard"]))

        dark_mode = self._is_dark(colors)

        return {
            "name": f"{p}+{f}+{d}+{l}",
            "colors": colors,
            "typography": typo,
            "dark_mode": dark_mode,
            "decoration": deco,
            "layout_variant": lay,
            "atoms": {"palette": p, "fonts": f, "decoration": d, "layout": l, "moods": detected_moods},
        }

    def _detect_moods(self, text: str) -> list[str]:
        text_lower = text.lower()
        moods = []
        for mood in _MOOD_PALETTE_MAP:
            if mood in text_lower:
                moods.append(mood)
        industry_hints = {
            "fintech": "fintech", "finance": "fintech", "bank": "fintech", "invest": "professional",
            "health": "health", "medical": "health", "bio": "health",
            "edu": "education", "university": "education", "school": "education",
            "sustainability": "sustainability", "esg": "sustainability", "green": "nature",
            "creative": "creative", "design": "creative", "art": "creative",
            "saas": "tech", "ai": "tech", "ml": "tech", "cloud": "tech",
            "luxury": "luxury", "premium": "elegant", "brand": "elegant",
            "startup": "startup", "pitch": "startup", "fundrais": "startup",
        }
        for hint, mood in industry_hints.items():
            if hint in text_lower and mood not in moods:
                moods.append(mood)
        return moods

    def _pick_from_mood(self, moods: list[str], mood_map: dict, rng: random.Random) -> str:
        for mood in moods:
            options = mood_map.get(mood, [])
            if options:
                return rng.choice(options)
        all_options = list(mood_map.values())
        if all_options:
            return rng.choice(rng.choice(all_options))
        return list(mood_map.keys())[0] if mood_map else "standard"

    def _is_dark(self, colors: dict[str, str]) -> bool:
        bg = colors.get("background", "#FFFFFF")
        r, g, b = int(bg[1:3], 16), int(bg[3:5], 16), int(bg[5:7], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        return luminance < 0.5

    @staticmethod
    def available_palettes() -> list[str]:
        return list(COLOR_PALETTES.keys())

    @staticmethod
    def available_fonts() -> list[str]:
        return list(FONT_PAIRS.keys())

    @staticmethod
    def available_decorations() -> list[str]:
        return list(DECORATION_STYLES.keys())

    @staticmethod
    def available_layouts() -> list[str]:
        return list(LAYOUT_VARIANTS.keys())

    @staticmethod
    def available_presets() -> list[str]:
        return list(_PRESET_ATOM_MAP.keys())

    @staticmethod
    def combination_count() -> int:
        return len(COLOR_PALETTES) * len(FONT_PAIRS) * len(DECORATION_STYLES) * len(LAYOUT_VARIANTS)
