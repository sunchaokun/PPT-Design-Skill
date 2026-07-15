# PPT Design Quality Upgrade Plan

> From "Template-Grade" to "International Designer-Grade"
> Version: 1.1 | Date: 2026-07-14

---

## Executive Summary

Our PPT generation system has complete functional coverage (P1-P15, 638 tests passing), but the visual output lacks the sophistication of international design consultancies. The core problem is not missing features — it's **design intelligence**: every slide uses identical hardcoded positions, flat colors without depth, binary typography (bold/not-bold), and no content-adaptive layout. This plan addresses 7 design dimensions with 28 specific upgrades, organized into 3 implementation tiers.

**Root Cause Analysis — Why Output Looks "Cheap":**

| # | Root Cause | Impact | Current Code Location |
|---|-----------|--------|----------------------|
| 1 | All positions hardcoded in `render_slide()` | Every deck looks identical | `precision_renderer.py:364-421` |
| 2 | No content measurement before placement | Text overflows or wastes space | `precision_renderer.py:402-418` |
| 3 | Binary typography (bold/not-bold, 2 sizes) | No visual hierarchy | `precision_renderer.py:365-384` |
| 4 | Flat 10-color palettes, no tint/shade scales | No depth, no subtle backgrounds | `theme_composer.py:27-178` |
| 5 | Uniform shadow (blur=4, alpha=15-25%) on everything | Clip-art feel | `precision_renderer.py:193,214` |
| 6 | Left accent strip on every content slide | "AI-generated" watermark | `precision_renderer.py:555` |
| 7 | No image color-grading or gradient overlays | Images clash with palette | `precision_renderer.py:559-565` |

---

## Tier 1: Critical Visual Impact (Implement First)

These 10 upgrades will transform the output from "obviously auto-generated" to "professionally designed" in a single pass.

---

### 1.1 Dynamic Layout Engine — Content-Adaptive Positioning

**Problem:** `render_slide()` uses hardcoded coordinates for every goal type. A 3-bullet slide and a 10-bullet slide get identical layout.

**Solution:** Replace hardcoded positions with a constraint-based layout system that measures content before placing it.

**Implementation:**

```python
# New module: src/ppt_pro_max/renderer/layout_engine.py

@dataclass
class Rect:
    """Slide rectangle with margins and computed content area."""
    width: float = 13.333   # SLIDE_WIDTH in inches
    height: float = 7.5     # SLIDE_HEIGHT in inches
    margin_left: float = 0.9
    margin_right: float = 0.9
    margin_top: float = 0.6
    margin_bottom: float = 0.5
    
    @property
    def content_width(self) -> float:
        return self.width - self.margin_left - self.margin_right
    
    @property
    def content_height(self) -> float:
        return self.height - self.margin_top - self.margin_bottom

@dataclass
class ContentLayout:
    """Computed positions for all elements on a slide."""
    title_x: float = 0.9
    title_y: float = 0.5
    title_w: float = 11.533
    title_h: float = 0.6
    content_x: float = 0.9
    content_y: float = 1.6
    content_w: float = 7.0
    image_x: float = 8.3
    image_y: float = 1.2
    image_w: float = 4.2
    bullet_columns: int = 1  # 1 or 2, decided by LayoutEngine

class LayoutEngine:
    """Content-adaptive layout calculator.
    
    Measures text length, counts bullets, detects image presence,
    then computes optimal positions using a spacing scale.
    """
    
    SPACING_SCALE = {
        "xs": Pt(4),   # 4pt — tight element gaps
        "sm": Pt(8),   # 8pt — between related items
        "md": Pt(12),  # 12pt — between sections
        "lg": Pt(24),  # 24pt — major separations
        "xl": Pt(48),  # 48pt — slide-level breathing room
    }
    
    MARGIN_PROFILES = {
        "consulting": {"left": 0.9, "right": 0.9, "top": 0.6, "bottom": 0.5},
        "keynote":    {"left": 1.5, "right": 1.5, "top": 1.2, "bottom": 0.8},
        "minimal":    {"left": 2.0, "right": 2.0, "top": 1.5, "bottom": 1.0},
    }
    
    def compute_content_layout(self, page: dict, slide_rect: Rect) -> ContentLayout:
        """Returns positions for title, body, image, cards based on content."""
        title = page.get("title", "")
        bullets = page.get("bullets", [])
        has_image = bool(page.get("image"))
        cards = page.get("cards", [])
        
        # Measure title height
        title_lines = self._estimate_lines(title, slide_rect.content_width, 28)
        title_h = max(0.5, title_lines * 0.4)
        
        # Content zone starts after title + gap
        content_y = slide_rect.margin_top + title_h + 0.4
        
        # Bullet layout: narrow text block for few, two-column for many
        if len(bullets) <= 4:
            text_w = slide_rect.content_width * 0.65 if has_image else slide_rect.content_width
        elif len(bullets) <= 7:
            text_w = slide_rect.content_width * 0.7 if has_image else slide_rect.content_width
        else:
            text_w = slide_rect.content_width  # two-column mode
        
        # Image placement: right side if present, no image = full-width text
        if has_image:
            image_w = slide_rect.content_width - text_w - 0.4
            image_x = slide_rect.margin_left + text_w + 0.4
        else:
            image_w = 0
            image_x = 0
        
        return ContentLayout(
            title_x=slide_rect.margin_left,
            title_y=slide_rect.margin_top,
            title_w=slide_rect.content_width,
            title_h=title_h,
            content_x=slide_rect.margin_left,
            content_y=content_y,
            content_w=text_w,
            image_x=image_x,
            image_y=content_y,
            image_w=image_w,
        )
```

**Key Changes to `precision_renderer.py`:**

1. Add `LayoutEngine` as an instance attribute in `PrecisionRenderer.__init__()` — no new parameter on `render_slide()`
2. Inside `render_slide()`, call `self._layout_engine.compute_content_layout(page, slide_rect)` to get positions, then use them instead of hardcoded values
3. Title height adapts to title length (short title = more whitespace below)
4. Bullet width adapts to bullet count (3 bullets = narrow column, 8+ = full width or two-column)
5. Image position adapts to text width (no image = text expands to full width)
6. Card width adapts to card count AND card content length

**LayoutEngine Integration Pattern:**

```python
# In PrecisionRenderer.__init__():
from ppt_pro_max.renderer.layout_engine import LayoutEngine
self._layout_engine = LayoutEngine()

# In render_slide(), replace hardcoded positions:
slide_rect = Rect(margin_left=0.9, margin_right=0.9, margin_top=0.6, margin_bottom=0.5,
                  width=SLIDE_WIDTH, height=SLIDE_HEIGHT)
layout = self._layout_engine.compute_content_layout(page, slide_rect)

# Then use layout.title_x, layout.title_y, layout.title_w, layout.title_h,
# layout.content_x, layout.content_y, layout.content_w,
# layout.image_x, layout.image_y, layout.image_w instead of hardcoded values
```

**`_estimate_lines()` Implementation:**

```python
@staticmethod
def _estimate_lines(text: str, width_inches: float, font_size_pt: int) -> int:
    """Estimate number of lines for text at given width and font size.
    
    Uses character-count heuristic: average character width ≈ font_size_pt * 0.012 inches
    for Latin text, font_size_pt * 0.024 for CJK. Mixed text uses weighted average.
    """
    if not text:
        return 1
    # Count CJK vs Latin characters
    cjk_count = sum(1 for ch in text if '\u4e00' <= ch <= '\u9fff' or '\u3000' <= ch <= '\u303f')
    latin_count = len(text) - cjk_count
    # Average character width in inches
    avg_char_w = (latin_count * font_size_pt * 0.012 + cjk_count * font_size_pt * 0.024) / max(1, len(text))
    chars_per_line = max(1, int(width_inches / avg_char_w))
    lines = (len(text) + chars_per_line - 1) // chars_per_line
    return max(1, lines)
```

**Files to Modify:**
- `src/ppt_pro_max/renderer/layout_engine.py` (NEW — contains `Rect`, `ContentLayout`, `LayoutEngine`)
- `src/ppt_pro_max/enterprise/precision_renderer.py` (refactor `render_slide()`)

**Estimated Effort:** 3-4 days

---

### 1.2 Typographic Scale System — 5-Level Hierarchy

**Problem:** Current system uses 2 effective sizes (28pt title, 12pt body = 2.3x ratio). Professional decks use 4-5 levels with 4x+ ratio between hero and body.

**Solution:** Implement a strict typographic scale with 5 levels, minimum 2pt gaps between adjacent levels, and proper line-height/letter-spacing.

**Typographic Scale Definition:**

| Level | Name | Size Range | Weight | Line Height | Letter Spacing | Use |
|-------|------|-----------|--------|-------------|----------------|-----|
| 1 | Display | 44-72pt | Bold (700) | 1.1x | -0.02em | Hero statements, section numbers |
| 2 | Title | 24-28pt | Semibold (600) | 1.2x | -0.01em | Slide titles |
| 3 | Subtitle | 18-22pt | Medium (500) | 1.3x | 0em | Section headers, card titles |
| 4 | Body | 14-16pt | Regular (400) | 1.5x | 0em | Bullets, descriptions |
| 5 | Caption | 10-12pt | Regular (400) | 1.4x | +0.02em | Sources, footnotes, badges |

**Implementation:**

```python
# New module: src/ppt_pro_max/renderer/typography.py

from dataclasses import dataclass
from pptx.util import Pt

@dataclass
class TypeScale:
    display:  int = 52   # pt
    title:    int = 28
    subtitle: int = 20
    body:     int = 14
    caption:  int = 11
    
    # Line height multipliers
    display_lh:  float = 1.1
    title_lh:    float = 1.2
    subtitle_lh: float = 1.3
    body_lh:     float = 1.5
    caption_lh:  float = 1.4
    
    # Letter spacing: stored as EM ratio, converted to OOXML hundredths-of-a-point at render time
    # OOXML a:spc val = em_value * font_size_pt * 100
    # Example: -0.02em at 28pt = -0.02 * 28 * 100 = -56
    display_tracking:  float = -0.02   # -0.02em (tight)
    title_tracking:    float = -0.01   # -0.01em
    subtitle_tracking: float = 0.0
    body_tracking:     float = 0.0
    caption_tracking:  float = 0.02    # +0.02em (relaxed)
    
    @classmethod
    def for_density(cls, density: int) -> "TypeScale":
        """Scale all sizes based on content density (1-10).
        
        Enforces minimum sizes: body >= 11pt, caption >= 9pt.
        """
        factor = 1.0 - (density - 1) * 0.025  # density 1 = 1.0x, density 10 = 0.775x
        return cls(
            display=max(36, int(52 * factor)),
            title=max(20, int(28 * factor)),
            subtitle=max(16, int(20 * factor)),
            body=max(11, int(14 * factor)),
            caption=max(9, int(11 * factor)),
        )
    
    @classmethod
    def for_mode(cls, mode: str) -> "TypeScale":
        """Select scale based on presentation mode."""
        if mode == "presenting":
            return cls(display=64, title=32, subtitle=22, body=18, caption=12)
        elif mode == "reading":
            return cls(display=44, title=24, subtitle=18, body=14, caption=11)
        else:
            return cls()  # default balanced
```

**Key Changes:**

1. Replace all hardcoded font sizes in `precision_renderer.py` with `TypeScale` lookups
2. Add letter-spacing via OOXML `<a:spc>` element on runs (already supported by `visual_effects.py` patterns)
3. Add line-height via `<a:lnSpc><a:spcPct val="150000"/>` in paragraph properties
4. Hero title: 52pt → 44-72pt (dynamic based on title length — short title gets larger font)
5. Body text: 12pt → 14pt minimum (12pt is unreadable on projected screens)
6. Card title: 15pt → 20pt (subtitle level, not between body and caption)
7. Card body: 11pt → 14pt (body level, not caption level)
8. Exercise badge: 11pt → 11pt with ALL CAPS + +0.08em tracking
9. **shape_factory.py fix**: `add_hexagon_card()` at line 148 uses `Pt(font_size - 3)` with default `font_size=13`, producing **10pt** — below the 11pt minimum. Must change to `Pt(max(11, font_size - 3))`.

**Letter-Spacing Implementation (OOXML):**

```python
def apply_letter_spacing(run, tracking_em: float, font_size_pt: int):
    """Set letter spacing on a run.
    
    tracking_em: EM ratio, e.g. -0.02 (tight), +0.08 (wide for ALL CAPS)
    font_size_pt: font size in points (needed because OOXML spc is font-size-dependent)
    
    OOXML a:spc val is in hundredths of a point.
    val = tracking_em * font_size_pt * 100
    Example: -0.02em at 28pt = -0.02 * 28 * 100 = -56
    """
    rPr = run._r.get_or_add_rPr()
    spc = etree.SubElement(rPr, qn("a:spc"))
    spc.set("val", str(int(tracking_em * font_size_pt * 100)))
```

**Files to Modify:**
- `src/ppt_pro_max/renderer/typography.py` (NEW)
- `src/ppt_pro_max/enterprise/precision_renderer.py` (replace hardcoded sizes)
- `src/ppt_pro_max/renderer/visual_effects.py` (add `apply_letter_spacing()`)
- `src/ppt_pro_max/renderer/ppt_renderer.py` (replace hardcoded sizes)
- `src/ppt_pro_max/renderer/shape_factory.py` (replace hardcoded sizes)

**Estimated Effort:** 2-3 days

---

### 1.3 Color Depth System — Tint/Shade Scales + Alpha Variants

**Problem:** Each palette has 10 flat hex colors. No opacity variants, no tint/shade scales, no gradient definitions. Cards use only 2 background options (dark/light).

**Solution:** Generate 11 tint/shade variants per primary color algorithmically, plus 5 alpha levels per color.

**Color Scale Generation:**

```python
# New module: src/ppt_pro_max/renderer/color_system.py

def generate_color_scale(base_hex: str, levels: int = 11) -> dict[str, str]:
    """Generate perceptually uniform tint/shade scale.
    
    Uses OKLCH color space for perceptually uniform steps.
    No external dependency — implements pure-Python OKLCH conversion.
    Fallback to HSL if OKLCH produces out-of-gamut colors.
    
    Returns: {"50": "#...", "100": "#...", ..., "500": base, ..., "950": "#..."}
    """
    # Pure-Python OKLCH conversion (no external library needed)
    # Step 1: hex → linear RGB → OKLab → OKLCH
    l, c, h = _hex_to_oklch(base_hex)
    
    scale = {}
    for i, level in enumerate([50, 100, 200, 300, 400, 500, 600, 700, 800, 900, 950]):
        if level <= 500:
            # Tint: increase lightness toward white
            t = (500 - level) / 500  # 1.0 at 50, 0.0 at 500
            new_l = l + (1.0 - l) * t * 0.9  # Don't go full white
            new_c = c * (1 - t * 0.5)  # Reduce chroma for tints
        else:
            # Shade: decrease lightness toward black
            t = (level - 500) / 500  # 0.0 at 500, 1.0 at 950
            new_l = l * (1 - t * 0.8)  # Don't go full black
            new_c = c * (1 - t * 0.3)  # Slightly reduce chroma for shades
        
        # Convert back: OKLCH → OKLab → linear RGB → sRGB → hex
        # With gamut clipping to ensure valid sRGB values
        scale[str(level)] = _oklch_to_hex(new_l, new_c, h)
    
    return scale

def _hex_to_oklch(hex_color: str) -> tuple[float, float, float]:
    """Convert hex color to OKLCH. Pure Python, no dependencies."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255
    # sRGB → linear RGB
    r = r / 12.92 if r <= 0.04045 else ((r + 0.055) / 1.055) ** 2.4
    g = g / 12.92 if g <= 0.04045 else ((g + 0.055) / 1.055) ** 2.4
    b = b / 12.92 if b <= 0.04045 else ((b + 0.055) / 1.055) ** 2.4
    # linear RGB → OKLab (via LMS)
    l_ = 0.4122214708 * r + 0.5363325363 * g + 0.0514459929 * b
    m_ = 0.2119034982 * r + 0.6806995451 * g + 0.1073969566 * b
    s_ = 0.0883024619 * r + 0.2817188376 * g + 0.6299787005 * b
    l_ = l_ ** (1/3) if l_ >= 0 else -((-l_) ** (1/3))
    m_ = m_ ** (1/3) if m_ >= 0 else -((-m_) ** (1/3))
    s_ = s_ ** (1/3) if s_ >= 0 else -((-s_) ** (1/3))
    L = 0.2104542553 * l_ + 0.7936177850 * m_ - 0.0040720468 * s_
    a = 1.9779984951 * l_ - 2.4285922050 * m_ + 0.4505937099 * s_
    b_ok = 0.0259040371 * l_ + 0.7827717662 * m_ - 0.8086757660 * s_
    # OKLab → OKLCH
    C = (a ** 2 + b_ok ** 2) ** 0.5
    H = __import__('math').degrees(__import__('math').atan2(b_ok, a)) % 360
    return L, C, H

def _oklch_to_hex(L: float, C: float, H: float) -> str:
    """Convert OKLCH to hex. Pure Python, with gamut clipping."""
    import math
    H_rad = H * math.pi / 180
    a = C * math.cos(H_rad)
    b_ok = C * math.sin(H_rad)
    # OKLab → LMS
    l_ = L + 0.3963377774 * a + 0.2158037573 * b_ok
    m_ = L - 0.1055613458 * a - 0.0638541728 * b_ok
    s_ = L - 0.0894841775 * a - 1.2914855480 * b_ok
    l_ = l_ ** 3 if l_ >= 0 else -((-l_) ** 3)
    m_ = m_ ** 3 if m_ >= 0 else -((-m_) ** 3)
    s_ = s_ ** 3 if s_ >= 0 else -((-s_) ** 3)
    # LMS → linear RGB
    r = +4.0767416621 * l_ - 3.3077115913 * m_ + 0.2309699292 * s_
    g = -1.2684380046 * l_ + 2.6097574011 * m_ - 0.3413193965 * s_
    b = -0.0041960863 * l_ - 0.7034186147 * m_ + 1.7076147010 * s_
    # linear RGB → sRGB (gamut clip)
    r = max(0, min(1, 12.92 * r if r <= 0.0031308 else 1.055 * r ** (1/2.4) - 0.055))
    g = max(0, min(1, 12.92 * g if g <= 0.0031308 else 1.055 * g ** (1/2.4) - 0.055))
    b = max(0, min(1, 12.92 * b if b <= 0.0031308 else 1.055 * b ** (1/2.4) - 0.055))
    return f"{int(r*255):02X}{int(g*255):02X}{int(b*255):02X}"

def alpha_color(shape, hex_color: str, opacity_pct: int) -> None:
    """Apply alpha to a shape's solid fill.
    
    Direct wrapper around set_solid_fill_with_alpha() from visual_effects.py.
    opacity_pct: 0-100 where 100 = fully opaque, 10 = 10% visible.
    """
    from ppt_pro_max.renderer.visual_effects import set_solid_fill_with_alpha
    set_solid_fill_with_alpha(shape, hex_color, opacity_pct)

# Pre-defined alpha levels (in percent, compatible with set_solid_fill_with_alpha)
ALPHA_LEVELS = {
    "ghost":   4,    # Barely visible — card backgrounds, grouped areas
    "subtle":  8,    # Subtle tint — hover states, secondary backgrounds
    "light":   15,   # Light overlay — accent highlights
    "medium":  40,   # Medium — disabled states, image overlays
    "strong":  65,   # Strong — hero overlays, modal backgrounds
}
```

**Extended Palette Structure:**

```python
# Current: 10 flat colors
palette = {
    "primary": "#2563EB", "on-primary": "#FFFFFF", "secondary": "#64748B",
    "accent": "#F97316", "background": "#FFFFFF", "foreground": "#1E293B",
    "muted": "#F1F5F9", "muted-foreground": "#94A3B8", "border": "#E2E8F0",
    "destructive": "#EF4444",
}

# New: 10 base colors + 11 tint/shade per primary/secondary/accent + alpha variants
palette = {
    # Base colors (same as before, backward compatible)
    "primary": "#2563EB", "on-primary": "#FFFFFF", ...
    
    # Auto-generated tint/shade scales
    "primary-50":  "#EFF6FF",  # Ghost background
    "primary-100": "#DBEAFE",  # Subtle background
    "primary-200": "#BFDBFE",  # Light border
    "primary-300": "#93C5FD",  # Soft accent
    "primary-400": "#60A5FA",  # Medium accent
    "primary-500": "#2563EB",  # Base (= primary)
    "primary-600": "#1D4ED8",  # Dark hover
    "primary-700": "#1E40AF",  # Deeper active
    "primary-800": "#1E3A8A",  # Very dark heading
    "primary-900": "#172554",  # Near black
    "primary-950": "#0C1A3A",  # Dark mode bg
    
    # Same for secondary and accent
    "secondary-50": ..., "secondary-500": ..., "secondary-950": ...,
    "accent-50": ..., "accent-500": ..., "accent-950": ...,
    
    # Surface colors (for elevation)
    "surface-0": "#FFFFFF",   # Base background
    "surface-1": "#F8FAFC",   # Card resting (z-1)
    "surface-2": "#F1F5F9",   # Card raised (z-2)
    "surface-3": "#E2E8F0",   # Card overlay (z-3)
}
```

**Key Changes:**

1. `theme_composer.py`: After composing a palette, auto-generate tint/shade scales and surface colors
2. `precision_renderer.py`: Card backgrounds use `surface-1`/`surface-2` instead of hardcoded `"#F8FAFC"`/`"#0D152A"`
3. Badge backgrounds use `primary-100` (10% tint) instead of solid `primary`
4. Muted text uses `foreground-400` instead of flat `muted-foreground`
5. Hero overlay uses gradient from `primary-900` to `primary-950` instead of flat black
6. Bottom bar uses `surface-1` gradient instead of flat `muted`
7. **Replace `_lighten()`/`_darken()`**: After §1.3 is implemented, `precision_renderer.py`'s naive RGB `_lighten()`/`_darken()` methods (lines 302-318) must delegate to `color_system.py`'s OKLCH-based tint/shade functions. Two coexisting color derivation systems will produce inconsistent results.

**Files to Modify:**
- `src/ppt_pro_max/renderer/color_system.py` (NEW)
- `src/ppt_pro_max/renderer/theme_composer.py` (auto-generate scales)
- `src/ppt_pro_max/enterprise/precision_renderer.py` (use scale colors)
- `src/ppt_pro_max/renderer/ppt_renderer.py` (use scale colors)

**Estimated Effort:** 2-3 days

---

### 1.4 Gradient Overlay Gradient System — Replace Flat Black Overlays

**Problem:** Hero slides use a flat overlay at 72% opacity (the call site `add_dark_overlay(slide, 0.72)` at line 562; method default is 0.65). The overlay uses black on light themes and dark-blue on dark themes. This kills the image and looks like a template. Professional decks use gradient overlays that fade from transparent to dark, preserving the image at the top.

**Solution:** Replace `add_dark_overlay()` with gradient overlay that fades from top (transparent) to bottom (dark).

**Implementation:**

```python
def add_gradient_overlay(self, slide, opacity_bottom: float = 0.72, 
                          opacity_top: float = 0.0, 
                          color_role: str = "background") -> None:
    """Add a vertical gradient overlay over the entire slide.
    
    Top: transparent (opacity_top) — preserves image
    Bottom: dark (opacity_bottom) — ensures text contrast
    
    Uses GradientFill from visual_effects.py for clean OOXML generation.
    Requires: from ppt_pro_max.renderer.visual_effects import GradientFill, GradientStop
    """
    bg_hex = self._c(color_role, "#000000")
    
    # Create a full-slide rectangle
    ov = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                                Inches(SLIDE_WIDTH), Inches(SLIDE_HEIGHT))
    ov.line.fill.background()
    
    # Apply gradient fill using existing GradientFill class
    # Alpha values: OOXML units 0-100000 where 100000 = fully opaque, 0 = fully transparent
    top_alpha = int(opacity_top * 100000)    # 0% = 0, 100% = 100000
    mid_alpha = int(opacity_bottom * 0.4 * 100000)
    bot_alpha = int(opacity_bottom * 100000)
    
    grad = GradientFill(
        stops=[
            GradientStop(color=bg_hex, position=0, alpha=top_alpha),
            GradientStop(color=bg_hex, position=40000, alpha=mid_alpha),
            GradientStop(color=bg_hex, position=100000, alpha=bot_alpha),
        ],
        angle=5400000,  # 90 degrees = top to bottom
    )
    grad.apply(ov)
```

**Key Changes:**

1. Replace `add_dark_overlay(slide, 0.72)` (line 562) with `add_gradient_overlay(slide, opacity_bottom=0.72, opacity_top=0.0)` — same bottom darkness, but top is transparent
2. Hero text position moves to bottom 40% of slide (where overlay is dark)
3. For CTA slides, use a center-focused radial gradient instead of linear
4. For dark-mode hero slides, use `primary-950` instead of black for the overlay color
5. Add optional vignette: radial gradient from center (transparent) to edges (5% dark)

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (replace `add_dark_overlay`)
- `src/ppt_pro_max/renderer/ppt_renderer.py` (replace `_apply_dark_overlay`)

**ppt_renderer.py Replacement:**

The `ppt_renderer.py` version (`_apply_dark_overlay`, lines 143-163) uses pure black `RGBColor(0,0,0)` with alpha 55000 (55%). It must also be replaced with gradient overlay:

```python
def _apply_gradient_overlay(self, slide, opacity_bottom: float = 0.72, opacity_top: float = 0.0) -> None:
    """Replace flat black overlay with gradient overlay."""
    from ppt_pro_max.renderer.visual_effects import GradientFill, GradientStop
    ov = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(0), Inches(0), Inches(SLIDE_WIDTH), Inches(SLIDE_HEIGHT))
    ov.line.fill.background()
    grad = GradientFill(
        stops=[
            GradientStop(color="000000", position=0, alpha=int(opacity_top * 100000)),
            GradientStop(color="000000", position=40000, alpha=int(opacity_bottom * 0.4 * 100000)),
            GradientStop(color="000000", position=100000, alpha=int(opacity_bottom * 100000)),
        ],
        angle=5400000,
    )
    grad.apply(ov)
```

**Estimated Effort:** 1-2 days

---

### 1.5 Shadow Elevation System — 5-Level Depth

**Problem:** All shadows use the same parameters (blur=4pt, alpha=15-25%). No elevation hierarchy. Elements at different visual depths look flat.

**Solution:** Implement a 5-level shadow system with proper blur/offset/alpha ratios.

**Shadow Scale:**

| Level | Name | Offset-Y | Blur | Alpha | Color | Use |
|-------|------|----------|------|-------|-------|-----|
| 0 | Flat | 0pt | 0pt | 0% | — | Background elements, full-bleed |
| 1 | Resting | 1pt | 3pt | 10% | foreground | Cards on light bg, static containers |
| 2 | Raised | 2pt | 6pt | 15% | foreground | Active cards, floating panels |
| 3 | Overlay | 4pt | 12pt | 20% | foreground | Key feature cards, modals |
| 4 | Hero | 8pt | 24pt | 25% | foreground | Hero image frame, primary CTA |

**Critical Rule:** Shadow color = foreground color (black on light slides, white on dark slides), NEVER gray.

**Implementation:**

```python
# In visual_effects.py or new elevation.py

ELEVATION_SCALE = {
    0: {"blur_pt": 0,  "distance_pt": 0, "alpha_pct": 0},
    1: {"blur_pt": 3,  "distance_pt": 1, "alpha_pct": 10},
    2: {"blur_pt": 6,  "distance_pt": 2, "alpha_pct": 15},
    3: {"blur_pt": 12, "distance_pt": 4, "alpha_pct": 20},
    4: {"blur_pt": 24, "distance_pt": 8, "alpha_pct": 25},
}

def apply_elevation(shape, level: int, is_dark: bool = False, 
                     foreground_hex: str = "#000000", primary_hex: str = "#2563EB"):
    """Apply shadow at the given elevation level.
    
    On dark backgrounds, use glow instead of shadow.
    foreground_hex: hex color for shadow (black on light, white on dark).
    primary_hex: hex color for glow on dark mode.
    
    Note: apply_shadow() signature is (shape, blur_pt, distance_pt, direction_deg=90, color, alpha_pct)
    """
    spec = ELEVATION_SCALE.get(level, ELEVATION_SCALE[1])
    
    if level == 0:
        return  # No shadow
    
    if is_dark:
        # Glow effect for dark mode (primary color, subtle)
        # Minimum alpha_pct=15 to ensure visibility (5% is invisible)
        apply_glow(shape, 
                   radius_pt=spec["blur_pt"],
                   color=primary_hex,
                   alpha_pct=max(15, spec["alpha_pct"] // 2))
    else:
        # Shadow for light mode (foreground color = black)
        # CRITICAL: direction_deg comes BEFORE color in apply_shadow signature
        apply_shadow(shape,
                     blur_pt=spec["blur_pt"],
                     distance_pt=spec["distance_pt"],
                     direction_deg=90,
                     color=foreground_hex,
                     alpha_pct=spec["alpha_pct"])
```

**Key Changes:**

1. Cards: elevation level 1 (resting) instead of current blur=4/alpha=15
2. Featured/hero card: elevation level 3 (overlay) — the first card in a features row gets higher elevation
3. CTA button: elevation level 2 (raised)
4. Code block: elevation level 1 (resting)
5. Dark mode: use glow instead of shadow (primary color at 5-10% alpha, 8pt radius)
6. Remove the 1pt border on cards — use shadow + background tint instead (borders are a template tell)

**Files to Modify:**
- `src/ppt_pro_max/renderer/visual_effects.py` (add `apply_elevation()`)
- `src/ppt_pro_max/enterprise/precision_renderer.py` (use elevation levels)
- `src/ppt_pro_max/renderer/shape_factory.py` (use elevation levels)

**Estimated Effort:** 1-2 days

---

### 1.6 Remove the "AI Watermark" — Conditional Brand Strip

**Problem:** `apply_brand_background()` adds a 0.06" left accent strip and 0.25" bottom muted bar to EVERY content slide. This is the #1 visual tell of auto-generation.

**Solution:** Make the brand strip conditional and varied.

**Implementation:**

```python
def apply_brand_background(self, slide, prs, goal="content", page_index=0, total_pages=0):
    # Background fill (keep)
    bg_hex = self._c("background", "#FFFFFF")
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = self._rgb(bg_hex)
    
    # Brand strip: only on certain slides, with variation
    # Note: BrandSpec does not have an 'extra' field.
    # Read strip_style from brand.json via spacing dict.
    strip_style = "auto"
    if self._brand and self._brand.spacing:
        strip_style = self._brand.spacing.get("strip_style", "auto")
    
    if strip_style == "none":
        pass  # No strip at all — cleanest
    elif strip_style == "auto":
        # Only add strip on non-first, non-last content slides
        # Vary the strip: left bar, bottom bar, or corner accent
        variant = page_index % 3
        if variant == 0 and page_index > 0:
            # Left accent strip (current behavior, but only 1/3 of slides)
            accent_hex = self._c("accent", self._c("primary", "#2563EB"))
            self.add_rect(slide, 0, 0, 0.06, SLIDE_HEIGHT, fill_hex=accent_hex, gradient=True)
        elif variant == 1:
            # Bottom accent line (thin, 2pt, full width)
            accent_hex = self._c("accent", self._c("primary", "#2563EB"))
            self.add_rect(slide, 0, SLIDE_HEIGHT - 0.03, SLIDE_WIDTH, 0.03, fill_hex=accent_hex)
        # variant == 2: no strip at all — breathing room
    elif strip_style == "left":
        accent_hex = self._c("accent", self._c("primary", "#2563EB"))
        self.add_rect(slide, 0, 0, 0.06, SLIDE_HEIGHT, fill_hex=accent_hex, gradient=True)
    
    # Bottom muted bar: reduced height, gradient from lighter muted to muted
    # Current gradient (add_rect with gradient=True) goes lighter→darker same hue
    muted_hex = self._c("muted", "#F1F5F9")
    # Only add if there's no footer text already
    if not (self._brand and self._brand.footer and self._brand.footer.get("show_footer_text")):
        self.add_rect(slide, 0, SLIDE_HEIGHT - 0.15, SLIDE_WIDTH, 0.15, 
                      fill_hex=muted_hex, gradient=True)
```

**Key Changes:**

1. Left accent strip appears on only ~1/3 of content slides (not every slide)
2. Alternative: bottom accent line (thin, 2pt) on another 1/3
3. Remaining 1/3 of slides have no strip at all — pure whitespace
4. Bottom muted bar reduced from 0.25" to 0.15" (less dominant)
5. User can override via `brand.json` `strip_style` field
6. First slide (hook) and last slide (CTA) never get strips
7. **Caller update required**: `render_slide()` (line 375) must pass `page_index` and `total_pages` when calling `apply_brand_background()`. Pipeline must provide these values to the renderer.

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (refactor `apply_brand_background`)

**Estimated Effort:** 0.5-1 day

---

### 1.7 Image Color-Grading — Palette Harmony

**Problem:** Images are inserted as-is, with no color relationship to the deck palette. Stock photos often clash with the brand colors.

**Solution:** Apply a subtle primary-color overlay (8-12% opacity) to all images using Pillow before insertion.

**Implementation:**

```python
# In precision_renderer.py or new image_processor.py

from PIL import Image as PILImage

def grade_image_to_palette(image_path: str, primary_hex: str, alpha: float = 0.10) -> str:
    """Apply color overlay to harmonize image with palette.
    
    Returns path to graded image (cached).
    Preserves original format (JPEG stays JPEG) to avoid file size bloat.
    """
    cache_key = hashlib.md5(f"{image_path}:{primary_hex}:{alpha}".encode()).hexdigest()
    cache_dir = os.path.join(tempfile.gettempdir(), "ppt-graded-images")
    os.makedirs(cache_dir, exist_ok=True)
    
    # Detect original format to preserve it
    ext = os.path.splitext(image_path)[1].lower()
    if ext in (".jpg", ".jpeg"):
        cache_path = os.path.join(cache_dir, f"{cache_key}.jpg")
    else:
        cache_path = os.path.join(cache_dir, f"{cache_key}.png")
    
    if os.path.exists(cache_path):
        return cache_path
    
    img = PILImage.open(image_path).convert("RGB")
    
    # Parse primary color
    h = primary_hex.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    
    # Create solid color overlay
    overlay = PILImage.new("RGB", img.size, (r, g, b))
    
    # Blend: result = img * (1 - alpha) + overlay * alpha
    graded = PILImage.blend(img, overlay, alpha)
    
    # Save in original format to avoid file size bloat
    if ext in (".jpg", ".jpeg"):
        graded.save(cache_path, "JPEG", quality=90)
    else:
        graded.save(cache_path, "PNG")
    
    return cache_path
```

**Key Changes:**

1. Before `add_image()`, call `grade_image_to_palette()` with the deck's primary color at 10% alpha
2. Hero images get stronger grading (12%) since they're larger
3. Side images get lighter grading (8%) since they're smaller
4. Cache graded images to avoid re-processing
5. For dark-mode decks, use a darkening overlay instead of color overlay (blend toward `primary-950`)

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (add grading before `add_image`)
- `src/ppt_pro_max/renderer/ppt_renderer.py` (add grading before image insertion)
- `src/ppt_pro_max/renderer/image_processor.py` (NEW)

**Estimated Effort:** 1 day

---

### 1.8 Card Design Upgrade — Visual Hierarchy Within Cards

** Cards

**Problem:** All cards are identical: same height (4.5"), same accent bar (0.4" × 0.05"), same margins. No card has visual priority. Card body at 11pt is too small.

**Solution:** Add visual hierarchy within and between cards.

**Card Design Specification:**

```
┌─────────────────────────────────┐
│ ▎  ← 3pt primary accent bar     │  ← Top: 3pt (0.04") accent bar, left-aligned
│ ▎                               │
│ ▎  Card Title                   │  ← 20pt semibold, primary color
│ ▎                               │
│ ▎  Body text line one           │  ← 14pt regular, foreground color
│ ▎  Body text line two           │  ← Line height: 1.5x
│ ▎  Body text line three         │
│                                 │
│  ┌──────────┐                   │  ← Optional: icon/image area (bottom-right)
│  │   🎯     │                   │
│  └──────────┘                   │
└─────────────────────────────────┘
  Background: surface-1 (z-1 elevation)
  Border: none (shadow provides separation)
  Corner radius: 8pt (medium)
  Padding: 0.2" all sides
```

**Featured Card (first card in a row):**

```
┌─────────────────────────────────┐
│ ██████████████████████████████  │  ← Full-width primary gradient bar (0.15" tall)
│                                 │
│  Featured Title                 │  ← 22pt bold, primary color
│                                 │
│  Body text with more emphasis   │  ← 14pt regular, foreground
│  and slightly larger size       │
│                                 │
└─────────────────────────────────┘
  Background: surface-2 (z-2 elevation, higher shadow)
  Border: none
  Corner radius: 8pt
```

**Key Changes:**

1. Card title: 15pt → 20pt (subtitle level)
2. Card body: 11pt → 14pt (body level)
3. First card in a row gets "featured" treatment: full-width gradient bar, larger title, higher elevation
4. Accent bar: 0.4" × 0.05" → 0.04" × full-card-width (thin line, not a small rectangle)
5. Card border: remove 1pt border, rely on shadow + background tint
6. Card height: dynamic based on content (not always 4.5")
7. Card padding: 0.15" → 0.2" (more breathing room)
8. Card corner radius: use consistent 8pt from the radius scale

**Implementation Order Note:** This upgrade (§1.8) must be implemented BEFORE §2.6 (Layout Variant Consumption). §1.8 changes `add_card()` internal logic (featured treatment, dynamic height). §2.6 then adds variant-driven card_style selection on top. If done in reverse order, §2.6's variant logic would be overwritten by §1.8's refactoring.

**Signature Change:** `add_card()` must be extended from `(slide, x, y, w, h, title, body, accent_role="primary")` to `(slide, x, y, w, h, title, body, accent_role="primary", featured=False)`. When `featured=True`, render full-width gradient bar, larger title (22pt), higher elevation (level 2), and `surface-2` background.

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (refactor `add_card()`)
- `src/ppt_pro_max/renderer/shape_factory.py` (update card rendering)

**Estimated Effort:** 1-2 days

---

### 1.9 Dark Mode Luminance Detection — Replace Hardcoded List

**Problem:** `_is_dark()` uses a hardcoded list of 7 hex values. Any dark background not in the list is treated as light, causing white text on dark backgrounds.

**Solution:** Use luminance calculation (already exists in `theme_composer.py` but not used in `precision_renderer.py`).

**Implementation:**

```python
def _is_dark(self) -> bool:
    """Detect dark mode via luminance calculation, not hardcoded list."""
    if self._brand and self._brand.dark_mode:
        return True
    bg_hex = self._c("background", "#FFFFFF").lstrip("#")
    r, g, b = int(bg_hex[0:2], 16), int(bg_hex[2:4], 16), int(bg_hex[4:6], 16)
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return luminance < 0.5
```

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (replace `_is_dark`)

**Estimated Effort:** 0.5 day

---

### 1.10 Code Block Redesign — Syntax Highlighting + Language Badge

**Problem:** Code blocks are a rounded rect with monospace text. No syntax highlighting, no line numbers, no language indicator.

**Solution:** Add basic syntax highlighting (keyword coloring) and a language badge.

**Implementation:**

```python
def _render_code_on_slide(self, slide, code_data) -> None:
    code_text = code_data if isinstance(code_data, str) else code_data.get("source", "")
    language = code_data.get("language", "") if isinstance(code_data, dict) else ""
    
    # Language badge (top-left, pill-shaped)
    # Note: add_rounded_rect does not yet support corner_radius param — 
    # this will be added as part of upgrade 3.3. For now, use default radius.
    if language:
        badge_text = language.upper()
        badge_w = len(badge_text) * 0.12 + 0.2
        self.add_rounded_rect(slide, 1.0, 1.3, badge_w, 0.3, fill_role="primary")
        # Note: add_text does not yet support tracking param — 
        # this will be added as part of upgrade 1.2. For now, use bold.
        self.add_text(slide, badge_text, 1.1, 1.32, badge_w - 0.2, 0.26,
                      size=10, color_role="on-primary", bold=True)
    
    # Code background with elevation
    # Note: add_rounded_rect does not yet support elevation param —
    # this will be added as part of upgrade 1.5. For now, use shadow=True.
    code_bg = self._c("muted", "#1E293B")  # Code blocks always dark
    self.add_rounded_rect(slide, 0.9, 1.5, 11.533, 5.0, 
                          fill_hex=code_bg, shadow=True)
    
    # Syntax highlighting (basic: keywords, strings, comments)
    # Note: Per-line rendering with individual text boxes is expensive.
    # For v1, use add_multiline with a single styled block.
    # Full per-line syntax highlighting is a future enhancement.
    lines = code_text.split("\n")[:25]
    header = f"  {language}" if language else ""
    all_lines = []
    if header:
        all_lines.append(header)
    all_lines.extend(f"  {line}" for line in lines)
    self.add_multiline(slide, all_lines, 1.2, 1.7, 11, 4.5,
                       font="Consolas", size=11, color_role="foreground", spacing=4)
```

**Key Changes:**

1. Code background: always dark (#1E293B) regardless of slide theme (code is always dark). Note: current code uses `self._c("muted", "#0D152A" if self._is_dark() else "#F8FAFC")` which makes code blocks light on light themes — this should change to always-dark.
2. Language badge: pill-shaped, ALL CAPS, +0.08em tracking
3. Line numbers: muted, right-aligned, 2-digit
4. Basic syntax highlighting: keywords in accent, strings in soft primary, comments in muted
5. Line height: 0.22" per line (comfortable for 11pt monospace)
6. Elevation: level 1 shadow on code block

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (refactor `_render_code_on_slide`)

**Estimated Effort:** 1-2 days

---

## Tier 2: Significant Quality Improvement

---

### 2.1 CJK + Latin Font Pairing System

**Problem:** CJK fallback is hardcoded as "Microsoft YaHei" for all Latin fonts. Serif Latin fonts (Playfair Display, Georgia) paired with sans-serif CJK looks jarring.

**Solution:** Define explicit CJK companions for each Latin font pair.

**Font Pairing Map:**

```python
CJK_COMPANIONS = {
    # Sans-serif Latin → Sans-serif CJK
    "Inter":          {"heading": "Microsoft YaHei", "body": "Microsoft YaHei"},
    "Space Grotesk":  {"heading": "Microsoft YaHei", "body": "Microsoft YaHei"},
    "Poppins":        {"heading": "Source Han Sans SC", "body": "Microsoft YaHei"},
    "Montserrat":     {"heading": "Source Han Sans SC", "body": "Microsoft YaHei"},
    "Segoe UI":       {"heading": "Microsoft YaHei", "body": "Microsoft YaHei"},
    "Calibri":        {"heading": "Microsoft YaHei", "body": "Microsoft YaHei"},
    "Nunito":         {"heading": "Microsoft YaHei", "body": "Microsoft YaHei"},
    "Arial":          {"heading": "SimHei", "body": "SimHei"},
    
    # Serif Latin → Serif CJK
    "Playfair Display": {"heading": "Noto Serif CJK SC", "body": "Microsoft YaHei"},
    "Georgia":          {"heading": "SimSun", "body": "Microsoft YaHei"},
    "Lora":             {"heading": "Noto Serif CJK SC", "body": "Microsoft YaHei"},
    "Times New Roman":  {"heading": "SimSun", "body": "SimSun"},
    
    # Mono Latin → Mono CJK
    "Consolas": {"heading": "Microsoft YaHei", "body": "Microsoft YaHei"},
}
```

**Implementation:** When setting `run.font.name`, also set the East Asian typeface via XML for CJK characters. python-pptx's `run.font.name` only sets the Latin typeface (`<a:latin>`). CJK font must be set separately via `<a:ea>` (East Asian) and `<a:cs>` (Complex Script) elements on run properties.

```python
def _set_font_with_cjk(self, run, latin_font: str, cjk_font: str | None = None) -> None:
    """Set both Latin and CJK fonts on a run.
    
    python-pptx's run.font.name only sets a:latin typeface.
    CJK companion must be set via XML: a:ea and a:cs elements.
    """
    run.font.name = latin_font
    if cjk_font:
        rPr = run._r.get_or_add_rPr()
        # Set East Asian typeface
        ea = rPr.find(qn("a:ea"))
        if ea is None:
            ea = etree.SubElement(rPr, qn("a:ea"))
        ea.set("typeface", cjk_font)
        # Set Complex Script typeface
        cs = rPr.find(qn("a:cs"))
        if cs is None:
            cs = etree.SubElement(rPr, qn("a:cs"))
        cs.set("typeface", cjk_font)
```

**Integration in PrecisionRenderer:** Add `_set_font_with_cjk()` as a method. In `add_text()` and `add_multiline()`, after creating the run, call `self._set_font_with_cjk(run, font, cjk_companion)` where `cjk_companion` is looked up from `CJK_COMPANIONS` based on the Latin font name.

**Files to Modify:**
- `src/ppt_pro_max/renderer/theme_mapper.py` (replace hardcoded CJK mapping)
- `src/ppt_pro_max/enterprise/precision_renderer.py` (use CJK companions)
- `src/ppt_pro_max/renderer/shape_factory.py` (use CJK companions)

**Estimated Effort:** 1 day

---

### 2.2 Content-Adaptive Margins — Density-Aware Whitespace

**Problem:** All content slides use 0.9" left margin regardless of content density. Sparse slides look barren; dense slides look crammed.

**Solution:** Adjust margins based on content density.

**Implementation:**

```python
def compute_margins(self, page: dict, mode: str = "balanced") -> dict:
    """Compute margins based on content density and presentation mode."""
    bullets = page.get("bullets", [])
    cards = page.get("cards", [])
    has_image = bool(page.get("image"))
    
    # Count content elements
    content_count = len(bullets) + len(cards) * 3
    if has_image:
        content_count += 2
    
    if mode == "presenting" or content_count <= 3:
        # Sparse content → generous margins (keynote style)
        return {"left": 1.5, "right": 1.5, "top": 1.2, "bottom": 0.8}
    elif mode == "reading" or content_count >= 8:
        # Dense content → tight margins (consulting style)
        return {"left": 0.8, "right": 0.8, "top": 0.6, "bottom": 0.5}
    else:
        # Balanced
        return {"left": 0.9, "right": 0.9, "top": 0.6, "bottom": 0.5}
```

**Files to Modify:**
- `src/ppt_pro_max/renderer/layout_engine.py` (add margin computation)
- `src/ppt_pro_max/enterprise/precision_renderer.py` (use computed margins)

**Estimated Effort:** 1 day

---

### 2.3 Badge/Tag Styling System

**Problem:** Exercise badge is a solid-color rounded rect with 11pt text. Looks like a Bootstrap badge, not a designed element.

**Solution:** Implement a proper badge/tag system with ALL CAPS, tracking, and subtle background.

**Badge Specification:**

```
┌──────────────────┐
│  EXERCISE 5 MIN  │  ← ALL CAPS, +0.08em tracking, 10-11pt, semibold
└──────────────────┘
  bg: primary at 10% alpha (primary-100 tint)
  text: primary at 100%
  radius: 4pt (small badge) or pill (999pt)
  padding: 0.1" horizontal, 0.05" vertical
  border: none
  shadow: none (badges are flat)
```

**Implementation:**

```python
def add_badge(self, slide, text: str, x: float, y: float, 
              variant: str = "default", size: str = "sm") -> None:
    """Add a styled badge/tag element.
    
    variant: "default" (primary tint bg), "outline" (border only), "solid" (primary bg)
    size: "sm" (10pt), "md" (11pt), "lg" (12pt)
    """
    font_size = {"sm": 10, "md": 11, "lg": 12}[size]
    badge_text = text.upper()
    
    # Calculate badge width from text
    # CJK characters are ~2x wider than Latin at same font size
    cjk_count = sum(1 for ch in badge_text if '\u4e00' <= ch <= '\u9fff' or '\u3000' <= ch <= '\u303f')
    latin_count = len(badge_text) - cjk_count
    char_width = font_size * 0.006  # Latin inches per character at given pt
    cjk_width = font_size * 0.012   # CJK inches per character
    badge_w = latin_count * char_width + cjk_count * cjk_width + 0.3  # + padding
    badge_h = 0.35
    
    if variant == "default":
        # Tinted background
        # Note: corner_radius param will be added in upgrade 3.3
        self.add_rounded_rect(slide, x, y, badge_w, badge_h,
                              fill_hex=self._c("primary-100", self._tint("primary", 0.10)))
        # Note: tracking param will be added in upgrade 1.2
        self.add_text(slide, badge_text, x + 0.1, y + 0.02, badge_w - 0.2, badge_h - 0.04,
                      size=font_size, color_role="primary", bold=True)
    elif variant == "solid":
        self.add_rounded_rect(slide, x, y, badge_w, badge_h,
                              fill_role="primary")
        self.add_text(slide, badge_text, x + 0.1, y + 0.02, badge_w - 0.2, badge_h - 0.04,
                      size=font_size, color_role="on-primary", bold=True)
```

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (add `add_badge()`, refactor exercise rendering)

**Estimated Effort:** 1 day

---

### 2.4 Section Divider Slide Generation

**Problem:** No section dividers. Every slide looks the same. No visual rhythm or cognitive reset between topics.

**Solution:** Auto-generate section divider slides when content shifts topic.

**Section Divider Design:**

```
┌─────────────────────────────────────────────┐
│                                             │
│                                             │
│          01                                 │  ← Section number: 72pt, primary-200 (20% tint)
│                                             │
│    Section Title Here                       │  ← 36-44pt bold, foreground
│    ─────────────                            │  ← 2" gradient line, primary → transparent
│    Optional subtitle                        │  ← 18pt, muted-foreground
│                                             │
│                                             │
└─────────────────────────────────────────────┘
  Background: primary-50 (light tint) for both light and dark modes — NOT the standard background
  Dark mode note: light tint on dark slides creates strong pattern break (intentional contrast)
  No left accent strip, no bottom bar — this is a pattern break
```

**Implementation:**

```python
def render_section_divider(self, prs, section_number: int, 
                            section_title: str, section_subtitle: str = "") -> None:
    slide = self.add_slide(prs)
    
    # Different background color (pattern break)
    # Dark mode: use primary-50 (light tint) for strong contrast break
    # Light mode: use primary-50 (light tint) for subtle contrast break
    # In both modes, the key is a different background from regular content slides
    if self._is_dark():
        bg = self._c("primary-50", self._lighten(self._c("primary"), 200))
    else:
        bg = self._c("primary-50", self._lighten(self._c("primary"), 220))
    
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = self._rgb(bg)
    
    # Section number (oversized, low opacity)
    # Note: add_text expects float inches, not Inches() objects
    self.add_text(slide, f"{section_number:02d}", 
                  2.0, 1.5, 9.333, 2.0,
                  size=72, color_role="primary-200", bold=True)
    
    # Section title
    self.add_text(slide, section_title,
                  2.0, 3.5, 9.333, 1.0,
                  size=40, color_role="foreground", bold=True)
    
    # Gradient accent line
    # Note: add_gradient_line will be implemented in upgrade 3.4
    # For now, use add_rect with gradient=True as a placeholder
    self.add_rect(slide, 2.0, 4.6, 3.0, 0.03, fill_role="accent", gradient=True)
    
    # Subtitle
    if section_subtitle:
        self.add_text(slide, section_subtitle,
                      2.0, 4.8, 9.333, 0.5,
                      font=self._font_b(), size=18, color_role="muted-foreground")
```

**When to Generate Section Dividers:**

1. When `goal` changes from one type to another (e.g., hook → content → features → data)
2. When the title contains section-like keywords ("Part", "Chapter", "Section", "Phase")
3. When there are 4+ consecutive content slides of the same type (insert divider every 3-4 slides)
4. User can disable via `brand.json` `"section_dividers": false`

**Cross-Upgrade Dependency:** This upgrade (§2.4) requires §1.3 (Color System) to be implemented first. Section dividers use `primary-200` and `primary-950` color roles which are generated by the Color System's tint/shade scale. If §1.3 is not yet implemented, `_c("primary-200", ...)` will fall back to the default value, producing incorrect colors. **Implementation order: §1.3 → §2.4.**

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (add `render_section_divider`)
- `src/ppt_pro_max/enterprise/pipeline.py` (insert dividers between sections)

**Estimated Effort:** 1-2 days

---

### 2.5 Decoration System Completion — Implement All 10 Styles

**Problem:** 10 decoration styles are defined in `theme_composer.py` but only 3 have rendering implementations (accent-bar, left-accent, title-underline). The other 7 (neon-lines, gold-trim, minimal-dots, diamond-bullets, gradient-bar, circle-accent, sidebar-nav) produce no visual output.

**Solution:** Implement rendering for all 10 decoration styles.

**Decoration Rendering Map:**

| Style | Title Treatment | Body Treatment | Card Treatment | Special |
|-------|----------------|---------------|----------------|---------|
| accent-bar | Left accent bar + underline | Standard bullets | Top accent bar | Current default |
| neon-lines | Glowing accent line (primary glow) | Standard | Top glow bar | Dark mode only |
| gold-trim | Gold (#D4A853) underline | Standard | Gold top bar | Luxury mood |
| minimal-dots | No bar, dot before title | Circle bullets (•) | No bar, dot before title | Minimal mood |
| diamond-bullets | Underline | Diamond bullets (◆) | Diamond before title | Bold mood |
| gradient-bar | Gradient underline (primary→transparent) | Standard | Gradient top bar | Tech mood |
| circle-accent | Circle decoration next to title | Standard | Circle icon area | Creative mood |
| sidebar-nav | Left sidebar with section labels | Indented | Sidebar highlight | Navigation mood |
| no-decoration | Title only, no bar | Standard | No bar | Ultra-minimal |
| full-bleed-overlay | No bar (hero overlay handles it) | Standard | No bar | Hero slides |

**Implementation Strategy:**

Each decoration style maps to a `DecorationRenderer class method:

```python
class DecorationRenderer:
    """Renders decoration elements based on style."""
    
    def apply_title_decoration(self, slide, title_x, title_y, title_w, style, colors):
        method = getattr(self, f"_title_{style.replace('-', '_')}", self._title_accent_bar)
        method(slide, title_x, title_y, title_w, colors)
    
    def _title_accent_bar(self, slide, x, y, w, colors):
        # Current behavior: left bar + underline
        ...
    
    def _title_gradient_bar(self, slide, x, y, w, colors):
        # Gradient underline: primary → transparent
        apply_gradient_line(slide, x, y + 0.35, w * 0.3, 0.03,
                           colors["primary"], "transparent")
    
    def _title_neon_lines(self, slide, x, y, w, colors):
        # Glowing line with primary color glow effect
        line = add_rect(slide, x, y + 0.35, w * 0.3, 0.03, fill=colors["primary"])
        apply_glow(line, radius_pt=6, color=colors["primary"], alpha_pct=30)
    
    def _title_minimal_dots(self, slide, x, y, w, colors):
        # Small dot before title
        add_oval(slide, x - 0.15, y + 0.1, 0.08, 0.08, fill=colors["primary"])
    
    # ... etc for all 10 styles
```

**Files to Modify:**
- `src/ppt_pro_max/renderer/decoration_renderer.py` (NEW)
- `src/ppt_pro_max/enterprise/precision_renderer.py` (use DecorationRenderer)
- `src/ppt_pro_max/renderer/ppt_renderer.py` (use DecorationRenderer)

**Estimated Effort:** 2-3 days

---

### 2.6 Layout Variant Consumption — Connect Theme to Renderer

**Problem:** `theme_composer.py` defines 8 layout variants with margin/alignment/card_style properties, but `precision_renderer.py` ignores all of them. The `layout_variant` dict is composed but never consumed.

**Solution:** Pass the layout variant from theme composition through to the renderer and use it to adjust positioning.

**Implementation:**

```python
# In pipeline.py — extract and pass layout_variant

# Current call (line 391):
#   precision.render_slide(prs, design, component_lib=component_lib)

# New call:
layout_variant = design.get("layout_variant") if isinstance(design, dict) else None
precision.render_slide(prs, design, component_lib=component_lib, layout_variant=layout_variant)

# In precision_renderer.py — add layout_variant parameter

def render_slide(self, prs: Presentation, page: dict[str, Any], 
                 component_lib=None, layout_variant: dict | None = None) -> object:
    variant = layout_variant or {}
    
    # Extract variant properties
    margin_left = variant.get("content_margin_left", 0.9)
    margin_right = variant.get("content_margin_right", 0.9)
    title_align = variant.get("title_alignment", "left")
    card_style = variant.get("card_style", "rounded")
    
    # Use these values instead of hardcoded 0.9, 0.9, "left", "rounded"
    ...
```

**Key Changes:**

1. `pipeline.py` passes `layout_variant` from theme composition to `render_slide()`
2. `precision_renderer.py` uses variant margins instead of hardcoded 0.9"
3. `centered` variant: margins = 2.0", title centered
4. `sidebar-left` variant: content starts at 4.2", sidebar at 0-3.5"
5. `wide-cards` variant: card gap = 0.3" instead of 0.4"
6. `asymmetric` variant: staggered card heights

**Files to Modify:**
- `src/ppt_pro_max/enterprise/pipeline.py` (pass layout_variant)
- `src/ppt_pro_max/enterprise/precision_renderer.py` (consume layout_variant)

**Estimated Effort:** 1-2 days

---

## Tier 3: Polish That Separates "Pro" from "Good"

---

### 3.1 Subtle Noise Texture Overlay

**Problem:** Flat solid backgrounds have no texture. Professional decks use subtle noise for a "printed" feel.

**Solution:** Generate a noise texture tile with Pillow and overlay it at 2-3% opacity on hero and section divider slides.

**Implementation:**

```python
def _generate_noise_tile(self, size: int = 200, opacity: float = 0.02, deck_title: str = "") -> str:
    """Generate a noise texture tile for background overlay.
    
    Pure Python implementation — no numpy dependency.
    Uses random module for Gaussian-like noise generation.
    deck_title: per-deck seed source; same title = same noise, different titles = different noise.
    """
    import random
    cache_path = os.path.join(tempfile.gettempdir(), "ppt-noise", f"noise_{int(opacity*1000)}_{hashlib.md5(deck_title.encode()).hexdigest()[:8] if deck_title else 'default'}.png")
    if os.path.exists(cache_path):
        return cache_path
    
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    
    # Generate noise using pure Python (Box-Muller transform for Gaussian)
    pixels = []
    # Per-deck deterministic seed: varies across decks, consistent within a deck
    # Falls back to 42 if no deck title available
    seed_val = int(hashlib.md5(deck_title.encode()).hexdigest()[:8], 16) if deck_title else 42
    random.seed(seed_val)
    for y in range(size):
        row = []
        for x in range(size):
            # Box-Muller transform for approximate Gaussian noise
            u1 = random.random() or 0.001
            u2 = random.random()
            z = (-2.0 * __import__('math').log(u1)) ** 0.5 * __import__('math').cos(2.0 * __import__('math').pi * u2)
            val = int(max(0, min(255, 128 + z * 20)))
            alpha = int(opacity * 255)
            row.append((val, val, val, alpha))
        pixels.append(row)
    
    # Build PNG using Pillow
    img = PILImage.new("RGBA", (size, size))
    for y, row in enumerate(pixels):
        for x, pixel in enumerate(row):
            img.putpixel((x, y), pixel)
    
    img.save(cache_path, "PNG")
    return cache_path
```

**When to Apply:**
- Hero slides (hook, cta)
- Section dividers
- NOT on data/content/code slides (reduces readability)

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (add noise overlay to hero/divider slides)

**Estimated Effort:** 0.5 day

---

### 3.2 Progress Bar Indicator

**Problem:** No visual indication of deck progress. Audience loses sense of where they are.

**Solution:** Add a thin progress bar at the bottom of each content slide.

**Implementation:**

```python
def add_progress_bar(self, slide, current: int, total: int):
    """Add a thin progress bar at the bottom of the slide.
    
    Placed at the very bottom edge. When present, the bottom muted bar
    (from §1.6) is reduced to avoid overlap — the progress bar replaces
    the bottom portion of the muted bar.
    """
    bar_y = SLIDE_HEIGHT - 0.04
    bar_h = 0.03
    
    # Background track
    self.add_rect(slide, 0, bar_y, SLIDE_WIDTH, bar_h, 
                  fill_hex=self._c("border", "#E2E8F0"))
    
    # Filled portion
    fill_w = SLIDE_WIDTH * (current / total)
    self.add_rect(slide, 0, bar_y, fill_w, bar_h,
                  fill_hex=self._c("primary", "#2563EB"))
```

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (add progress bar to content slides)
- `src/ppt_pro_max/enterprise/pipeline.py` (pass slide index/total)

**Estimated Effort:** 0.5 day

---

### 3.3 Rounded Corner Consistency Enforcement

**Problem:** Corner radius varies inconsistently across elements. Some cards use 4pt, others use 8pt, badges use different values.

**Solution:** Define a consistent corner radius scale and enforce it.

**Radius Scale:**

| Element Type | Radius (pt) | When |
|-------------|-------------|------|
| Small (badges, tags, pills) | 4pt | Badges, category labels |
| Medium (cards, buttons, images) | 8pt | Feature cards, CTA buttons |
| Large (panels, code blocks) | 12pt | Full-width panels, code backgrounds |
| Pill (avatars, full-round badges) | 50% of height | Tag pills, avatars |

**Implementation:**

```python
CORNER_RADIUS_SCALE = {
    "sm": 4,    # 4pt — badges, tags
    "md": 8,    # 8pt — cards, buttons, images
    "lg": 12,   # 12pt — panels, code blocks
    "pill": 50, # 50% — pill shapes
}

# In add_rounded_rect():
def add_rounded_rect(self, slide, x, y, w, h, corner_radius="md", ...):
    if isinstance(corner_radius, str):
        radius_pt = CORNER_RADIUS_SCALE.get(corner_radius, 8)
    else:
        radius_pt = corner_radius
    # Convert pt to EMU for OOXML
    # <a:prstGeom prst="roundRect"> <a:avLst> <a:gd name="adj" fmla="val 10000"/>
    # adj value = radius / (min(w, h) / 2) * 100000
    ...
```

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (use radius scale)
- `src/ppt_pro_max/renderer/shape_factory.py` (use radius scale)

**Estimated Effort:** 0.5 day

---

### 3.4 Gradient Line Accents for Title Underlines

**Problem:** Title underlines are solid-color thin rectangles. Professional decks use gradient lines that fade from primary to transparent.

**Solution:** Replace solid accent bars with gradient lines.

**Implementation:**

```python
def add_gradient_line(self, slide, x, y, width, height, 
                       from_color: str, to_color: str = "transparent"):
    """Add a gradient line that fades from solid to transparent."""
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   Inches(x), Inches(y), Inches(width), Inches(height))
    rect.line.fill.background()
    
    # Use GradientFill class directly (supports alpha per stop)
    # apply_gradient(shape, color1, color2) does NOT support alpha
    grad = GradientFill(
        stops=[
            GradientStop(color=from_color, position=0),
            GradientStop(color=from_color if to_color == "transparent" else to_color, 
                         position=100000, alpha=0),
        ],
        angle=0,  # Left to right
    )
    grad.apply(rect)
```

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (replace `add_rect` for title underlines)

**Estimated Effort:** 0.5 day

---

### 3.5 Image Masking — Rounded Rectangle with Padding

**Problem:** Side images are placed as-is with no frame or mask. When image background differs from slide background, it looks jarring.

**Solution:** Add optional rounded-rectangle mask with padding for side images.

**Implementation:**

```python
def add_masked_image(self, slide, image_path, x, y, w, h, 
                      padding: float = 0.1, corner_radius: str = "md",
                      elevation: int = 1) -> None:
    """Add image with rounded-rectangle mask and padding.
    
    1. Draw rounded rect background (white or surface-1)
    2. Crop image to fit inside with padding
    3. Place image inside the rounded rect
    4. Apply shadow for elevation
    """
    # Background frame
    # Note: corner_radius and elevation params will be added in upgrades 3.3 and 1.5
    frame_bg = "#FFFFFF" if not self._is_dark() else self._c("surface-1", "#1E293B")
    self.add_rounded_rect(slide, x, y, w, h, fill_hex=frame_bg, shadow=True)
    
    # Image inside with padding
    img_x = x + padding
    img_y = y + padding
    img_w = w - 2 * padding
    img_h = h - 2 * padding
    self.add_image(slide, image_path, img_x, img_y, img_w, img_h)
```

**When to Use:**
- Side images on content slides (not full-bleed hero images)
- Product screenshots
- Team photos
- NOT for hero/CTA full-bleed images

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (add `add_masked_image`, use for side images)

**Estimated Effort:** 1 day

---

### 3.6 Two-Column Bullet Layout for 6+ Items

**Problem:** 8+ bullets in a single column at 12pt creates a dense text wall. Professional decks split into two columns.

**Solution:** When bullet count >= 6, split into two columns with a thin vertical separator.

**Implementation:**

```python
def _render_bullets(self, slide, bullets, x, y, w, h, layout_variant=None, layout=None):
    """Render bullets, using LayoutEngine's column decision if available.
    
    LayoutEngine (§1.1) decides whether to use 1 or 2 columns via
    layout.bullet_columns. This method only executes that decision.
    If no layout is provided, falls back to bullet count heuristic.
    """
    num_cols = 1
    if layout and hasattr(layout, 'bullet_columns'):
        num_cols = layout.bullet_columns
    elif len(bullets) >= 6:
        num_cols = 2
    
    if num_cols >= 2:
        # Two-column layout
        col_w = (w - 0.3) / 2  # 0.3" gap between columns
        mid = (len(bullets) + 1) // 2
        
        # Left column
        left_lines = [f"•  {b}" for b in bullets[:mid]]
        self.add_multiline(slide, left_lines, x, y, col_w, h,
                          size=14, color_role="foreground", spacing=8)
        
        # Right column
        right_lines = [f"•  {b}" for b in bullets[mid:]]
        self.add_multiline(slide, right_lines, x + col_w + 0.3, y, col_w, h,
                          size=14, color_role="foreground", spacing=8)
        
        # Thin vertical separator
        self.add_rect(slide, x + col_w + 0.15, y + 0.1, 0.01, h - 0.2,
                      fill_hex=self._c("border", "#E2E8F0"))
    else:
        # Single column
        lines = [f"•  {b}" for b in bullets[:8]]
        self.add_multiline(slide, lines, x, y, w, h,
                          size=14, color_role="foreground", spacing=8)
```

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (refactor bullet rendering)

**Estimated Effort:** 0.5 day

---

### 3.7 Hero Slide Layout Variety — 4 Composition Patterns

**Problem:** All hero/CTA slides use the same centered-text-over-image pattern. Professional decks vary hero compositions.

**Solution:** Implement 4 hero composition patterns and select based on content and mood.

**Hero Composition Patterns:**

| Pattern | Layout | When |
|---------|--------|------|
| **Full-bleed centered** | Image full-slide, text centered vertically | Default, most common |
| **Split-left** | Image right 60%, text left 40% on solid bg | When subtitle is long |
| **Bottom-fade** | Image top 60%, gradient fade, text bottom 40% | When image is the hero |
| **Asymmetric** | Image left 70% offset, text right 30% at power-point, gradient overlay on right | Creative/modern moods |

**Asymmetric Layout Specification:**
- Image: (0, 0, 9.333, 7.5) — full-height, 70% width, left-aligned
- Gradient overlay: right 40% of slide, from transparent to background color
- Title: (8.0, 2.0, 4.8, 1.5) — right-aligned, 36pt bold
- Subtitle: (8.0, 3.6, 4.8, 0.8) — right-aligned, 18pt

**Implementation:**

```python
def _select_hero_pattern(self, page: dict, mood: str) -> str:
    subtitle = page.get("subtitle", "")
    has_image = bool(page.get("image"))
    
    if not has_image:
        return "gradient"  # No image = gradient background
    
    if len(subtitle) > 60:
        return "split-left"  # Long subtitle needs more text space
    
    if mood in ("creative", "bold", "vibrant", "startup"):
        return "asymmetric"
    
    return "bottom-fade"  # Default: image with bottom gradient
```

**Files to Modify:**
- `src/ppt_pro_max/enterprise/precision_renderer.py` (add hero rendering with pattern selection)

**Estimated Effort:** 2 days

---

## Implementation Roadmap

### Phase A: Tier 1 (Weeks 1-3) — "Stop Looking Cheap"

| Week | Upgrades | Deliverable |
|------|----------|-------------|
| 1 | 1.1 Layout Engine + 1.6 Remove AI Watermark + 1.9 Dark Mode Fix | Dynamic positioning, no more "AI-generated" tell |
| 2 | 1.2 Typography Scale + 1.5 Shadow Elevation | Visual hierarchy, depth system |
| 3 | 1.3 Color Depth + 1.4 Gradient Overlays + 1.7 Image Grading | Color sophistication, image harmony |
| 3.5 | 1.8 Card Upgrade + 1.10 Code Block | Component quality |

**Phase A Success Metric:** Side-by-side comparison of same-content PPT before/after should show:
- No two slides look identical in structure
- Text is readable at 3m projection distance (min 14pt body)
- Images feel integrated with the color palette
- No flat black overlays visible

### Phase B: Tier 2 (Weeks 4-5) — "Look Professional"

| Week | Upgrades | Deliverable |
|------|----------|-------------|
| 4 | 2.1 CJK Font Pairing + 2.2 Adaptive Margins + 2.3 Badge System | CJK quality, density awareness |
| 4.5 | 2.4 Section Dividers + 2.5 Decoration Completion | Visual rhythm, all styles work |
| 5 | 2.6 Layout Variant Consumption | Theme system fully connected |

**Phase B Success Metric:**
- CJK text renders with proper companion fonts
- Section dividers create visual rhythm
- All 10 decoration styles produce visible output
- Layout variants actually change the rendered output

### Phase C: Tier 3 (Week 6) — "Look Designer-Grade"

| Week | Upgrades | Deliverable |
|------|----------|-------------|
| 6 | 3.1 Noise Texture + 3.2 Progress Bar + 3.3 Corner Radius + 3.4 Gradient Lines + 3.5 Image Masking + 3.6 Two-Column + 3.7 Hero Variety | Polish layer |

**Phase C Success Metric:**
- Hero slides have 4 distinct composition patterns
- Noise texture visible on hero/divider slides
- Progress bar on all content slides
- Consistent corner radius across all elements

### Phase D: Integration Verification (Week 7) — "No Regressions"

After all 28 upgrades are implemented, run a full integration verification:

| Check | Method | Pass Criteria |
|-------|--------|---------------|
| Existing tests still pass | `python -m pytest tests/ -q` | 0 failures |
| E1-E5 scenarios regenerate | Run E1-E5 with upgraded renderer | Same or more slides/shapes, no visual regressions |
| Cross-upgrade compatibility | Generate PPT with all upgrade features active | No overlapping elements, no color fallbacks, no missing shapes |
| Performance | Time E1 generation | < 2x baseline (caching should help) |
| Visual quality | Manual review by 2+ people | "Designer-grade" rating on 5+ of 7 dimensions |

**Critical cross-upgrade dependencies (must implement in this order):**

```
§1.3 Color System ──────→ §2.4 Section Dividers (needs primary-200, primary-950)
§1.2 Typography ────────→ §2.3 Badges (needs letter-spacing)
§1.5 Elevation ─────────→ §1.8 Cards (needs elevation levels) ──→ §2.6 Layout Variants
§1.1 Layout Engine ────→ §2.6 Layout Variants (needs computed positions)
§3.3 Corner Radius ────→ §1.10 Code Block (needs radius param)
§1.3 Color System ──────→ §1.4 Gradient Overlay (needs primary-950 for dark overlay)
```

---

## New Files to Create

| File | Purpose | Lines (est.) |
|------|---------|-------------|
| `src/ppt_pro_max/renderer/layout_engine.py` | Content-adaptive layout calculator | ~200 |
| `src/ppt_pro_max/renderer/typography.py` | TypeScale, line-height, letter-spacing | ~120 |
| `src/ppt_pro_max/renderer/color_system.py` | Tint/shade generation, alpha levels, surface colors | ~180 |
| `src/ppt_pro_max/renderer/image_processor.py` | Color-grading, noise texture, image masking | ~150 |
| `src/ppt_pro_max/renderer/decoration_renderer.py` | All 10 decoration style implementations | ~300 |
| `src/ppt_pro_max/renderer/elevation.py` | Shadow scale, glow-for-dark-mode | ~80 |

**Total new code:** ~1,030 lines

## Existing Files to Modify

| File | Changes | Impact |
|------|---------|--------|
| `precision_renderer.py` | Refactor `render_slide()`, `add_card()`, `_render_code_on_slide()`, `_render_exercise_on_slide()`, `apply_brand_background()`, `apply_hero_overlay()`, `add_dark_overlay()` | Major |
| `ppt_renderer.py` | Replace hardcoded sizes, add image grading | Medium |
| `theme_composer.py` | Auto-generate tint/shade scales after palette composition | Medium |
| `visual_effects.py` | Add `apply_letter_spacing()`, `apply_elevation()`, `add_gradient_line()` | Medium |
| `shape_factory.py` | Use TypeScale, CJK companions, corner radius scale | Medium |
| `theme_mapper.py` | Replace hardcoded CJK mapping with CJK_COMPANIONS dict | Small |
| `brand_spec.py` | Add `strip_style` to spacing dict support | Small |
| `pipeline.py` | Pass layout_variant, insert section dividers, pass slide index | Medium |

---

## Testing Strategy

### Visual Regression Tests

For each Tier 1 upgrade, create a test that:
1. Generates a PPT with specific content (3 bullets, 8 bullets, hero, cards, code, exercise)
2. Verifies the new design properties (font sizes, positions, colors, shadows)
3. Compares output file size (should increase slightly due to gradient XML)

### Quality Gate Tests

```python
# test_design_quality.py

def test_no_hardcoded_positions():
    """Verify render_slide uses LayoutEngine, not hardcoded coordinates."""
    # Mock LayoutEngine, verify it's called for each goal type
    
def test_typographic_scale_enforced():
    """Verify all font sizes come from TypeScale, not magic numbers."""
    # Generate PPT, check all run.font.size values are in the scale
    
def test_color_scale_generated():
    """Verify each palette has tint/shade variants."""
    # Compose theme, check palette has primary-50 through primary-950
    
def test_no_flat_black_overlay():
    """Verify hero slides use gradient overlay, not flat black."""
    # Generate hero slide, check overlay has gradient fill XML
    
def test_shadow_elevation_levels():
    """Verify cards use elevation level 1, featured cards use level 2+."""
    # Generate features slide, check shadow parameters on card shapes
    
def test_no_ai_watermark_strip():
    """Verify left accent strip doesn't appear on every content slide."""
    # Generate 5 content slides, verify strip appears on <= 2
    
def test_image_color_grading():
    """Verify images are color-graded before insertion."""
    # Generate slide with image, verify graded image path is used
    
def test_cjk_font_pairing():
    """Verify CJK text uses companion font, not default Microsoft YaHei."""
    # Generate slide with CJK text, check run.font.name for CJK companion
```

### End-to-End Evaluation

After each phase, run the existing E1-E5 evaluation scenarios and compare:
- Slide count (should increase with section dividers)
- Shape count (should increase with badges, progress bars)
- Min font size (should be >= 11pt, ideally >= 14pt for body)
- Visual quality (manual review by 2+ people)

---

## Risk Mitigation

| Risk | Mitigation |
|------|-----------|
| Breaking existing 638 tests | All changes are additive; new modules don't affect existing code until wired in. Wire in one upgrade at a time with feature flags. |
| Performance regression (color grading, noise generation) | All image processing is cached (MD5-based). First run is slower, subsequent runs are instant. |
| OKLCH dependency for color scale generation | Implement a pure-Python OKLCH converter (no external dependency). Fallback to HSL if OKLCH fails. |
| Layout engine produces overlapping elements | Add collision detection in LayoutEngine: verify no two shapes overlap. Add QA check in render_slide(). |
| Gradient overlay makes text unreadable | Always use bottom-fade (dark at bottom where text sits). Test with 10+ different hero images. |
| Shadow system looks different in PowerPoint vs LibreOffice | Test in both. Use only OOXML-standard shadow attributes. Avoid proprietary effects. |

---

## Appendix A: Current vs Target Design Comparison

| Dimension | Current | Target | Improvement |
|-----------|---------|--------|-------------|
| Typographic levels | 2 (title 28pt, body 12pt) | 5 (display 52pt → caption 11pt) | 4x scale contrast |
| Color variants per palette | 10 flat colors | 10 base + 33 tint/shade + 5 surface | 4.8x more colors |
| Shadow levels | 1 (blur=4, alpha=15-25%) | 5 (flat → hero) | 5x depth range |
| Hero overlay | Flat overlay 72% | Gradient 0%→72% | Image preserved |
| Layout positions | Hardcoded per goal | Content-adaptive | Dynamic |
| Card design | Identical, 11pt body | Featured + standard, 14pt body | Hierarchy |
| Decoration styles rendered | 3 of 10 | 10 of 10 | Complete |
| CJK font pairing | 1 fallback for all | 12 specific pairings | Proper pairing |
| Section dividers | None | Auto-generated | Visual rhythm |
| Image treatment | Raw insertion | Color-graded + masked | Palette harmony |
| Badge styling | Solid rect, 11pt | Tinted bg, ALL CAPS, tracked | Professional |
| Code blocks | Plain monospace | Syntax highlighted + badge | Readable |
| Corner radius | Inconsistent | 3-level scale (4/8/12pt) | Consistent |
| Progress indicator | None | Thin bar at bottom | Navigation |
| Noise texture | None | 2% overlay on hero/divider | Printed feel |

---

## Appendix B: Reference Design Standards

| Standard | Key Takeaway | Applied In |
|----------|-------------|-----------|
| McKinsey/BCG presentation standards | 4-column grid, 2-color charts, 24pt titles, left-aligned body | Layout engine, chart colors, typography |
| Apple Keynote design | 4x scale contrast, gradient overlays, generous whitespace | Hero patterns, typographic scale |
| Duarte slide:ology | Bottom-fade overlays, color-grading images, visual hierarchy | Image treatment, overlay system |
| Material Design 3 | 5-level elevation, OKLCH color system, dynamic type scale | Shadow system, color system, typography |
| Apple HIG | Consistent corner radius, glow on dark, SF-style tracking | Corner radius, dark mode, letter-spacing |
| NN/g visual design principles | 3+1 color rule, Z-pattern reading, rule of thirds | Color hierarchy, layout composition |

---

## Appendix C: Errata — Errors Found and Fixed During Self-Audit

The following errors were identified during a code-level audit against the actual codebase and OOXML specification, and have been corrected in this document:

| # | Section | Error | Fix |
|---|---------|-------|-----|
| 1 | 1.2 | Letter-spacing `a:spc` val described as "hundredths of a percentage point" — WRONG. OOXML spec says hundredths of a **point**, which is font-size-dependent. | Changed TypeScale tracking from int (hundredths-of-pct) to float (EM ratio). `apply_letter_spacing()` now takes `(run, tracking_em, font_size_pt)` and computes `val = tracking_em * font_size_pt * 100`. |
| 2 | 1.2 | `TypeScale.for_density(10)` produced body=10pt, below 11pt minimum | Added `max()` floor constraints: body >= 11pt, caption >= 9pt, display >= 36pt |
| 3 | 1.2 | `for_mode("reading")` caption=10pt, below 11pt minimum | Changed to caption=11pt |
| 4 | 1.3 | `alpha_color()` was a no-op function that just returned the hex string | Replaced with wrapper around existing `set_solid_fill_with_alpha()` from `visual_effects.py` |
| 5 | 1.3 | `ALPHA_LEVELS` used float 0.0-1.0, incompatible with `set_solid_fill_with_alpha()` which expects int 0-100 | Changed to int percent values (4, 8, 15, 40, 65) |
| 6 | 1.3 | `hex_to_oklch()` and `oklch_to_hex()` called but never defined; `oklch` library not installed | Implemented full pure-Python OKLCH conversion functions inline |
| 7 | 1.4 | Gradient overlay code manually built XML instead of using existing `GradientFill` class | Rewrote to use `GradientFill` + `GradientStop` from `visual_effects.py` |
| 8 | 1.5 | `apply_elevation()` had duplicate code block (copy-paste error) and passed `color="foreground"` (role string) to `apply_glow()` which expects hex color | Fixed: single code path, accepts `foreground_hex` and `primary_hex` as explicit hex parameters |
| 9 | 1.6 | `self._brand.extra.get("strip_style")` — `BrandSpec` has no `extra` field | Changed to `self._brand.spacing.get("strip_style", "auto")` using existing `spacing` dict |
| 10 | 1.6 | Bottom bar description said "subtle gradient fade" implying opaque→transparent, but `add_rect(gradient=True)` actually does lighter→darker same hue | Clarified description to match actual behavior |
| 11 | 1.7 | `from PIL import ImageChops` imported but never used | Removed unused import |
| 12 | 1.7 | All graded images saved as PNG, causing file size bloat for JPEG photos | Added format detection: JPEG photos stay JPEG (quality=90), PNG stays PNG |
| 13 | 1.8 | Accent bar height 0.03" (2.16pt) too thin to see on screen | Changed to 0.04" (~3pt) for visibility |
| 14 | 1.10 | `add_rounded_rect()` called with `corner_radius` and `elevation` params that don't exist | Added notes that these params will be added in upgrades 3.3 and 1.5; used existing `shadow=True` |
| 15 | 1.10 | `add_text()` called with `tracking=800` param that doesn't exist | Added note that tracking will be added in upgrade 1.2; removed for now |
| 16 | 1.10 | Per-line syntax highlighting creates 25+ individual text boxes (expensive, fragile) | Simplified to `add_multiline()` with language badge; full per-line highlighting deferred to future |
| 17 | 1.10 | Code segment width=10 inches for individual text runs — way too wide | Removed per-line rendering approach entirely |
| 18 | 2.3 | `char_width = font_size * 0 * 0.008` — multiplication by zero gives 0 | Fixed to `font_size * 0.006` |
| 19 | 2.3 | `add_rounded_rect()` called with `corner_radius=4` param that doesn't exist | Added note; will use default radius until upgrade 3.3 |
| 20 | 2.4 | `add_text()` called with `Inches(2.0)` but function expects float inches (internally calls `Inches(x)`) | Changed all `Inches(N)` to bare `N` floats |
| 21 | 2.4 | `self.add_gradient_line()` called but method doesn't exist yet | Replaced with `self.add_rect(..., gradient=True)` placeholder until upgrade 3.4 |
| 22 | 3.4 | `apply_gradient()` called with list of `GradientStop` objects, but function signature is `(shape, color1, color2)` | Rewrote to use `GradientFill` class directly, which supports alpha per stop |
| 23 | 3.5 | `add_rounded_rect()` called with `corner_radius` and `elevation` params that don't exist | Added note; used `shadow=True` as placeholder |
| 24 | 3.6 | Bullet font size 12pt contradicts upgrade 1.2 which raises body to 14pt | Changed to 14pt to match new typographic scale |

### Second Audit Errata (Design Logic Issues)

| # | Section | Error | Fix |
|---|---------|-------|-----|
| 25 | 1.1 | LayoutEngine produces `ContentLayout` but no integration pattern shown for `render_slide()` | Added explicit integration pattern: LayoutEngine as instance attribute in `__init__()`, called inside `render_slide()` |
| 26 | 1.1 | `_estimate_lines()` referenced but never implemented | Added full implementation with CJK/Latin character width differentiation |
| 27 | 1.3 | `alpha_color()` was a no-op `pass` statement — function body did nothing | Changed to actual wrapper that calls `set_solid_fill_with_alpha(shape, hex_color, opacity_pct)` with `shape` as first parameter |
| 28 | 1.5 | `apply_shadow()` called with wrong parameter order — `alpha_pct` before `color`, but actual signature is `(shape, blur_pt, distance_pt, direction_deg=90, color, alpha_pct)` | Fixed: added `direction_deg=90` between `distance_pt` and `color` |
| 29 | 1.6 | `apply_brand_background()` uses `page_index` parameter but current signature is `(self, slide, prs, goal="content")` — no page_index/total_pages | Updated signature to `(self, slide, prs, goal="content", page_index=0, total_pages=0)` |
| 30 | 1.8 | Card upgrade (§1.8) and layout variant (§2.6) both modify `add_card()` with no coordination | Added implementation order note: §1.8 must be before §2.6 |
| 31 | 2.3 | Badge width estimation `font_size * 0.006` assumes Latin-only; CJK characters are ~2x wider | Added CJK character counting and separate width calculation |
| 32 | 3.1 | Noise texture uses `import numpy as np` but numpy is not a project dependency | Replaced with pure Python implementation using Box-Muller transform + Pillow |
| 33 | 2.4 | Section dividers use `primary-200` and `primary-950` color roles that don't exist until §1.3 Color System is implemented | Added cross-upgrade dependency warning: §1.3 must be implemented before §2.4 |
| 34 | Roadmap | No integration verification step after all upgrades | Added Phase D: Integration Verification with cross-upgrade dependency graph |

### Third Audit Errata (Code Verification Against Actual Codebase)

The following errors were identified by verifying every code claim against the actual source files. Line numbers reference `src/ppt_pro_max/enterprise/precision_renderer.py` unless noted.

| # | Section | Error | Fix |
|---|---------|-------|-----|
| 35 | 1.4 | Problem statement said "flat black overlay at 72% opacity" — partially inaccurate. The call site `add_dark_overlay(slide, 0.72)` at line 562 does use 72%, but the method default is 0.65. Also, PrecisionRenderer uses `#060B18` (not black) in dark mode; only `ppt_renderer.py` uses pure black. | Corrected: "flat overlay at 72% opacity (call site line 562; method default is 0.65). Uses black on light, dark-blue on dark." Gradient overlay default changed to `opacity_bottom=0.72` to match current call site behavior at bottom while adding transparency at top. |
| 36 | 1.5 | Dark mode glow `alpha_pct=spec["alpha_pct"] // 2` produces 5% for level 1 — effectively invisible (5% = 5000 OOXML units). | Changed to `max(15, spec["alpha_pct"] // 2)` to ensure minimum visible glow. Updated level 1 glow to alpha_pct=15, level 2 to 15, level 3 to 20, level 4 to 25. |
| 37 | 1.3 | OKLCH `_hex_to_oklch()` has negative cube root bug: `l_ ** (1/3) if l_ > 0 else l_` passes negative LMS values through unchanged, producing incorrect OKLab coordinates. Same bug in `_oklch_to_hex()`. | Fixed: `l_ ** (1/3) if l_ >= 0 else -((-l_) ** (1/3))` for all three LMS channels in both functions. |
| 38 | 1.3 | `_lighten()`/`_darken()` (precision_renderer.py:302-318) use naive RGB arithmetic (`r ± amount`). Plan introduces OKLCH but never replaces these methods — two systems will coexist and produce inconsistent color derivations. | Added explicit instruction: after §1.3 is implemented, `_lighten()`/`_darken()` must delegate to `color_system.py`'s OKLCH-based tint/shade functions instead of naive RGB arithmetic. |
| 39 | 1.6 | `apply_brand_background()` signature changed to add `page_index`/`total_pages`, but `render_slide()` (line 375) calls it as `self.apply_brand_background(slide, prs, goal=goal)` without these new params. | Added note: `render_slide()` must be updated to pass `page_index` and `total_pages` when calling `apply_brand_background()`. Pipeline must provide these values. |
| 40 | 1.8 | `add_card()` signature (line 569) is `(slide, x, y, w, h, title, body, accent_role="primary")` — no `featured` parameter for the new "featured card" treatment. | Added: `add_card()` signature must be extended with `featured: bool = False` parameter. When `featured=True`, render full-width gradient bar, larger title, higher elevation. |
| 41 | 2.1 | Plan says "specifying both Latin and CJK names in the font stack" but python-pptx does NOT support font stacks. `run.font.name` only sets the Latin typeface. CJK font must be set via XML: `<a:ea typeface="...">` on run properties. PrecisionRenderer never sets `a:ea`/`a:cs` — only `ppt_renderer.py` does this via `theme_mapper._ensure_ea_font()`. | Added implementation code showing how to set CJK font via XML: `rPr = run._r.get_or_add_rPr(); ea = etree.SubElement(rPr, qn('a:ea')); ea.set('typeface', cjk_font)`. Also: PrecisionRenderer must import and use CJK companion lookup. |
| 42 | 2.4 | Dark mode section divider uses `primary-950` for background, which may be nearly identical to the dark-mode slide background (e.g., `#060B18`), defeating the "pattern break" purpose. | Changed: in dark mode, use `primary-50` (light tint) instead of `primary-950` for section divider background. This creates actual contrast — a light divider between dark content slides is a strong pattern break. |
| 43 | 3.1 | `random.seed(42)` makes all presentations share identical noise texture. If slides from different decks are compared side-by-side, identical noise looks like a copy-paste error. | Changed: use `hashlib.md5(presentation_title.encode()).hexdigest()[:8]` as seed instead of fixed 42. Falls back to 42 if no title. This gives per-deck deterministic noise while varying across decks. |
| 44 | 3.2 | Progress bar at `SLIDE_HEIGHT - 0.06` overlaps with bottom muted bar at `SLIDE_HEIGHT - 0.15` (from §1.6). Both occupy the same bottom region. | Fixed: progress bar placed at `SLIDE_HEIGHT - 0.04` (inside the muted bar area). The muted bar is reduced to 0.12" height (from 0.15") and the progress bar sits at its bottom edge. Alternatively, when progress bar is present, the muted bar is omitted. |
| 45 | 3.6 | Two-column bullet logic duplicates §1.1 LayoutEngine's bullet handling. LayoutEngine decides layout (full-width vs two-column), renderer should only execute. | Added note: §3.6 rendering code must check `layout.bullet_columns` from LayoutEngine output rather than independently deciding column count. LayoutEngine owns the layout decision; renderer owns the execution. |
| 46 | 3.7 | "Asymmetric" hero pattern has no layout specification — no coordinates for text position or image offset. | Added specification: image at (0, 0, 9.333, 7.5) left-aligned, gradient overlay on right 40%, title at (8.0, 2.0, 4.8, 1.5) right-aligned, subtitle at (8.0, 3.6, 4.8, 0.8) right-aligned. |
| 47 | 1.1 | `ContentLayout` dataclass is returned by `compute_content_layout()` but never defined anywhere in the plan. | Added `ContentLayout` dataclass definition with all fields: `title_x, title_y, title_w, title_h, content_x, content_y, content_w, image_x, image_y, image_w, bullet_columns`. |
| 48 | 1.1 | `Rect` dataclass needs computed properties (`content_width`, `margin_top`, etc.) but plan only says "add Rect dataclass" without specifying these. | Added `Rect` dataclass definition with fields: `width, height, margin_left, margin_right, margin_top, margin_bottom` and computed property `content_width = width - margin_left - margin_right`, `content_height = height - margin_top - margin_bottom`. |
| 49 | 1.4 | `ppt_renderer.py`'s `_apply_dark_overlay()` (lines 143-163) uses pure black `RGBColor(0,0,0)` with alpha 55000 (55%). Plan only shows PrecisionRenderer replacement code, not ppt_renderer replacement. | Added: ppt_renderer's `_apply_dark_overlay()` must also be replaced with gradient overlay using the same `GradientFill` approach. Code added to §1.4. |
| 50 | 1.2 | `shape_factory.py:148` uses `Pt(font_size - 3)` with default `font_size=13`, producing **10pt** — below the 11pt minimum. Plan lists shape_factory.py in "Files to Modify" but has no specific fix. | Added explicit instruction: in shape_factory.py, change `Pt(font_size - 3)` to `Pt(max(11, font_size - 3))` to enforce 11pt minimum. |
| 51 | 2.6 | `pipeline.py` calls `precision.render_slide(prs, design, component_lib=component_lib)` (line 391). Plan says "pass layout_variant" but doesn't show how to extract it from `design` dict or pass it. | Added code: `layout_variant = design.get("layout_variant") if isinstance(design, dict) else None; precision.render_slide(prs, design, component_lib=component_lib, layout_variant=layout_variant)`. Also: `render_slide()` signature must add `layout_variant: dict | None = None` parameter. |
| 52 | 1.4 | Gradient overlay alpha comment says "percentage * 1000 (e.g., 80% = 80000)" but `GradientStop.alpha` stores raw OOXML units (0-100000 where 100000 = opaque). The math `int(0.80 * 100000) = 80000` is correct, but the comment is misleading — it implies the scale is 0-1000 when it's actually 0-100000. | Corrected comment: "Alpha values: OOXML units 0-100000 where 100000 = fully opaque, 0 = fully transparent." |
