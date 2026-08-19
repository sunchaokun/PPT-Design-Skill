import hashlib
import math
import os
import random
import tempfile
from pathlib import Path

from PIL import Image as PILImage
from PIL import ImageFilter, ImageDraw


_CACHE_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "..", ".cache", "graded")
_CACHE_DIR = os.path.normpath(_CACHE_DIR)

_EFFECTS_CACHE_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "..", ".cache", "effects")
_EFFECTS_CACHE_DIR = os.path.normpath(_EFFECTS_CACHE_DIR)


def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def grade_image_to_palette(image_path: str, palette_hex: str,
                            alpha: float = 0.10) -> str:
    src = Path(image_path)
    if not src.exists():
        return image_path
    key = hashlib.md5(f"{src.stat().st_mtime}_{palette_hex}_{alpha}".encode()).hexdigest()
    ext = src.suffix.lower()
    if ext in (".jpg", ".jpeg"):
        out_name = f"{key}.jpg"
    else:
        out_name = f"{key}.png"
    out_dir = Path(_CACHE_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / out_name
    if out_path.exists():
        return str(out_path)
    img = PILImage.open(str(src)).convert("RGB")
    overlay = PILImage.new("RGB", img.size, _hex_to_rgb(palette_hex))
    blended = PILImage.blend(img, overlay, alpha)
    if ext in (".jpg", ".jpeg"):
        blended.save(str(out_path), "JPEG", quality=92)
    else:
        blended.save(str(out_path), "PNG")
    return str(out_path)


def generate_noise_tile(size: int = 200, opacity: float = 0.02,
                         deck_title: str = "") -> str:
    noise_dir = os.path.join(tempfile.gettempdir(), "ppt-noise")
    os.makedirs(noise_dir, exist_ok=True)
    seed_part = hashlib.md5(deck_title.encode()).hexdigest()[:8] if deck_title else "default"
    cache_path = os.path.join(noise_dir, f"noise_{int(opacity * 1000)}_{seed_part}.png")
    if os.path.exists(cache_path):
        return cache_path
    seed_val = int(hashlib.md5(deck_title.encode()).hexdigest()[:8], 16) if deck_title else 42
    rng = random.Random(seed_val)
    img = PILImage.new("RGBA", (size, size))
    for y in range(size):
        for x in range(size):
            u1 = rng.random() or 0.001
            u2 = rng.random()
            z = (-2.0 * math.log(u1)) ** 0.5 * math.cos(2.0 * math.pi * u2)
            val = int(max(0, min(255, 128 + z * 20)))
            alpha_val = int(opacity * 255)
            img.putpixel((x, y), (val, val, val, alpha_val))
    img.save(cache_path, "PNG")
    return cache_path


def _effects_cache_path(image_path: str, effect_name: str, **kwargs) -> Path:
    out_dir = Path(_EFFECTS_CACHE_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)
    src = Path(image_path)
    parts = [str(src.stat().st_mtime), effect_name]
    for v in kwargs.values():
        parts.append(str(v))
    key = hashlib.md5("_".join(parts).encode()).hexdigest()
    ext = src.suffix.lower()
    if ext in (".jpg", ".jpeg"):
        return out_dir / f"{key}.jpg"
    return out_dir / f"{key}.png"


def _save_with_format(img: PILImage.Image, out_path: Path, src_path: str) -> str:
    ext = Path(src_path).suffix.lower()
    if ext in (".jpg", ".jpeg"):
        rgb = img.convert("RGB")
        rgb.save(str(out_path), "JPEG", quality=92)
    else:
        img.save(str(out_path), "PNG")
    return str(out_path)


def apply_grayscale(image_path: str) -> str:
    src = Path(image_path)
    if not src.exists():
        return image_path
    out_path = _effects_cache_path(image_path, "grayscale")
    if out_path.exists():
        return str(out_path)
    img = PILImage.open(str(src)).convert("L")
    return _save_with_format(img, out_path, image_path)


def apply_sepia(image_path: str, intensity: float = 0.5) -> str:
    src = Path(image_path)
    if not src.exists():
        return image_path
    if intensity <= 0.0:
        return image_path
    out_path = _effects_cache_path(image_path, "sepia", intensity=intensity)
    if out_path.exists():
        return str(out_path)
    img = PILImage.open(str(src)).convert("RGB")

    try:
        import numpy as np
        arr = np.array(img, dtype=np.float32)
        sepia_matrix = np.array([
            [0.393, 0.769, 0.189],
            [0.349, 0.686, 0.168],
            [0.272, 0.534, 0.131]
        ])
        sepia = arr @ sepia_matrix.T
        sepia = np.clip(sepia, 0, 255)
        result = arr + (sepia - arr) * intensity
        result = np.clip(result, 0, 255).astype(np.uint8)
        img = PILImage.fromarray(result, 'RGB')
    except ImportError:
        pixels = img.load()
        w, h = img.size
        for y in range(h):
            for x in range(w):
                r, g, b = pixels[x, y]
                tr = int(min(255, 0.393 * r + 0.769 * g + 0.189 * b))
                tg = int(min(255, 0.349 * r + 0.686 * g + 0.168 * b))
                tb = int(min(255, 0.272 * r + 0.534 * g + 0.131 * b))
                nr = int(r + (tr - r) * intensity)
                ng = int(g + (tg - g) * intensity)
                nb = int(b + (tb - b) * intensity)
                pixels[x, y] = (nr, ng, nb)

    return _save_with_format(img, out_path, image_path)


def apply_duotone(image_path: str, color1: str, color2: str) -> str:
    src = Path(image_path)
    if not src.exists():
        return image_path
    out_path = _effects_cache_path(image_path, "duotone", c1=color1, c2=color2)
    if out_path.exists():
        return str(out_path)
    img = PILImage.open(str(src)).convert("L")
    c1 = _hex_to_rgb(color1)
    c2 = _hex_to_rgb(color2)

    try:
        import numpy as np
        gray = np.array(img, dtype=np.float32) / 255.0
        result = np.zeros((*gray.shape, 3), dtype=np.float32)
        for i in range(3):
            result[:, :, i] = c2[i] + (c1[i] - c2[i]) * gray
        result = np.clip(result, 0, 255).astype(np.uint8)
        img = PILImage.fromarray(result, 'RGB')
    except ImportError:
        result = PILImage.new("RGB", img.size)
        pixels = result.load()
        gray_pixels = img.load()
        w, h = img.size
        for y in range(h):
            for x in range(w):
                t = gray_pixels[x, y] / 255.0
                r = int(c2[0] + (c1[0] - c2[0]) * t)
                g = int(c2[1] + (c1[1] - c2[1]) * t)
                b = int(c2[2] + (c1[2] - c2[2]) * t)
                pixels[x, y] = (r, g, b)

    return _save_with_format(img, out_path, image_path)


def apply_ink_wash(image_path: str, contrast: float = 1.5,
                   brightness: float = 0.0) -> str:
    src = Path(image_path)
    if not src.exists():
        return image_path
    out_path = _effects_cache_path(image_path, "ink_wash",
                                   contrast=contrast, brightness=brightness)
    if out_path.exists():
        return str(out_path)
    img = PILImage.open(str(src)).convert("L")
    from PIL import ImageEnhance
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(contrast)
    if brightness != 0.0:
        enhancer = ImageEnhance.Brightness(img)
        img = enhancer.enhance(1.0 + brightness)
    tint = PILImage.new("RGB", img.size, (245, 240, 232))
    img_rgb = img.convert("RGB")
    result = PILImage.blend(img_rgb, tint, 0.15)
    return _save_with_format(result, out_path, image_path)


def apply_blur(image_path: str, radius: int = 5) -> str:
    src = Path(image_path)
    if not src.exists():
        return image_path
    if radius <= 0:
        return image_path
    out_path = _effects_cache_path(image_path, "blur", radius=radius)
    if out_path.exists():
        return str(out_path)
    img = PILImage.open(str(src))
    result = img.filter(ImageFilter.GaussianBlur(radius=radius))
    return _save_with_format(result, out_path, image_path)


def apply_vignette(image_path: str, intensity: float = 0.5) -> str:
    src = Path(image_path)
    if not src.exists():
        return image_path
    if intensity <= 0.0:
        return image_path
    out_path = _effects_cache_path(image_path, "vignette", intensity=intensity)
    if out_path.exists():
        return str(out_path)
    img = PILImage.open(str(src)).convert("RGBA")
    w, h = img.size
    mask = PILImage.new("L", (w, h), 255)
    draw = ImageDraw.Draw(mask)
    cx, cy = w // 2, h // 2
    max_radius = int(math.sqrt(cx * cx + cy * cy))
    steps = 20
    for i in range(steps):
        ratio = 1.0 - (i / steps) * 0.5
        r = int(max_radius * ratio)
        alpha = int(255 * (1.0 - intensity * (1.0 - i / steps)))
        draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=alpha)
    img.putalpha(mask)
    return _save_with_format(img, out_path, image_path)


def apply_edge_fade(image_path: str, margin_pct: float = 0.1,
                    bg_color: str | None = None) -> str:
    src = Path(image_path)
    if not src.exists():
        return image_path
    out_path = _effects_cache_path(image_path, "edge_fade",
                                   margin=margin_pct, bg=bg_color or "none")
    if out_path.exists():
        return str(out_path)
    img = PILImage.open(str(src)).convert("RGBA")
    w, h = img.size
    margin_x = int(w * margin_pct)
    margin_y = int(h * margin_pct)
    alpha = PILImage.new("L", (w, h), 255)
    pixels = alpha.load()
    for y in range(h):
        for x in range(w):
            fx = 1.0
            if x < margin_x:
                fx = x / max(margin_x, 1)
            elif x > w - margin_x:
                fx = (w - x) / max(margin_x, 1)
            fy = 1.0
            if y < margin_y:
                fy = y / max(margin_y, 1)
            elif y > h - margin_y:
                fy = (h - y) / max(margin_y, 1)
            pixels[x, y] = int(min(fx, fy) * 255)
    img.putalpha(alpha)
    if bg_color:
        bg = PILImage.new("RGB", (w, h), _hex_to_rgb(bg_color))
        bg_rgba = bg.convert("RGBA")
        result = PILImage.alpha_composite(bg_rgba, img)
        return _save_with_format(result, out_path, image_path)
    return _save_with_format(img, out_path, image_path)


# ── Multi-layer Composition ──

def compose_images(layers, width=1920, height=1080, bg_color='#000000'):
    """Composite multiple PIL Image layers into a single image.

    Args:
        layers: List of dicts, each with:
            - 'image': PILImage.Image or str (path) (required)
            - 'opacity': float 0.0-1.0 (default 1.0)
            - 'position': tuple (x, y) in pixels (default (0, 0))
        width: Output width in pixels
        height: Output height in pixels
        bg_color: Background color hex (default '#000000')

    Returns:
        PILImage.Image (RGBA)
    """
    result = PILImage.new('RGBA', (width, height), _hex_to_rgb(bg_color) + (255,))

    for layer in layers:
        img = layer.get('image')
        if isinstance(img, (str, Path)):
            img = PILImage.open(img)
        if img is None:
            continue

        img = img.convert('RGBA')
        opacity = layer.get('opacity', 1.0)
        pos = layer.get('position', (0, 0))

        if opacity < 1.0:
            alpha = img.split()[3]
            alpha = alpha.point(lambda p: int(p * opacity))
            img.putalpha(alpha)

        result.paste(img, pos, img)

    return result


# ── Directional Gradient Mask ──

def apply_gradient_mask(image_path, direction='bottom', color='#000000',
                        start_opacity=0.0, end_opacity=0.8):
    """Apply directional gradient mask to an image.

    Args:
        image_path: Path to source image
        direction: 'top'|'bottom'|'left'|'right'|'diagonal_tl'|'diagonal_br'
        color: Mask color hex
        start_opacity: Starting opacity (0.0-1.0)
        end_opacity: Ending opacity (0.0-1.0)

    Returns:
        Path to processed image
    """
    src = Path(image_path)
    if not src.exists():
        return image_path

    out_path = _effects_cache_path(image_path, "grad_mask",
                                   dir=direction, c=color,
                                   s=start_opacity, e=end_opacity)
    if out_path.exists():
        return str(out_path)

    img = PILImage.open(str(src)).convert('RGBA')
    w, h = img.size
    mask = PILImage.new('L', (w, h), 0)
    pixels = mask.load()

    r, g, b = _hex_to_rgb(color)

    for y in range(h):
        for x in range(w):
            if direction == 'bottom':
                t = y / max(h - 1, 1)
            elif direction == 'top':
                t = 1.0 - y / max(h - 1, 1)
            elif direction == 'right':
                t = x / max(w - 1, 1)
            elif direction == 'left':
                t = 1.0 - x / max(w - 1, 1)
            elif direction == 'diagonal_tl':
                t = (x / max(w - 1, 1) + y / max(h - 1, 1)) / 2
            elif direction == 'diagonal_br':
                t = 1.0 - (x / max(w - 1, 1) + y / max(h - 1, 1)) / 2
            else:
                t = 0

            opacity = start_opacity + (end_opacity - start_opacity) * t
            pixels[x, y] = int(opacity * 255)

    overlay = PILImage.new('RGBA', (w, h), (r, g, b, 0))
    overlay.putalpha(mask)
    result = PILImage.alpha_composite(img, overlay)

    return _save_with_format(result, out_path, image_path)


# ── Scatter Particles ──

def apply_scatter(image_path, count=50, color='#FFFFFF',
                  min_size=2, max_size=15, min_alpha=20, max_alpha=120,
                  distribution='random', seed=None):
    """Add scattered particles to an image.

    Args:
        image_path: Path to source image
        count: Number of particles
        color: Particle color hex
        min_size: Minimum particle size in pixels
        max_size: Maximum particle size in pixels
        min_alpha: Minimum alpha (0-255)
        max_alpha: Maximum alpha (0-255)
        distribution: 'random'|'center'|'edge'|'top'|'bottom'
        seed: Random seed for reproducibility

    Returns:
        Path to processed image
    """
    import random as _random

    src = Path(image_path)
    if not src.exists():
        return image_path

    out_path = _effects_cache_path(image_path, "scatter",
                                   count=count, c=color, dist=distribution,
                                   seed=seed)
    if out_path.exists():
        return str(out_path)

    img = PILImage.open(str(src)).convert('RGBA')
    w, h = img.size
    overlay = PILImage.new('RGBA', (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    rng = _random.Random(seed)
    r, g, b = _hex_to_rgb(color)

    for _ in range(count):
        size = rng.randint(min_size, max_size)
        alpha = rng.randint(min_alpha, max_alpha)

        if distribution == 'center':
            cx, cy = w // 2 + rng.randint(-w // 4, w // 4), h // 2 + rng.randint(-h // 4, h // 4)
        elif distribution == 'edge':
            if rng.random() < 0.5:
                cx = rng.choice([rng.randint(0, w // 6), rng.randint(w * 5 // 6, w - 1)])
                cy = rng.randint(0, h - 1)
            else:
                cx = rng.randint(0, w - 1)
                cy = rng.choice([rng.randint(0, h // 6), rng.randint(h * 5 // 6, h - 1)])
        elif distribution == 'top':
            cx, cy = rng.randint(0, w - 1), rng.randint(0, h // 3)
        elif distribution == 'bottom':
            cx, cy = rng.randint(0, w - 1), rng.randint(h * 2 // 3, h - 1)
        else:  # random
            cx, cy = rng.randint(0, w - 1), rng.randint(0, h - 1)

        x1, y1 = cx - size // 2, cy - size // 2
        x2, y2 = cx + size // 2, cy + size // 2
        draw.ellipse([x1, y1, x2, y2], fill=(r, g, b, alpha))

    result = PILImage.alpha_composite(img, overlay)
    return _save_with_format(result, out_path, image_path)
