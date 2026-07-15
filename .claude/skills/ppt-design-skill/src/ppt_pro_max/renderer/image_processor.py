import hashlib
import math
import os
import random
import tempfile
from pathlib import Path

from PIL import Image as PILImage


_CACHE_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "..", ".cache", "graded")
_CACHE_DIR = os.path.normpath(_CACHE_DIR)


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
