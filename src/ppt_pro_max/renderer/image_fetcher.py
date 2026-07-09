"""Image Fetcher — five modes: placeholder / search / generate / enhance / auto."""

from __future__ import annotations

import base64
import hashlib
import json
import os
import tempfile
import time
import urllib.request
import urllib.parse
import urllib.error
from pathlib import Path
from typing import Any


class ImageFetcher:
    def __init__(
        self,
        mode: str = "placeholder",
        unsplash_access_key: str | None = None,
        pexels_api_key: str | None = None,
        llm_provider: str | None = None,
        llm_api_key: str | None = None,
        llm_base_url: str | None = None,
        llm_model: str | None = None,
        image_cache_dir: str | None = None,
    ):
        self.mode = mode
        self.unsplash_access_key = unsplash_access_key or os.environ.get("UNSPLASH_ACCESS_KEY", "")
        self.pexels_api_key = pexels_api_key or os.environ.get("PEXELS_API_KEY", "")
        self.llm_provider = (llm_provider or os.environ.get("PPT_IMAGE_LLM_PROVIDER", "")).lower()
        self.llm_api_key = llm_api_key or os.environ.get("PPT_IMAGE_LLM_API_KEY", "")
        self.llm_base_url = llm_base_url or os.environ.get("PPT_IMAGE_LLM_BASE_URL", "")
        self.llm_model = llm_model or os.environ.get("PPT_IMAGE_LLM_MODEL", "")

        if image_cache_dir:
            self._cache_dir = Path(image_cache_dir)
        else:
            self._cache_dir = Path(tempfile.gettempdir()) / "ppt-design-skill-images"
        self._cache_dir.mkdir(parents=True, exist_ok=True)

    def fetch(self, keywords: str, emotion: str = "", goal: str = "", width: int = 1920, height: int = 1080) -> str | None:
        if self.mode == "placeholder":
            return None

        if self.mode == "search":
            return self._fetch_from_search(keywords, emotion, width, height)

        if self.mode == "generate":
            return self._fetch_from_llm_generate(keywords, emotion, goal, width, height)

        if self.mode == "enhance":
            return self._fetch_with_llm_enhance(keywords, emotion, goal, width, height)

        if self.mode == "auto":
            return self._fetch_auto(keywords, emotion, goal, width, height)

        return None

    def _fetch_auto(self, keywords: str, emotion: str, goal: str, width: int, height: int) -> str | None:
        if self.llm_provider:
            result = self._fetch_from_llm_generate(keywords, emotion, goal, width, height)
            if result:
                return result
        result = self._fetch_from_search(keywords, emotion, width, height)
        return result

    def _fetch_from_search(self, keywords: str, emotion: str, width: int, height: int) -> str | None:
        cached = self._check_cache(f"search:{keywords}:{width}x{height}")
        if cached:
            return cached

        if self.unsplash_access_key:
            result = self._fetch_unsplash(keywords, width, height)
            if result:
                return result

        if self.pexels_api_key:
            result = self._fetch_pexels(keywords, width, height)
            if result:
                return result

        result = self._fetch_unsplash_source(keywords, width, height)
        return result

    def _fetch_unsplash(self, keywords: str, width: int, height: int) -> str | None:
        try:
            query = urllib.parse.quote(keywords)
            url = f"https://api.unsplash.com/search/photos?query={query}&per_page=1&orientation=landscape"
            req = urllib.request.Request(url, headers={
                "Authorization": f"Client-ID {self.unsplash_access_key}",
                "Accept-Version": "v1",
            })
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            results = data.get("results", [])
            if not results:
                return None
            photo_url = results[0]["urls"].get("regular", "")
            if not photo_url:
                return None
            photo_url = photo_url.replace("w=1080", f"w={width}")
            photo_url += f"&h={height}&fit=crop"
            return self._download_and_cache(photo_url, f"unsplash:{keywords}:{width}x{height}")
        except Exception:
            return None

    def _fetch_pexels(self, keywords: str, width: int, height: int) -> str | None:
        try:
            query = urllib.parse.quote(keywords)
            url = f"https://api.pexels.com/v1/search?query={query}&per_page=1&orientation=landscape"
            req = urllib.request.Request(url, headers={
                "Authorization": self.pexels_api_key,
            })
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            photos = data.get("photos", [])
            if not photos:
                return None
            photo_url = photos[0]["src"].get("large2x", photos[0]["src"].get("large", ""))
            if not photo_url:
                return None
            return self._download_and_cache(photo_url, f"pexels:{keywords}:{width}x{height}")
        except Exception:
            return None

    def _fetch_unsplash_source(self, keywords: str, width: int, height: int) -> str | None:
        try:
            query = urllib.parse.quote(keywords.replace(" ", "-"))
            url = f"https://source.unsplash.com/{width}x{height}/?{query}"
            return self._download_and_cache(url, f"unsplash-src:{keywords}:{width}x{height}")
        except Exception:
            return None

    def _fetch_from_llm_generate(self, keywords: str, emotion: str, goal: str, width: int, height: int) -> str | None:
        prompt = self._build_image_prompt(keywords, emotion, goal)

        if self.llm_provider in ("seedream", "doubao", "volcengine"):
            return self._generate_seedream(prompt, width, height)

        if self.llm_provider in ("gpt-image", "gpt_image", "gptimage"):
            return self._generate_gpt_image(prompt, width, height)

        if self.llm_provider in ("dalle", "openai", "dall-e"):
            return self._generate_dalle(prompt, width, height)

        if self.llm_provider in ("wanx", "tongyi", "aliyun"):
            return self._generate_wanx(prompt, width, height)

        if self.llm_provider in ("kimi", "moonshot"):
            enhanced_keywords = self._kimi_enhance_keywords(keywords, emotion, goal)
            if enhanced_keywords:
                return self._fetch_from_search(enhanced_keywords, emotion, width, height)
            return self._fetch_from_search(keywords, emotion, width, height)

        return None

    def _fetch_with_llm_enhance(self, keywords: str, emotion: str, goal: str, width: int, height: int) -> str | None:
        enhanced = self._kimi_enhance_keywords(keywords, emotion, goal)
        search_keywords = enhanced if enhanced else keywords
        return self._fetch_from_search(search_keywords, emotion, width, height)

    def _build_image_prompt(self, keywords: str, emotion: str, goal: str) -> str:
        emotion_style = {
            "curiosity": "mysterious and intriguing",
            "frustration": "dark and tense",
            "hope": "bright and uplifting",
            "confidence": "bold and professional",
            "trust": "warm and authentic",
            "warmth": "cozy and inviting",
            "urgency": "dynamic and energetic",
            "fear": "dramatic and intense",
        }
        style = emotion_style.get(emotion, "professional and clean")

        goal_context = {
            "hook": "attention-grabbing hero image",
            "problem": "visualizing pain and frustration",
            "agitation": "conveying urgency and risk",
            "solution": "showing hope and transformation",
            "features": "clean product interface",
            "traction": "upward growth and success",
            "market": "global opportunity and scale",
            "team": "collaboration and expertise",
            "financial": "financial growth and charts",
            "demo": "product in action",
            "testimonials": "happy people and success",
            "pricing": "value and choice",
            "cta": "action and forward momentum",
            "product": "product showcase",
            "proof": "evidence and trust",
            "offer": "excitement and opportunity",
        }
        context = goal_context.get(goal, "professional presentation slide")
        return f"A {style} image representing {keywords}, suitable for a presentation slide about {context}, high quality, no text overlay, clean composition, suitable for corporate presentation"

    def _generate_seedream(self, prompt: str, width: int, height: int) -> str | None:
        try:
            api_key = self.llm_api_key or os.environ.get("ARK_API_KEY", "")
            if not api_key:
                return None

            base_url = self.llm_base_url or "https://ark.cn-beijing.volces.com/api/v3"
            model = self.llm_model or "doubao-seedream-5-0-pro-260628"

            if max(width, height) >= 2048:
                img_size = "2K"
            else:
                img_size = "1K"

            cache_key = f"seedream:{model}:{prompt[:80]}:{img_size}"
            cached = self._check_cache(cache_key)
            if cached:
                return cached

            payload = json.dumps({
                "model": model,
                "prompt": prompt,
                "size": img_size,
                "output_format": "png",
                "response_format": "url",
                "watermark": False,
            }).encode("utf-8")

            url = f"{base_url}/images/generations"
            req = urllib.request.Request(url, data=payload, headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
            })

            with urllib.request.urlopen(req, timeout=120) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            image_url = data.get("data", [{}])[0].get("url", "")
            if not image_url:
                return None
            return self._download_and_cache(image_url, f"seedream:{model}:{prompt[:50]}:{img_size}")
        except Exception:
            return None

    def _generate_gpt_image(self, prompt: str, width: int, height: int) -> str | None:
        try:
            api_key = self.llm_api_key or os.environ.get("OPENAI_API_KEY", "")
            if not api_key:
                return None

            base_url = self.llm_base_url or "https://api.openai.com/v1"
            model = self.llm_model or "gpt-image-1"

            if width > height:
                size = "1536x1024"
            elif height > width:
                size = "1024x1536"
            else:
                size = "1024x1024"

            cache_key = f"gpt-image:{model}:{prompt[:80]}:{size}"
            cached = self._check_cache(cache_key)
            if cached:
                return cached

            payload = json.dumps({
                "model": model,
                "prompt": prompt,
                "n": 1,
                "size": size,
                "quality": "low",
                "output_format": "png",
            }).encode("utf-8")

            url = f"{base_url}/images/generations"
            req = urllib.request.Request(url, data=payload, headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
            })

            with urllib.request.urlopen(req, timeout=120) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            item = data.get("data", [{}])[0]

            image_url = item.get("url", "")
            b64 = item.get("b64_json", "")

            if image_url:
                return self._download_and_cache(image_url, f"gpt-image:{model}:{prompt[:50]}")

            if b64:
                cache_key = f"gpt-image:{model}:{prompt[:50]}"
                cache_hash = hashlib.md5(cache_key.encode()).hexdigest()
                cache_path = self._cache_dir / f"{cache_hash}.png"
                if not (cache_path.exists() and cache_path.stat().st_size > 1000):
                    cache_path.write_bytes(base64.b64decode(b64))
                return str(cache_path)

            return None
        except Exception:
            return None

    def _generate_dalle(self, prompt: str, width: int, height: int) -> str | None:
        try:
            api_key = self.llm_api_key or os.environ.get("OPENAI_API_KEY", "")
            if not api_key:
                return None

            base_url = self.llm_base_url or "https://api.openai.com/v1"
            size = "1792x1024" if width > height else "1024x1792"
            if width <= 1024:
                size = "1024x1024"

            cache_key = f"dalle:{self.llm_model or 'dall-e-3'}:{prompt[:80]}:{size}"
            cached = self._check_cache(cache_key)
            if cached:
                return cached

            payload = json.dumps({
                "model": self.llm_model or "dall-e-3",
                "prompt": prompt,
                "n": 1,
                "size": size,
                "quality": "standard",
            }).encode("utf-8")

            url = f"{base_url}/images/generations"
            req = urllib.request.Request(url, data=payload, headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
            })
            with urllib.request.urlopen(req, timeout=60) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            image_url = data.get("data", [{}])[0].get("url", "")
            if not image_url:
                return None
            return self._download_and_cache(image_url, f"dalle:{prompt[:50]}:{size}")
        except Exception:
            return None

    def _generate_wanx(self, prompt: str, width: int, height: int) -> str | None:
        try:
            api_key = self.llm_api_key or os.environ.get("DASHSCOPE_API_KEY", "")
            if not api_key:
                return None

            base_url = self.llm_base_url or "https://dashscope.aliyuncs.com/api/v1"
            size_str = f"{width}*{height}"

            payload = json.dumps({
                "model": self.llm_model or "wanx-v1",
                "input": {"prompt": prompt},
                "parameters": {
                    "size": size_str,
                    "n": 1,
                },
            }).encode("utf-8")

            url = f"{base_url}/services/aigc/text2image/image-synthesis"
            req = urllib.request.Request(url, data=payload, headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
                "X-DashScope-Async": "enable",
            })
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            task_id = data.get("output", {}).get("task_id", "")
            if not task_id:
                return None

            for _ in range(30):
                time.sleep(2)
                status_url = f"{base_url}/tasks/{task_id}"
                status_req = urllib.request.Request(status_url, headers={
                    "Authorization": f"Bearer {api_key}",
                })
                with urllib.request.urlopen(status_req, timeout=10) as status_resp:
                    status_data = json.loads(status_resp.read().decode("utf-8"))

                task_status = status_data.get("output", {}).get("task_status", "")
                if task_status == "SUCCEEDED":
                    results = status_data.get("output", {}).get("results", [])
                    if results:
                        image_url = results[0].get("url", "")
                        if image_url:
                            return self._download_and_cache(image_url, f"wanx:{prompt[:50]}")
                    return None
                if task_status in ("FAILED", "UNKNOWN"):
                    return None

            return None
        except Exception:
            return None

    def _kimi_enhance_keywords(self, keywords: str, emotion: str, goal: str) -> str | None:
        try:
            api_key = self.llm_api_key or os.environ.get("MOONSHOT_API_KEY", "")
            if not api_key:
                return None

            base_url = self.llm_base_url or "https://api.moonshot.cn/v1"
            model = self.llm_model or "kimi-k2-0711-preview"

            system_msg = "You are an expert at selecting the perfect stock photo search keywords for presentation slides. Return ONLY 3-5 English search keywords separated by spaces, nothing else."
            user_msg = f"I need a presentation image for a slide about '{goal}' with emotion '{emotion}'. Original keywords: '{keywords}'. Give me better Unsplash search keywords."

            payload = json.dumps({
                "model": model,
                "messages": [
                    {"role": "system", "content": system_msg},
                    {"role": "user", "content": user_msg},
                ],
                "temperature": 0.3,
                "max_tokens": 50,
            }).encode("utf-8")

            url = f"{base_url}/chat/completions"
            req = urllib.request.Request(url, data=payload, headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
            })
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            enhanced = data.get("choices", [{}])[0].get("message", {}).get("content", "").strip()
            if enhanced and len(enhanced) < 100:
                return enhanced
            return None
        except Exception:
            return None

    def _download_and_cache(self, url: str, cache_key: str) -> str | None:
        try:
            cache_hash = hashlib.md5(cache_key.encode()).hexdigest()
            ext = ".jpg"
            if ".png" in url.lower():
                ext = ".png"
            cache_path = self._cache_dir / f"{cache_hash}{ext}"

            if cache_path.exists() and cache_path.stat().st_size > 1000:
                return str(cache_path)

            req = urllib.request.Request(url, headers={
                "User-Agent": "PPT-Design-Skill/0.1.0",
            })
            with urllib.request.urlopen(req, timeout=30) as resp:
                image_data = resp.read()

            if len(image_data) < 1000:
                return None

            cache_path.write_bytes(image_data)
            return str(cache_path)
        except Exception:
            return None

    def _check_cache(self, cache_key: str) -> str | None:
        cache_hash = hashlib.md5(cache_key.encode()).hexdigest()
        for ext in (".jpg", ".png", ".webp"):
            cache_path = self._cache_dir / f"{cache_hash}{ext}"
            if cache_path.exists() and cache_path.stat().st_size > 1000:
                return str(cache_path)
        return None

    @staticmethod
    def available_modes() -> dict[str, str]:
        return {
            "placeholder": "Gradient placeholder image (default, no API key needed)",
            "search": "Download from Unsplash/Pexels by keywords (needs API key)",
            "generate": "AI image generation: Seedream / GPT Image / DALL-E / Wanx (needs API key)",
            "enhance": "Use Kimi K2.6 to enhance search keywords, then download (needs Moonshot API key)",
            "auto": "Try AI generation first, fall back to search (recommended)",
        }

    @staticmethod
    def available_providers() -> dict[str, dict[str, str]]:
        return {
            "seedream": {
                "name": "Doubao Seedream (ByteDance Volcengine)",
                "env_key": "ARK_API_KEY",
                "default_model": "doubao-seedream-5-0-pro-260628",
                "models": "doubao-seedream-5-0-pro-260628, doubao-seedream-5-0-lite-250415",
                "base_url": "https://ark.cn-beijing.volces.com/api/v3",
            },
            "gpt-image": {
                "name": "GPT Image (OpenAI)",
                "env_key": "OPENAI_API_KEY",
                "default_model": "gpt-image-1",
                "models": "gpt-image-2, gpt-image-1.5, gpt-image-1, dall-e-3",
                "base_url": "https://api.openai.com/v1",
            },
            "dalle": {
                "name": "DALL-E 3 (OpenAI)",
                "env_key": "OPENAI_API_KEY",
                "default_model": "dall-e-3",
                "models": "dall-e-3",
                "base_url": "https://api.openai.com/v1",
            },
            "wanx": {
                "name": "Wanx (Alibaba DashScope)",
                "env_key": "DASHSCOPE_API_KEY",
                "default_model": "wanx-v1",
                "models": "wanx-v1",
                "base_url": "https://dashscope.aliyuncs.com/api/v1",
            },
            "kimi": {
                "name": "Kimi K2.6 (Moonshot, enhance-only)",
                "env_key": "MOONSHOT_API_KEY",
                "default_model": "kimi-k2-0711-preview",
                "models": "kimi-k2-0711-preview",
                "base_url": "https://api.moonshot.cn/v1",
            },
        }
