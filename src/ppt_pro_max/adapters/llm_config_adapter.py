"""LLM Config Adapter — auto-detect LLM configuration from host AI coding tools.

Supported hosts:
- OpenCode  (opencode.json / auth.json)
- Codex CLI (config.toml / auth.json)
- Claude Code (settings.json / .credentials.json)

Image-capable providers are mapped:
- OpenAI (gpt-image / dalle) — model name determines which
- Gemini (gemini-*-image) — native image generation
- Doubao/Volcengine → Seedream
- Tongyi/Aliyun → Wanx
- Moonshot/Kimi → Kimi enhance
"""

from __future__ import annotations

import json
import os
import platform
import re
from pathlib import Path
from typing import Any


HOST_TO_PPT_PROVIDER: dict[str, str] = {
    "doubao": "seedream",
    "volcengine": "seedream",
    "tongyi": "wanx",
    "aliyun": "wanx",
    "moonshot": "kimi",
    "kimi": "kimi",
    "gemini": "gemini",
    "google": "gemini",
}

IMAGE_CAPABLE_PROVIDER_IDS: set[str] = set(HOST_TO_PPT_PROVIDER.keys()) | {"openai"}

PROVIDER_ENV_KEY_MAP: dict[str, str] = {
    "seedream": "ARK_API_KEY",
    "doubao": "ARK_API_KEY",
    "volcengine": "ARK_API_KEY",
    "dalle": "OPENAI_API_KEY",
    "gpt-image": "OPENAI_API_KEY",
    "openai": "OPENAI_API_KEY",
    "gemini": "GEMINI_API_KEY",
    "google": "GEMINI_API_KEY",
    "wanx": "DASHSCOPE_API_KEY",
    "tongyi": "DASHSCOPE_API_KEY",
    "aliyun": "DASHSCOPE_API_KEY",
    "kimi": "MOONSHOT_API_KEY",
    "moonshot": "MOONSHOT_API_KEY",
}

PROVIDER_DEFAULT_MODEL: dict[str, str] = {
    "seedream": "doubao-seedream-5-0-260128",
    "gpt-image": "gpt-image-1",
    "dalle": "dall-e-3",
    "gemini": "gemini-2.5-flash-image",
    "wanx": "wanx-v1",
    "kimi": "kimi-k2-0711-preview",
}

PROVIDER_DEFAULT_BASE_URL: dict[str, str] = {
    "seedream": "https://ark.cn-beijing.volces.com/api/v3",
    "gpt-image": "https://api.openai.com/v1",
    "dalle": "https://api.openai.com/v1",
    "gemini": "https://generativelanguage.googleapis.com/v1beta",
    "wanx": "https://dashscope.aliyuncs.com/api/v1",
    "kimi": "https://api.moonshot.cn/v1",
}


def detect_host_llm_config() -> dict[str, str]:
    for detector in (_detect_opencode, _detect_codex, _detect_claude_code):
        try:
            config = detector()
            if config and config.get("llm_api_key"):
                return config
        except Exception:
            continue
    return {}


def _opencode_config_path() -> Path:
    override = os.environ.get("OPENCODE_CONFIG")
    if override:
        return Path(override)
    config_dir = os.environ.get("OPENCODE_CONFIG_DIR")
    if config_dir:
        return Path(config_dir) / "opencode.json"
    if platform.system() == "Windows":
        return Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")) / "opencode" / "opencode.json"
    return Path.home() / ".config" / "opencode" / "opencode.json"


def _opencode_auth_path() -> Path:
    if platform.system() == "Windows":
        return Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / "opencode" / "auth.json"
    return Path.home() / ".local" / "share" / "opencode" / "auth.json"


def _resolve_opencode_var(value: str) -> str:
    if not isinstance(value, str):
        return value
    if value.startswith("{env:") and value.endswith("}"):
        return os.environ.get(value[5:-1], "")
    if value.startswith("{file:") and value.endswith("}"):
        file_path = value[6:-1].replace("~", str(Path.home()))
        try:
            return Path(file_path).read_text(encoding="utf-8").strip()
        except (FileNotFoundError, OSError):
            return ""
    return value


def _read_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}


def _deep_merge(base: dict, override: dict) -> dict:
    result = base.copy()
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def _map_provider(host_provider_id: str, model: str = "") -> str | None:
    direct = HOST_TO_PPT_PROVIDER.get(host_provider_id)
    if direct:
        return direct

    if host_provider_id == "openai":
        model_lower = (model or "").lower()
        if "gpt-image" in model_lower:
            return "gpt-image"
        if "dall-e" in model_lower or "dalle" in model_lower:
            return "dalle"
        return "gpt-image"

    model_lower = (model or "").lower()
    if "gpt-image" in model_lower:
        return "gpt-image"
    if "dall-e" in model_lower:
        return "dalle"
    if "gemini" in model_lower and "image" in model_lower:
        return "gemini"
    if "seedream" in model_lower:
        return "seedream"
    if "wanx" in model_lower:
        return "wanx"
    return None


def _resolve_api_key_for_provider(ppt_provider: str, explicit_key: str = "") -> str:
    if explicit_key:
        return explicit_key
    env_var = PROVIDER_ENV_KEY_MAP.get(ppt_provider, "")
    if env_var:
        return os.environ.get(env_var, "")
    return ""


def _detect_opencode() -> dict[str, str]:
    global_config = _read_json(_opencode_config_path())

    project_config_path = Path("opencode.json")
    project_config = _read_json(project_config_path)

    config = _deep_merge(global_config, project_config)

    model_str = config.get("model", "")
    provider_id = model_str.split("/")[0] if "/" in model_str else ""

    if not provider_id:
        return {}

    ppt_provider = _map_provider(provider_id, model_str)
    if not ppt_provider:
        return {}

    provider_conf = config.get("provider", {}).get(provider_id, {})
    api_key_raw = provider_conf.get("options", {}).get("apiKey", "")
    api_key = _resolve_opencode_var(api_key_raw)

    if not api_key:
        auth = _read_json(_opencode_auth_path())
        api_key = auth.get(provider_id, {}).get("apiKey", "")
        if not api_key:
            api_key = auth.get(provider_id, "")

    api_key = _resolve_api_key_for_provider(ppt_provider, api_key)
    if not api_key:
        return {}

    base_url = _resolve_opencode_var(provider_conf.get("options", {}).get("baseURL", ""))
    llm_model = ""
    model_id = model_str.split("/")[-1] if "/" in model_str else ""
    model_conf = provider_conf.get("models", {}).get(model_id, {})
    if model_conf:
        llm_model = model_conf.get("id", model_id)

    result: dict[str, str] = {"llm_provider": ppt_provider, "llm_api_key": api_key}
    if base_url:
        result["llm_base_url"] = base_url
    if llm_model:
        result["llm_model"] = llm_model
    result["detected_from"] = "opencode"
    return result


def _codex_config_path() -> Path:
    base = os.environ.get("CODEX_HOME", str(Path.home() / ".codex"))
    return Path(base) / "config.toml"


def _parse_toml_simple(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        import tomllib
        with open(path, "rb") as f:
            return tomllib.load(f)
    except ImportError:
        pass
    except Exception:
        return {}
    try:
        import tomli
        with open(path, "rb") as f:
            return tomli.load(f)
    except ImportError:
        pass
    except Exception:
        return {}
    return _parse_toml_regex(path)


def _parse_toml_regex(path: Path) -> dict[str, Any]:
    result: dict[str, Any] = {}
    current_section = result

    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return {}

    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue

        section_match = re.match(r'^\[([^\]]+)\]$', stripped)
        if section_match:
            section_path = section_match.group(1).strip().split(".")
            current_section = result
            for part in section_path:
                if part not in current_section:
                    current_section[part] = {}
                current_section = current_section[part]
            continue

        kv_match = re.match(r'^(\w+)\s*=\s*"([^"]*)"', stripped)
        if kv_match:
            current_section[kv_match.group(1)] = kv_match.group(2)
            continue

        kv_match2 = re.match(r"^(\w+)\s*=\s*'([^']*)'", stripped)
        if kv_match2:
            current_section[kv_match2.group(1)] = kv_match2.group(2)
            continue

        kv_bare = re.match(r'^(\w+)\s*=\s*(\S+)', stripped)
        if kv_bare:
            val = kv_bare.group(2)
            if val.lower() == "true":
                current_section[kv_bare.group(1)] = True
            elif val.lower() == "false":
                current_section[kv_bare.group(1)] = False
            else:
                try:
                    current_section[kv_bare.group(1)] = int(val)
                except ValueError:
                    try:
                        current_section[kv_bare.group(1)] = float(val)
                    except ValueError:
                        current_section[kv_bare.group(1)] = val

    return result


def _detect_codex() -> dict[str, str]:
    config = _parse_toml_simple(_codex_config_path())

    model = config.get("model", "")
    provider_id = str(config.get("model_provider", "openai")).lower()

    ppt_provider = _map_provider(provider_id, model)
    if not ppt_provider:
        return {}

    providers = config.get("model_providers", {})
    provider_conf = providers.get(provider_id, {})
    env_key = str(provider_conf.get("env_key", ""))
    base_url = str(provider_conf.get("base_url", ""))

    api_key = ""
    if env_key:
        api_key = os.environ.get(env_key, "")

    if not api_key:
        codex_home = os.environ.get("CODEX_HOME", str(Path.home() / ".codex"))
        auth = _read_json(Path(codex_home) / "auth.json")
        api_key = auth.get("access_token", "") or auth.get("api_key", "")

    api_key = _resolve_api_key_for_provider(ppt_provider, api_key)
    if not api_key:
        return {}

    result: dict[str, str] = {"llm_provider": ppt_provider, "llm_api_key": api_key}
    if base_url:
        result["llm_base_url"] = base_url
    if model:
        result["llm_model"] = model
    result["detected_from"] = "codex"
    return result


def _claude_config_dir() -> Path:
    override = os.environ.get("CLAUDE_CONFIG_DIR")
    if override:
        return Path(override)
    return Path.home() / ".claude"


def _detect_claude_code() -> dict[str, str]:
    config_dir = _claude_config_dir()
    settings = _read_json(config_dir / "settings.json")

    env_block = settings.get("env", {})

    for key_var in ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN"):
        api_key = os.environ.get(key_var, "") or env_block.get(key_var, "")
        if api_key:
            model = os.environ.get("ANTHROPIC_MODEL", "")
            base_url = os.environ.get("ANTHROPIC_BASE_URL", "") or env_block.get("ANTHROPIC_BASE_URL", "")

            ppt_provider = _infer_provider_from_base_url(base_url) or _infer_provider_from_model(model) or "gpt-image"

            if ppt_provider in ("seedream",):
                api_key = os.environ.get("ARK_API_KEY", api_key)
            elif ppt_provider in ("wanx",):
                api_key = os.environ.get("DASHSCOPE_API_KEY", api_key)
            elif ppt_provider in ("kimi",):
                api_key = os.environ.get("MOONSHOT_API_KEY", api_key)
            elif ppt_provider == "gemini":
                api_key = os.environ.get("GEMINI_API_KEY", api_key)

            api_key = _resolve_api_key_for_provider(ppt_provider, api_key)
            if not api_key:
                continue

            result: dict[str, str] = {"llm_provider": ppt_provider, "llm_api_key": api_key}
            if base_url:
                result["llm_base_url"] = base_url
            if model:
                result["llm_model"] = model
            result["detected_from"] = "claude-code"
            return result

    helper = settings.get("apiKeyHelper")
    if helper:
        try:
            import subprocess
            proc = subprocess.run(helper, shell=True, capture_output=True, text=True, timeout=30)
            if proc.returncode == 0 and proc.stdout.strip():
                api_key = proc.stdout.strip()
                base_url = os.environ.get("ANTHROPIC_BASE_URL", "")
                ppt_provider = _infer_provider_from_base_url(base_url) or "gpt-image"

                result = {"llm_provider": ppt_provider, "llm_api_key": api_key}
                if base_url:
                    result["llm_base_url"] = base_url
                result["detected_from"] = "claude-code"
                return result
        except Exception:
            pass

    return {}


def _infer_provider_from_base_url(base_url: str) -> str | None:
    if not base_url:
        return None
    if "volces.com" in base_url or "ark." in base_url:
        return "seedream"
    if "dashscope" in base_url or "aliyuncs" in base_url:
        return "wanx"
    if "moonshot" in base_url or "kimi" in base_url:
        return "kimi"
    if "generativelanguage.googleapis" in base_url or "aiplatform.googleapis" in base_url:
        return "gemini"
    return None


def _infer_provider_from_model(model: str) -> str | None:
    if not model:
        return None
    model_lower = model.lower()
    if "gpt-image" in model_lower:
        return "gpt-image"
    if "dall-e" in model_lower or "dalle" in model_lower:
        return "dalle"
    if "gemini" in model_lower and "image" in model_lower:
        return "gemini"
    if "seedream" in model_lower:
        return "seedream"
    if "wanx" in model_lower:
        return "wanx"
    return None
