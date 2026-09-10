"""Small, privacy-conscious settings and credential helpers.

Ordinary preferences are stored as JSON in the operating system's user config
directory. API keys are never written to that JSON file; when the user opts in,
they are stored through the operating system credential service via ``keyring``.
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any

from .version import APP_NAME

SETTINGS_SCHEMA_VERSION = 1
_CREDENTIAL_SERVICE = "word-freq-analyzer"
_CREDENTIAL_ACCOUNT = "llm-api-key"


def settings_dir() -> Path:
    if sys.platform == "win32":
        base = Path(os.getenv("APPDATA") or Path.home() / "AppData" / "Roaming")
        return base / "WordFreqAnalyzer"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "WordFreqAnalyzer"
    base = Path(os.getenv("XDG_CONFIG_HOME") or Path.home() / ".config")
    return base / "word-freq-analyzer"


def settings_path() -> Path:
    return settings_dir() / "settings.json"


def load_settings() -> dict[str, Any]:
    path = settings_path()
    try:
        with path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
        if not isinstance(payload, dict):
            return {}
        payload.pop("api_key", None)  # discard credentials saved by older/dev builds
        return payload
    except (OSError, ValueError, TypeError):
        return {}


def save_settings(settings: dict[str, Any]) -> None:
    payload = dict(settings)
    payload.pop("api_key", None)
    payload["schema_version"] = SETTINGS_SCHEMA_VERSION
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".json.tmp")
    with temp.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
    os.replace(temp, path)


def load_api_key() -> str:
    try:
        import keyring

        return keyring.get_password(_CREDENTIAL_SERVICE, _CREDENTIAL_ACCOUNT) or ""
    except Exception:
        return ""


def save_api_key(api_key: str) -> bool:
    try:
        import keyring

        if api_key:
            keyring.set_password(_CREDENTIAL_SERVICE, _CREDENTIAL_ACCOUNT, api_key)
        else:
            try:
                keyring.delete_password(_CREDENTIAL_SERVICE, _CREDENTIAL_ACCOUNT)
            except Exception:
                pass
        return True
    except Exception:
        return False


def app_display_name() -> str:
    return APP_NAME
