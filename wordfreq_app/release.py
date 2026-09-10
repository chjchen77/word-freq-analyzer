"""Manual update checks against the project's GitHub Releases page."""

from __future__ import annotations

import json
from urllib.request import Request, urlopen

from .version import APP_VERSION, GITHUB_REPOSITORY, version_tuple


def check_latest_release(timeout: float = 8.0) -> dict[str, str | bool]:
    url = f"https://api.github.com/repos/{GITHUB_REPOSITORY}/releases/latest"
    request = Request(url, headers={
        "Accept": "application/vnd.github+json",
        "User-Agent": f"word-freq-analyzer/{APP_VERSION}",
    })
    with urlopen(request, timeout=timeout) as response:
        payload = json.load(response)
    tag = str(payload.get("tag_name") or "").strip()
    html_url = str(payload.get("html_url") or "").strip()
    return {
        "current": APP_VERSION,
        "latest": tag.lstrip("v") or "未知",
        "update_available": bool(tag) and version_tuple(tag) > version_tuple(APP_VERSION),
        "url": html_url,
    }
