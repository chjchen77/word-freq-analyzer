"""Single source of truth for release metadata."""

APP_NAME = "中文文本词频统计分析工具"
APP_VERSION = "4.2.0"
GITHUB_REPOSITORY = "chjchen77/word-freq-analyzer"


def version_tuple(value: str) -> tuple[int, ...]:
    """Convert tags such as ``v4.2.0`` to a safely comparable tuple."""
    cleaned = value.strip().lower().lstrip("v")
    numbers: list[int] = []
    for part in cleaned.split("."):
        digits = "".join(ch for ch in part if ch.isdigit())
        if not digits:
            break
        numbers.append(int(digits))
    return tuple(numbers or [0])
