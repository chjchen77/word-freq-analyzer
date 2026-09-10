"""Data-quality reporting and reproducibility metadata."""

from __future__ import annotations

import hashlib
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

import pandas as pd

QUALITY_NUMBER_FIELDS = (
    "原始行数",
    "有效记录数",
    "进入统计记录数",
    "公司代码为空",
    "年份无法识别",
    "月份无法识别",
    "文本为空",
    "关键词命中次数",
)


def empty_quality_stats(source_file: str, status: str = "成功") -> dict[str, Any]:
    row: dict[str, Any] = {
        "文件": source_file,
        "状态": status,
        "工作表": "",
        "表头行": "",
        "表头识别置信度": "",
    }
    row.update({field: 0 for field in QUALITY_NUMBER_FIELDS})
    return row


def merge_quality_stats(target: dict[str, Any], update: dict[str, Any]) -> dict[str, Any]:
    for field in QUALITY_NUMBER_FIELDS:
        target[field] = int(target.get(field, 0) or 0) + int(update.get(field, 0) or 0)
    if update.get("状态") and update.get("状态") != "成功":
        target["状态"] = update["状态"]
    for field in ("工作表", "表头行", "表头识别置信度"):
        if update.get(field) not in (None, ""):
            target[field] = update[field]
    return target


def quality_report_dataframe(rows: Iterable[dict[str, Any]]) -> pd.DataFrame:
    records = [dict(row) for row in rows]
    for row in records:
        raw = int(row.get("原始行数", 0) or 0)
        valid = int(row.get("有效记录数", 0) or 0)
        known_month = max(0, valid - int(row.get("月份无法识别", 0) or 0))
        row["有效记录率(%)"] = round(valid / raw * 100, 2) if raw else 0.0
        row["月份识别率(%)"] = round(known_month / valid * 100, 2) if valid else 0.0

    if records:
        total = empty_quality_stats("合计")
        for row in records:
            merge_quality_stats(total, row)
        total["状态"] = (
            "完成（含异常）"
            if any(str(row.get("状态", "成功")) != "成功" for row in records)
            else "完成"
        )
        raw = int(total["原始行数"])
        valid = int(total["有效记录数"])
        known_month = max(0, valid - int(total["月份无法识别"]))
        total["有效记录率(%)"] = round(valid / raw * 100, 2) if raw else 0.0
        total["月份识别率(%)"] = round(known_month / valid * 100, 2) if valid else 0.0
        records.append(total)

    columns = [
        "文件", "状态", "工作表", "表头行", "表头识别置信度", *QUALITY_NUMBER_FIELDS,
        "有效记录率(%)", "月份识别率(%)",
    ]
    return pd.DataFrame(records, columns=columns)


def _fingerprint_file(path: str, full_hash_limit: int = 200 * 1024 * 1024) -> tuple[str, str]:
    """Return an auditable hash without rereading multi-gigabyte inputs in full."""
    size = os.path.getsize(path)
    digest = hashlib.sha256()
    if size <= full_hash_limit:
        method = "SHA256"
        with open(path, "rb") as handle:
            for block in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(block)
    else:
        method = "SHA256(首尾各1MB+文件大小)"
        with open(path, "rb") as handle:
            digest.update(handle.read(1024 * 1024))
            handle.seek(max(0, size - 1024 * 1024))
            digest.update(handle.read(1024 * 1024))
        digest.update(str(size).encode("ascii"))
    return method, digest.hexdigest()


def input_manifest_dataframe(files: Iterable[str]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for raw_path in files:
        path = os.path.abspath(raw_path)
        try:
            stat = os.stat(path)
            method, fingerprint = _fingerprint_file(path)
            rows.append({
                "输入文件": path,
                "文件大小(字节)": stat.st_size,
                "修改时间(UTC)": datetime.fromtimestamp(
                    stat.st_mtime, tz=timezone.utc
                ).isoformat(timespec="seconds"),
                "指纹算法": method,
                "文件指纹": fingerprint,
            })
        except OSError as exc:
            rows.append({
                "输入文件": path,
                "文件大小(字节)": "",
                "修改时间(UTC)": "",
                "指纹算法": "",
                "文件指纹": f"读取失败：{exc}",
            })
    return pd.DataFrame(rows, columns=[
        "输入文件", "文件大小(字节)", "修改时间(UTC)", "指纹算法", "文件指纹",
    ])


def dictionary_sha256(dictionary: dict[str, list[str]]) -> str:
    digest = hashlib.sha256()
    for category in sorted(dictionary):
        digest.update(category.encode("utf-8"))
        digest.update(b"\0")
        for word in sorted(dictionary[category]):
            digest.update(word.encode("utf-8"))
            digest.update(b"\0")
    return digest.hexdigest()


def metadata_dataframe(items: dict[str, Any]) -> pd.DataFrame:
    return pd.DataFrame(
        [{"项目": key, "值": value} for key, value in items.items()],
        columns=["项目", "值"],
    )
