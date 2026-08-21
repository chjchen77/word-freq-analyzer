#!/usr/bin/env python3
"""Extract one deterministic batch of MDA texts and write its index rows."""

from __future__ import annotations

import argparse
import csv
from pathlib import Path

from extract_mda_sections import process_one


FIELDNAMES = [
    "公司代码", "年份", "公司名称", "报告标题", "报告日期",
    "source_file", "output_file", "source_encoding",
    "status", "message", "mda_chars",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="分批提取年报 MDA 文本。")
    parser.add_argument("--src-root", required=True, help="源年报目录")
    parser.add_argument("--out-root", required=True, help="MDA 输出目录")
    parser.add_argument("--start", required=True, type=int, help="按文件排序后的起始位置，从 0 开始")
    parser.add_argument("--limit", required=True, type=int, help="本批最多处理的文件数")
    parser.add_argument("--index-output", required=True, help="本批索引 CSV 输出路径")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    src_root = Path(args.src_root).expanduser().resolve()
    out_root = Path(args.out_root).expanduser().resolve()
    index_output = Path(args.index_output).expanduser().resolve()

    if not src_root.is_dir():
        raise SystemExit(f"源目录不存在：{src_root}")
    if args.start < 0 or args.limit <= 0:
        raise SystemExit("start 必须大于等于 0，limit 必须大于 0")

    all_files = sorted(
        p for p in src_root.rglob("*.txt")
        if p.is_file() and not p.name.startswith(".")
    )
    files = all_files[args.start:args.start + args.limit]
    if not files:
        raise SystemExit("指定批次没有可处理的 txt 文件")

    out_root.mkdir(parents=True, exist_ok=True)
    rows = [process_one(path, src_root, out_root) for path in files]

    index_output.parent.mkdir(parents=True, exist_ok=True)
    with index_output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)

    success = sum(row["status"] == "success" for row in rows)
    print(
        f"Done. start={args.start}, total={len(rows)}, "
        f"success={success}, failed={len(rows) - success}, index={index_output}"
    )


if __name__ == "__main__":
    main()
