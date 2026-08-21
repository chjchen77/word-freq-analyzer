#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd


def load_stata_variable_labels(path: Path) -> dict[str, str]:
    reader = pd.read_stata(path, iterator=True)
    return reader.variable_labels()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="将气候风险指标中的 *_ratio 改为以 MDA 句子数为分母重新计算。"
    )
    parser.add_argument("--risk-input", required=True, help="原始风险指标 .dta 文件")
    parser.add_argument("--sentence-input", required=True, help="MDA 句子统计面板 .csv 文件")
    parser.add_argument("--output-dta", required=True, help="输出 .dta 文件")
    parser.add_argument("--output-csv", help="可选：同步输出 .csv 文件")
    parser.add_argument("--issues-output", help="可选：输出无法计算句子分母的记录")
    return parser.parse_args()


def build_keys(df: pd.DataFrame, stock_col: str, year_col: str) -> pd.DataFrame:
    out = df.copy()
    out["stkcd_key"] = pd.to_numeric(out[stock_col], errors="coerce").astype("Int64").astype(str).str.zfill(6)
    out["year_key"] = pd.to_numeric(out[year_col], errors="coerce").astype("Int64").astype(str)
    return out


def load_sentence_panel(path: Path) -> pd.DataFrame:
    sent = pd.read_csv(path, encoding="utf-8-sig", dtype=str)
    sent = build_keys(sent, "公司代码", "年份")

    sent = sent.rename(
        columns={
            "公司代码": "mda_company_code",
            "年份": "mda_year",
            "公司名称": "mda_company_name",
            "提取状态": "mda_extract_status",
            "句子数": "mda_sentence_count",
            "原始句子数": "mda_raw_sentence_count",
            "疑似表格段落数": "mda_table_like_paragraphs",
        }
    )

    numeric_cols = ["mda_sentence_count", "mda_raw_sentence_count", "mda_table_like_paragraphs"]
    for col in numeric_cols:
        sent[col] = pd.to_numeric(sent[col], errors="coerce")

    return sent[
        [
            "stkcd_key",
            "year_key",
            "mda_company_name",
            "mda_extract_status",
            "mda_sentence_count",
            "mda_raw_sentence_count",
            "mda_table_like_paragraphs",
        ]
    ]


def add_missing_reason(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["mda_sentence_issue"] = ""

    unmatched = out["mda_extract_status"].isna()
    failed = out["mda_extract_status"].eq("failed")
    zero_or_negative = out["mda_sentence_count"].fillna(np.nan).le(0)
    missing_count = out["mda_sentence_count"].isna()

    out.loc[unmatched, "mda_sentence_issue"] = "missing_mda_panel"
    out.loc[failed, "mda_sentence_issue"] = "mda_extraction_failed"
    out.loc[~failed & ~unmatched & zero_or_negative, "mda_sentence_issue"] = "nonpositive_sentence_count"
    out.loc[~failed & ~unmatched & missing_count, "mda_sentence_issue"] = "missing_sentence_count"

    return out


def recalculate_ratios(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    out = df.copy()
    ratio_cols = [col for col in out.columns if col.endswith("_ratio")]

    denom = out["mda_sentence_count"].astype(float)
    valid = denom.gt(0)

    for ratio_col in ratio_cols:
        freq_col = ratio_col.replace("_ratio", "_freq")
        out[ratio_col] = np.where(valid, out[freq_col].astype(float) / denom, np.nan)

    # 额外补三列，方便直接使用“总/物理/转型”占 MDA 句子数的比例
    for freq_col, new_col in [
        ("total_freq", "total_sentence_ratio"),
        ("physical_freq", "physical_sentence_ratio"),
        ("transition_freq", "transition_sentence_ratio"),
    ]:
        out[new_col] = np.where(valid, out[freq_col].astype(float) / denom, np.nan)

    return out, ratio_cols


def build_variable_labels(
    original_labels: dict[str, str],
    columns: list[str],
    ratio_cols: list[str],
) -> dict[str, str]:
    labels = dict(original_labels)

    for ratio_col in ratio_cols:
        original_label = labels.get(ratio_col, "")
        if original_label:
            labels[ratio_col] = original_label.replace("占比", "占MDA句子数比重")

    labels.update(
        {
            "stkcd_key": "股票代码（六位字符串键）",
            "year_key": "年份（字符串键）",
            "mda_company_name": "公司名称（MDA文本）",
            "mda_extract_status": "MDA提取状态",
            "mda_sentence_count": "MDA句子数（清洗表格后）",
            "mda_raw_sentence_count": "MDA原始句子数（清洗前）",
            "mda_table_like_paragraphs": "MDA疑似表格段落数",
            "mda_sentence_issue": "MDA句子分母异常原因",
            "total_sentence_ratio": "气候风险关键词命中总句数占MDA句子数比重",
            "physical_sentence_ratio": "物理风险关键词命中句数占MDA句子数比重",
            "transition_sentence_ratio": "转型风险关键词命中句数占MDA句子数比重",
        }
    )

    return {col: labels.get(col, col) for col in columns}


def main() -> None:
    args = parse_args()

    risk_input = Path(args.risk_input).expanduser().resolve()
    sentence_input = Path(args.sentence_input).expanduser().resolve()
    output_dta = Path(args.output_dta).expanduser().resolve()
    output_csv = Path(args.output_csv).expanduser().resolve() if args.output_csv else None
    issues_output = Path(args.issues_output).expanduser().resolve() if args.issues_output else None

    original_labels = load_stata_variable_labels(risk_input)
    risk = pd.read_stata(risk_input)
    risk = build_keys(risk, "stkcd", "year")
    sentence_panel = load_sentence_panel(sentence_input)

    merged = risk.merge(sentence_panel, on=["stkcd_key", "year_key"], how="left")
    merged = add_missing_reason(merged)
    merged, ratio_cols = recalculate_ratios(merged)
    variable_labels = build_variable_labels(original_labels, list(merged.columns), ratio_cols)

    output_dta.parent.mkdir(parents=True, exist_ok=True)
    merged.to_stata(
        output_dta,
        write_index=False,
        version=118,
        variable_labels=variable_labels,
    )

    if output_csv is not None:
        merged.to_csv(output_csv, index=False, encoding="utf-8-sig")

    if issues_output is not None:
        issues = merged[merged["mda_sentence_issue"] != ""].copy()
        issues.to_csv(issues_output, index=False, encoding="utf-8-sig")

    print(f"输入行数：{len(merged)}")
    print(f"重算 ratio 列数：{len(ratio_cols)}")
    print(f"成功拿到有效句子分母的行数：{int(merged['mda_sentence_count'].fillna(0).gt(0).sum())}")
    print(f"句子分母缺失/不可用的行数：{int((merged['mda_sentence_issue'] != '').sum())}")
    print(f"DTA 输出：{output_dta}")
    if output_csv is not None:
        print(f"CSV 输出：{output_csv}")
    if issues_output is not None:
        print(f"问题清单：{issues_output}")


if __name__ == "__main__":
    main()
