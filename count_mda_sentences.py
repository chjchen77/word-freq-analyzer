#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import re
from pathlib import Path


ENCODINGS = (
    "utf-8",
    "utf-8-sig",
    "gb18030",
    "gbk",
    "big5",
    "utf-16",
    "utf-16le",
    "utf-16be",
)

CHINESE_RE = re.compile(r"[\u4e00-\u9fff]")
YEAR_RE = re.compile(r"(?<!\d)((?:19|20)\d{2})(?!\d)")
CODE_RE = re.compile(r"(?<!\d)(\d{6})(?!\d)")
FIELD_RE = re.compile(
    r"^(公司代码|年份|公司名称|报告标题|报告日期|来源文件|源文件编码|提取状态|说明)：(.*)$",
    re.M,
)
BODY_MARKER = "===== MDA（管理层讨论与分析）正文 ====="
END_PUNCT_RE = re.compile(r"[。！？!?；;]$")
TABLE_HINT_RE = re.compile(
    r"(单位[:：]|币种|附表|分行业|分产品|分地区|分销售模式|项目名称|药品名称|"
    r"研发管线|营业收入|营业成本|同比增减|期末余额|期初余额|前五名客户|"
    r"前五名供应商|按治疗领域|按细分行业)"
)
ANNUAL_REPORT_LINE_RE = re.compile(r"^(?:20\d{2}\s*年)?年度报告(?:全文)?(?:（.*?）)?$")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="统计提取后 MDA 文本的句子数，并整理为公司-年份面板数据。"
    )
    parser.add_argument("--src-root", required=True, help="MDA 文本根目录")
    parser.add_argument("--output", required=True, help="输出 CSV 路径")
    parser.add_argument("--excel-output", help="可选：同步输出 Excel 路径")
    parser.add_argument("--limit", type=int, default=0, help="仅处理前 N 份文件，0 表示全部")
    return parser.parse_args()


def score_text(text: str) -> int:
    sample = text[:200000]
    chinese = len(CHINESE_RE.findall(sample))
    keywords = 0
    for kw in ("管理层讨论与分析", "经营情况讨论与分析", "公司代码", "提取状态", "年度报告"):
        keywords += sample.count(kw) * 200
    replacement_penalty = sample.count("\ufffd") * 20
    return chinese + keywords - replacement_penalty


def normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\x00", "")


def read_best_text(path: Path) -> tuple[str, str]:
    raw = path.read_bytes()

    for encoding in ("utf-8", "utf-8-sig"):
        try:
            return normalize_newlines(raw.decode(encoding)), encoding
        except UnicodeDecodeError:
            pass

    best_text = ""
    best_encoding = ""
    best_score = -10**18

    for encoding in ENCODINGS:
        try:
            text = normalize_newlines(raw.decode(encoding))
        except UnicodeDecodeError:
            continue
        current_score = score_text(text)
        if current_score > best_score:
            best_text = text
            best_encoding = encoding
            best_score = current_score

    if not best_encoding:
        raise UnicodeDecodeError("unknown", b"", 0, 1, f"无法识别编码：{path}")

    return best_text, best_encoding


def parse_filename_metadata(path: Path) -> dict[str, str]:
    stem = path.stem
    parts = stem.split("_")

    code_match = CODE_RE.search(stem)
    year_match = YEAR_RE.search(stem)

    company_name = parts[2] if len(parts) >= 3 else ""
    report_date = parts[-1] if len(parts) >= 4 else ""
    report_title = "_".join(parts[3:-1]) if len(parts) > 4 else (parts[3] if len(parts) == 4 else "")

    return {
        "公司代码": code_match.group(1) if code_match else (parts[0] if parts else ""),
        "年份": year_match.group(1) if year_match else "",
        "公司名称": company_name,
        "报告标题": report_title,
        "报告日期": report_date,
    }


def parse_header_metadata(text: str) -> dict[str, str]:
    return {key: value.strip() for key, value in FIELD_RE.findall(text)}


def normalize_line(line: str) -> str:
    return " ".join(line.split())


def is_page_artifact_line(line: str) -> bool:
    return bool(
        re.fullmatch(r"\d+", line)
        or ANNUAL_REPORT_LINE_RE.fullmatch(line)
        or re.fullmatch(r"[-_=]{3,}", line)
    )


def is_ascii_heavy_line(line: str) -> bool:
    digits = sum(ch.isdigit() for ch in line)
    ascii_letters = sum(ch.isascii() and ch.isalpha() for ch in line)
    chinese = sum("\u4e00" <= ch <= "\u9fff" for ch in line)

    if re.fullmatch(r"[\d\W_]+", line):
        return True
    if digits >= 4 or ascii_letters >= 3:
        return True
    if len(line) <= 12 and chinese <= 8:
        return True
    return False


def is_table_like_paragraph(lines: list[str]) -> bool:
    punctuated_lines = sum(bool(END_PUNCT_RE.search(line)) for line in lines)
    if punctuated_lines > 0:
        return False

    avg_line_length = sum(len(line) for line in lines) / len(lines)
    ascii_heavy_lines = sum(is_ascii_heavy_line(line) for line in lines)
    table_hint_lines = sum(bool(TABLE_HINT_RE.search(line)) for line in lines)

    if table_hint_lines >= 1:
        return True
    if len(lines) == 1:
        return is_ascii_heavy_line(lines[0]) or len(lines[0]) <= 20
    if len(lines) == 2:
        return avg_line_length <= 18 or ascii_heavy_lines >= 1
    return avg_line_length <= 24 or (ascii_heavy_lines / len(lines)) >= 0.2


def remove_table_like_paragraphs(text: str) -> tuple[str, int]:
    paragraphs = re.split(r"\n\s*\n+", normalize_newlines(text))
    kept_paragraphs: list[str] = []
    dropped_count = 0

    for paragraph in paragraphs:
        lines = [
            normalized
            for raw_line in paragraph.split("\n")
            if (normalized := normalize_line(raw_line)) and not is_page_artifact_line(normalized)
        ]
        if not lines:
            continue
        if is_table_like_paragraph(lines):
            dropped_count += 1
            continue
        kept_paragraphs.append("\n".join(lines))

    return "\n\n".join(kept_paragraphs), dropped_count


# 分句逻辑统一引用 sentence_split（与 word_freq_analyzer 同源），
# 保证分母(MD&A句数)与分子(命中句)口径完全一致。
# compact=True 仅改变返回形态，不影响句子数量。
from sentence_split import is_table_like, split_sentences as _split  # noqa: E402,F401


def split_sentences(text: str) -> list[str]:
    return _split(text, compact=True)


def count_body_chars(text: str) -> int:
    return len(re.sub(r"\s+", "", normalize_newlines(text)))


def build_row(path: Path, src_root: Path) -> dict[str, str | int]:
    text, detected_encoding = read_best_text(path)
    header_text, marker, body = text.partition(BODY_MARKER)
    header_meta = parse_header_metadata(header_text)
    fallback_meta = parse_filename_metadata(path)

    metadata = fallback_meta | header_meta
    metadata.setdefault("来源文件", str(path.relative_to(src_root)))
    metadata.setdefault("源文件编码", detected_encoding)

    status = metadata.get("提取状态", "")
    note = metadata.get("说明", "")

    sentence_count: int | str = ""
    raw_sentence_count: int | str = ""
    char_count: int | str = ""
    dropped_table_paragraphs: int | str = ""

    if status == "success" and marker:
        body = body.strip()
        cleaned_body, dropped_table_paragraphs = remove_table_like_paragraphs(body)
        raw_sentence_count = len(split_sentences(body))
        sentence_count = len(split_sentences(cleaned_body))
        char_count = count_body_chars(cleaned_body)
    elif status == "success" and not marker:
        status = "failed"
        note = "缺少 MDA 正文标记，无法统计句子数"

    return {
        "公司代码": metadata.get("公司代码", ""),
        "年份": metadata.get("年份", ""),
        "公司名称": metadata.get("公司名称", ""),
        "报告标题": metadata.get("报告标题", ""),
        "报告日期": metadata.get("报告日期", ""),
        "提取状态": status,
        "说明": note,
        "原始句子数": raw_sentence_count,
        "句子数": sentence_count,
        "正文字数": char_count,
        "疑似表格段落数": dropped_table_paragraphs,
        "来源文件": metadata.get("来源文件", ""),
        "源文件编码": metadata.get("源文件编码", detected_encoding),
    }


def write_csv(rows: list[dict[str, str | int]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "公司代码",
        "年份",
        "公司名称",
        "报告标题",
        "报告日期",
        "提取状态",
        "说明",
        "原始句子数",
        "句子数",
        "正文字数",
        "疑似表格段落数",
        "来源文件",
        "源文件编码",
    ]
    with output_path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def write_excel(rows: list[dict[str, str | int]], output_path: Path) -> None:
    try:
        import pandas as pd
    except ImportError as exc:
        raise RuntimeError("未安装 pandas，无法输出 Excel 文件") from exc

    output_path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)
    if "公司代码" in df.columns:
        df["公司代码"] = df["公司代码"].astype(str).str.zfill(6)
    df.to_excel(output_path, index=False)


def main() -> None:
    args = parse_args()

    src_root = Path(args.src_root).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    excel_output_path = Path(args.excel_output).expanduser().resolve() if args.excel_output else None

    files = sorted(src_root.rglob("*.txt"))
    if args.limit > 0:
        files = files[:args.limit]

    rows = [build_row(path, src_root) for path in files]
    rows.sort(key=lambda row: (str(row["年份"]), str(row["公司代码"]), str(row["来源文件"])))

    write_csv(rows, output_path)
    if excel_output_path is not None:
        write_excel(rows, excel_output_path)

    success_rows = sum(1 for row in rows if row["提取状态"] == "success")
    failed_rows = len(rows) - success_rows

    print(f"已处理文件：{len(rows)}")
    print(f"成功统计：{success_rows}")
    print(f"失败/未统计：{failed_rows}")
    print(f"CSV 输出：{output_path}")
    if excel_output_path is not None:
        print(f"Excel 输出：{excel_output_path}")


if __name__ == "__main__":
    main()
