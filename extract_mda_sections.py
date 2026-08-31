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

PRIMARY_TITLES = (
    "董事会报告及管理层讨论与分析",
    "董事會報告及管理層討論與分析",
    "管理层讨论与分析",
    "管理層討論與分析",
    "管理者讨论与分析",
    "管理者討論與分析",
    "经理层讨论与分析",
    "经营情况讨论与分析",
    "經營情況討論與分析",
    "董事会工作报告",
    "董事會工作報告",
    "董事局工作报告",
    "董事局工作報告",
    "董事会报告",
    "董事會報告",
    "董事局报告",
    "董事局報告",
)

FALLBACK_TITLES = (
    "经营情况回顾与分析",
    "經營情況回顧與分析",
    "经营情况分析",
    "經營情況分析",
    "业务回顾与展望",
    "業務回顧與展望",
)

ALL_TITLES = tuple(
    sorted(PRIMARY_TITLES + FALLBACK_TITLES, key=len, reverse=True)
)
TITLE_PATTERN = "|".join(re.escape(title) for title in ALL_TITLES)
PRIMARY_TITLE_SET = set(PRIMARY_TITLES)
FALLBACK_TITLE_SET = set(FALLBACK_TITLES)
BODY_HINTS = (
    "报告期内",
    "主要业务",
    "主营业务",
    "经营情况",
    "财务状况",
    "风险管理",
    "核心竞争力",
    "未来展望",
    "业务回顾",
    "公司治理",
)
PAGE_NUM_RE = re.compile(r"^[0-9Ｏ○oO]+$")
LEADER_DOTS_RE = re.compile(r"[\.·•…]{3,}")
TOC_PAGE_RE = re.compile(r"[\.·•…]{2,}\s*[0-9Ｏ○oO]{1,4}\s*$")

SECTION_TITLE_RE = re.compile(
    rf"(?m)^\s*第\s*[一二三四五六七八九十百0-9]+\s*[章节節]\s*[\u3000 \t\r\f\v\n]*({TITLE_PATTERN})"
)
ALT_START_RE = re.compile(
    rf"(?m)^\s*(?:第\s*[一二三四五六七八九十百0-9]+\s*[章节節]\s*)?({TITLE_PATTERN})\s*$"
)
NEXT_SECTION_RE = re.compile(r"(?m)^\s*第\s*[一二三四五六七八九十百0-9]+\s*[章节節](?:[^\n]*)")
SUBHEADING_RE = re.compile(
    r"(?m)"
    r"(^\s*[一二三四五六七八九十]\s*[、\.．])|"
    r"(^\s*[（(]\s*一\s*[)）])|"
    r"(^\s*\d+\s*[、\.．])"
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="批量提取年报中的 MDA（管理层讨论与分析）部分。")
    parser.add_argument("--src-root", required=True, help="源年报根目录，例如 /Users/chj/Desktop/Work/年报/解压")
    parser.add_argument("--out-root", required=True, help="输出目录，例如 /Users/chj/Desktop/Work/年报/MDA")
    parser.add_argument("--limit", type=int, default=0, help="仅处理前 N 份文件，0 表示处理全部")
    return parser.parse_args()


def score_text(text: str) -> int:
    sample = text[:250000]
    chinese = len(CHINESE_RE.findall(sample))
    keywords = 0
    for kw in ("管理层讨论与分析", "经营情况讨论与分析", "年度报告", "目录", "公司代码"):
        keywords += sample.count(kw) * 200
    replacement_penalty = sample.count("\ufffd") * 20
    return chinese + keywords - replacement_penalty


def read_best_text(path: Path) -> tuple[str, str]:
    raw = path.read_bytes()
    best_text = ""
    best_encoding = ""
    best_score = -10**18

    for enc in ENCODINGS:
        try:
            text = raw.decode(enc)
        except UnicodeDecodeError:
            continue
        text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\x00", "")
        current_score = score_text(text)
        if current_score > best_score:
            best_text = text
            best_encoding = enc
            best_score = current_score

    if not best_encoding:
        raise UnicodeDecodeError("unknown", b"", 0, 1, f"无法识别编码：{path}")

    return best_text, best_encoding


def normalize_inline_text(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text)


def is_probable_toc_line(line: str) -> bool:
    compact = normalize_inline_text(line)
    if not compact:
        return False
    if "目录" in compact or "目錄" in compact:
        return True
    if TOC_PAGE_RE.search(compact):
        return True
    return bool(LEADER_DOTS_RE.search(compact))


def is_probable_toc_excerpt(extracted: str) -> bool:
    lines = [line.strip() for line in extracted.splitlines() if line.strip()][:8]
    if not lines:
        return True

    if any("目录" in normalize_inline_text(line) or "目錄" in normalize_inline_text(line) for line in lines):
        return True

    digit_only = sum(bool(PAGE_NUM_RE.fullmatch(normalize_inline_text(line))) for line in lines)
    leader_lines = sum(bool(LEADER_DOTS_RE.search(normalize_inline_text(line))) for line in lines)
    punct = sum(line.count(ch) for line in lines for ch in "。！？!?；;")
    chinese = len(CHINESE_RE.findall("".join(lines)))

    if leader_lines > 0:
        return True
    if punct == 0 and digit_only >= 2 and chinese < 80:
        return True
    if punct == 0 and chinese < 50:
        return True
    return False


def iter_line_candidates(text: str) -> list[tuple[int, str]]:
    candidates: list[tuple[int, str]] = []
    offset = 0

    for raw_line in text.splitlines(keepends=True):
        line = raw_line.rstrip("\n")
        compact = normalize_inline_text(line)
        if compact:
            if not is_probable_toc_line(line):
                if not any(token in compact for token in ("详见", "參見", "请参阅", "請參閱", "请查阅", "參閱")):
                    for title in ALL_TITLES:
                        idx = compact.find(title)
                        if idx < 0:
                            continue
                        prefix = compact[:idx]
                        suffix = compact[idx + len(title):]
                        if len(prefix) <= 12 and len(suffix) <= 20:
                            if any(token in prefix for token in ("报告", "報告", "公司在", "详见", "詳見", "参见", "參見")):
                                continue
                            if suffix and PAGE_NUM_RE.fullmatch(suffix):
                                continue
                            candidates.append((offset, title))
                            break
        offset += len(raw_line)

    return candidates


def parse_filename_metadata(path: Path) -> dict[str, str]:
    stem = path.stem
    parts = stem.split("_")

    code_match = CODE_RE.search(stem)
    year_match = YEAR_RE.search(stem)
    code = code_match.group(1) if code_match else (parts[0] if parts else "")
    year = year_match.group(1) if year_match else ""

    company_name = ""
    report_title = ""
    report_date = ""

    # 兼容两类常见文件名：
    # 1) 600276_2022_恒瑞医药_恒瑞医药2022年年度报告_2023-04-21
    # 2) 000001_平安银行_2025
    if len(parts) >= 3 and re.fullmatch(r"(?:19|20)\d{2}", parts[1] or ""):
        company_name = parts[2]
        if len(parts) >= 4:
            report_date = parts[-1]
            report_title = "_".join(parts[3:-1]) if len(parts) > 4 else parts[3]
    elif len(parts) >= 3 and re.fullmatch(r"(?:19|20)\d{2}", parts[-1] or ""):
        company_name = "_".join(parts[1:-1]).strip()
        report_title = f"{company_name}{parts[-1]}年年度报告" if company_name else f"{parts[-1]}年年度报告"
    else:
        if len(parts) >= 3:
            company_name = parts[2]
        if len(parts) >= 4:
            report_date = parts[-1]
            report_title = "_".join(parts[3:-1]) if len(parts) > 4 else parts[3]

    return {
        "公司代码": code,
        "年份": year,
        "公司名称": company_name,
        "报告标题": report_title,
        "报告日期": report_date,
    }


def compute_candidate_end(text: str, start: int) -> int:
    end = len(text)
    for match in NEXT_SECTION_RE.finditer(text, start + 20):
        pos = match.start()
        if pos <= start + 20:
            continue
        end = pos
        break
    return end


def score_candidate(text: str, start: int, end: int, title: str, base_score: int) -> int:
    extracted = text[start:end].strip()
    preview = extracted[:4000]
    score = base_score

    chinese = len(CHINESE_RE.findall(preview))
    punct = sum(preview.count(ch) for ch in "。！？!?；;")
    subheads = sum(1 for _ in SUBHEADING_RE.finditer(preview[:2500]))

    score += min(chinese // 150, 25)
    score += min(punct, 20) * 3
    score += min(subheads, 12) * 3

    if title in PRIMARY_TITLE_SET:
        score += 8
    if title in FALLBACK_TITLE_SET:
        score -= 6
    if "管理层" in title or "管理層" in title or "经理层" in title:
        score += 6
    if "经营情况讨论与分析" in title or "經營情況討論與分析" in title:
        score += 5
    if "董事会报告及管理层讨论与分析" in title or "董事會報告及管理層討論與分析" in title:
        score += 6
    elif "董事会" in title or "董事會" in title or "董事局" in title:
        score += 2

    if start > 3000:
        score += 4
    if len(extracted) >= 500:
        score += 12
    elif len(extracted) >= 200:
        score += 8
    elif len(extracted) >= 120:
        score += 3
    else:
        score -= 12

    if any(hint in preview for hint in BODY_HINTS):
        score += 6
    if is_probable_toc_excerpt(extracted):
        score -= 35

    return score


def find_mda_bounds(text: str) -> tuple[int, int] | None:
    raw_candidates: list[tuple[int, str, int]] = []

    for match in SECTION_TITLE_RE.finditer(text):
        raw_candidates.append((match.start(), match.group(1), 18))

    for match in ALT_START_RE.finditer(text):
        raw_candidates.append((match.start(), match.group(1), 10))

    for pos, title in iter_line_candidates(text):
        raw_candidates.append((pos, title, 12 if title in PRIMARY_TITLE_SET else 5))

    if not raw_candidates:
        return None

    dedup: dict[int, tuple[str, int]] = {}
    for pos, title, base_score in raw_candidates:
        prev = dedup.get(pos)
        if prev is None or base_score > prev[1]:
            dedup[pos] = (title, base_score)

    evaluated: list[tuple[int, int, int]] = []
    for pos, (title, base_score) in dedup.items():
        end = compute_candidate_end(text, pos)
        if end <= pos:
            continue
        score = score_candidate(text, pos, end, title, base_score)
        evaluated.append((score, pos, end))

    if not evaluated:
        return None

    evaluated.sort(key=lambda item: (-item[0], item[1]))
    best_score, start, end = evaluated[0]
    if best_score < 8:
        return None
    return start, end


def build_output_text(metadata: dict[str, str], source_rel: str, encoding: str, status: str, message: str, mda_text: str) -> str:
    lines = [
        f"公司代码：{metadata.get('公司代码', '')}",
        f"年份：{metadata.get('年份', '')}",
        f"公司名称：{metadata.get('公司名称', '')}",
        f"报告标题：{metadata.get('报告标题', '')}",
        f"报告日期：{metadata.get('报告日期', '')}",
        f"来源文件：{source_rel}",
        f"源文件编码：{encoding}",
        f"提取状态：{status}",
        f"说明：{message}",
        "",
        "===== MDA（管理层讨论与分析）正文 =====",
        mda_text.strip(),
        "",
    ]
    return "\n".join(lines)


def process_one(path: Path, src_root: Path, out_root: Path) -> dict[str, str | int]:
    metadata = parse_filename_metadata(path)
    rel_path = path.relative_to(src_root)
    out_path = out_root / rel_path
    out_path.parent.mkdir(parents=True, exist_ok=True)

    status = "success"
    message = "提取成功"
    encoding = ""
    extracted = ""

    try:
        text, encoding = read_best_text(path)
        bounds = find_mda_bounds(text)
        if not bounds:
            status = "failed"
            message = "未识别到“管理层讨论与分析”章节边界"
        else:
            start, end = bounds
            extracted = text[start:end].strip()
            if len(extracted) < 120:
                status = "failed"
                message = "识别到的 MDA 文本过短，疑似误匹配"
    except Exception as exc:
        status = "failed"
        message = str(exc)

    out_text = build_output_text(
        metadata=metadata,
        source_rel=str(rel_path),
        encoding=encoding,
        status=status,
        message=message,
        mda_text=extracted,
    )
    out_path.write_text(out_text, encoding="utf-8")

    row: dict[str, str | int] = {
        "公司代码": metadata.get("公司代码", ""),
        "年份": metadata.get("年份", ""),
        "公司名称": metadata.get("公司名称", ""),
        "报告标题": metadata.get("报告标题", ""),
        "报告日期": metadata.get("报告日期", ""),
        "source_file": str(rel_path),
        "output_file": str(out_path.relative_to(out_root)),
        "source_encoding": encoding,
        "status": status,
        "message": message,
        "mda_chars": len(extracted),
    }
    return row


def main():
    args = parse_args()
    src_root = Path(args.src_root).expanduser().resolve()
    out_root = Path(args.out_root).expanduser().resolve()

    if not src_root.is_dir():
        raise SystemExit(f"源目录不存在：{src_root}")
    if out_root == src_root:
        raise SystemExit("输出目录不能与源目录相同，否则会覆盖原始 txt 文件。")

    def _is_inside(path: Path, parent: Path) -> bool:
        try:
            path.resolve().relative_to(parent)
            return True
        except ValueError:
            return False

    files = sorted(
        p for p in src_root.rglob("*.txt")
        if p.is_file() and not p.name.startswith(".")
        and not _is_inside(p, out_root)
    )
    if args.limit > 0:
        files = files[:args.limit]

    if not files:
        raise SystemExit(f"源目录下未找到 txt 文件：{src_root}")

    out_root.mkdir(parents=True, exist_ok=True)
    rows: list[dict[str, str | int]] = []

    total = len(files)
    for idx, path in enumerate(files, 1):
        rows.append(process_one(path, src_root, out_root))
        if idx == total or idx % 500 == 0:
            success = sum(1 for row in rows if row["status"] == "success")
            print(f"[{idx}/{total}] processed, success={success}, failed={idx - success}")

    index_path = out_root / "mda_index.csv"
    with index_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "公司代码", "年份", "公司名称", "报告标题", "报告日期",
                "source_file", "output_file", "source_encoding",
                "status", "message", "mda_chars",
            ],
        )
        writer.writeheader()
        writer.writerows(rows)

    success = sum(1 for row in rows if row["status"] == "success")
    failed = total - success
    print(f"Done. total={total}, success={success}, failed={failed}, index={index_path}")


if __name__ == "__main__":
    main()
