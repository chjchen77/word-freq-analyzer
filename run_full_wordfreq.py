#!/usr/bin/env python3
"""全量词频统计：扫描 MD&A 文本，产出「公司-年份面板」与「命中句明细」。

命中句明细体量小（约 30MB），是后续 LLM 双轨评分的唯一输入——
无需把数 GB 的年报文本搬到打分机器上。

用法：
    python3 run_full_wordfreq.py \\
        --mda-root /path/to/MDA \\
        --dict /path/to/biodiversity_risk_dictionary_v26.xlsx \\
        --outdir /path/to/结果
"""
from __future__ import annotations

import argparse
import glob
import multiprocessing as mp
import re
import sys
import time
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sentence_split import split_sentences  # noqa: E402

BODY_MARKER = "===== MDA（管理层讨论与分析）正文 ====="
_TERMS: list[tuple[str, int, str]] = []
_PROBE: re.Pattern | None = None


def _init(terms: list[tuple[str, int, str]]) -> None:
    """worker 初始化：预编译整句预筛正则，避免对无命中句逐词搜索。"""
    global _TERMS, _PROBE
    _TERMS = terms
    _PROBE = re.compile("|".join(
        re.escape(t) for t, _, _ in sorted(terms, key=lambda x: -len(x[0]))))


def _one_file(fp: str) -> tuple[dict | None, list[dict]]:
    """返回 (该公司-年份的词频行, 命中句列表)。"""
    try:
        raw = Path(fp).read_text(encoding="utf-8", errors="ignore")
    except OSError:
        return None, []
    _, marker, body = raw.partition(BODY_MARKER)
    if not marker:
        return None, []

    stem = Path(fp).stem.split("_")
    code, year = stem[0], (stem[1] if len(stem) > 1 else "")
    if not re.fullmatch(r"\d{6}", code) or not re.fullmatch(r"(?:19|20)\d{2}", year):
        return None, []

    sents = split_sentences(body)
    freq: dict[str, int] = {}
    hits: list[dict] = []
    for sent in sents:
        if not _PROBE.search(sent):     # 绝大多数句子在此被跳过
            continue
        low = sent.lower()
        for term, tier, tname in _TERMS:
            n = low.count(term.lower())
            if n:
                freq[term] = freq.get(term, 0) + n
                hits.append({"公司代码": code, "年份": int(year),
                             "命中关键词": term, "分类": tname,
                             "命中句子": sent})
    row = {"公司代码": code, "年份": int(year), "mda_sent": len(sents), **freq}
    return row, hits


def main() -> None:
    ap = argparse.ArgumentParser(description="全量词频统计")
    ap.add_argument("--mda-root", required=True)
    ap.add_argument("--dict", required=True)
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--procs", type=int, default=max(1, mp.cpu_count() - 1))
    args = ap.parse_args()

    d = pd.read_excel(args.dict, sheet_name="词典")[["term", "tier", "tier_name"]].dropna(subset=["term"])
    d["term"] = d["term"].astype(str).str.strip()
    d = d[d["term"] != ""]
    terms = [(r.term, int(r.tier), r.tier_name) for r in d.itertuples()]
    tier_of = {t: tr for t, tr, _ in terms}

    files = sorted(glob.glob(f"{args.mda_root}/*/*.txt"))
    print(f"词典 {len(terms)} 词｜MD&A 文件 {len(files):,} 份｜进程 {args.procs}", flush=True)

    t0 = time.time()
    rows: list[dict] = []
    hits: list[dict] = []
    with mp.Pool(args.procs, initializer=_init, initargs=(terms,)) as pool:
        for i, (row, hs) in enumerate(pool.imap_unordered(_one_file, files, chunksize=64), 1):
            if row:
                rows.append(row)
                hits.extend(hs)
            if i % 5000 == 0:
                el = time.time() - t0
                print(f"  {i:,}/{len(files):,}  已用 {el / 60:.1f} 分钟，"
                      f"预计剩余 {(len(files) - i) / (i / el) / 60:.1f} 分钟", flush=True)

    out = Path(args.outdir)
    out.mkdir(parents=True, exist_ok=True)

    # ---- 命中句明细（LLM 打分的唯一输入）----
    hit_df = pd.DataFrame(hits).drop_duplicates(
        subset=["公司代码", "年份", "命中关键词", "命中句子"])
    hit_df.to_excel(out / "命中句明细.xlsx", index=False)

    # ---- 公司-年份面板 ----
    panel = pd.DataFrame(rows).fillna(0)
    kw_cols = [t for t, _, _ in terms if t in panel.columns]
    panel[kw_cols] = panel[kw_cols].astype(int)
    panel["total_freq"] = panel[kw_cols].sum(axis=1)
    # mda_sent 为 0 时不产生 inf：置空表示无法计算而非 0
    panel["total_ratio"] = (panel["total_freq"] / panel["mda_sent"]
                            .replace(0, pd.NA))
    for tr in sorted({t for t in tier_of.values()}):
        cols = [t for t in kw_cols if tier_of[t] == tr]
        panel[f"t{tr}_freq"] = panel[cols].sum(axis=1) if cols else 0
        panel[f"t{tr}_ratio"] = panel[f"t{tr}_freq"] / panel["mda_sent"].replace(0, pd.NA)
    panel = panel.sort_values(["公司代码", "年份"])
    panel.to_excel(out / "公司年份面板.xlsx", index=False)

    print(f"\n完成，耗时 {(time.time() - t0) / 60:.1f} 分钟")
    print(f"  公司-年份面板 {len(panel):,} 行 → {out / '公司年份面板.xlsx'}")
    print(f"  命中句明细   {len(hit_df):,} 行 → {out / '命中句明细.xlsx'}")
    print(f"  命中总次数 {int(panel['total_freq'].sum()):,}"
          f"｜有命中的公司-年份 {(panel['total_freq'] > 0).sum():,}")
    print(f"  MD&A 句子数中位 {panel['mda_sent'].median():.0f}")


if __name__ == "__main__":
    main()
