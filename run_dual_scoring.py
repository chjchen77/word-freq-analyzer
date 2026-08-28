#!/usr/bin/env python3
"""全量双轨评分：对命中句逐句打 phys/trans 两个 0-4 分。

面向长时间无人值守运行（全量约 5-6 天）设计：

- **断点续跑**：结果写入 SQLite 缓存，重启后已完成的句子秒级跳过，
  只补未完成部分。中断、重启、换机器都不必从头再来。
- **分片交付**：按年份切片，每跑完一年落一次盘，中途即可取用已完成年份。
- **进度可见**：每片打印完成量与预计剩余时间。

典型用法（服务器上）：

    export DASHSCOPE_API_KEY=sk-xxx
    screen -S bio                       # 断开 SSH 也不中断
    python3 run_dual_scoring.py \\
        --sentences 命中句.xlsx \\
        --out 打分结果.xlsx \\
        --cache dual_cache.db
    # Ctrl+A D 脱离；screen -r bio 回来看进度

缓存文件（--cache）是最重要的资产：跑了几天的标注都在里面，
务必与结果文件一同备份；只要它还在，任何中断都能秒级恢复。
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from llm_sentence_analyzer import LLMAnalyzerConfig, QwenSentenceAnalyzer  # noqa: E402

DEFAULT_RUBRIC = Path(__file__).resolve().parent / "rubrics" / "生物多样性_双轨评分标准.md"


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="全量双轨评分（支持断点续跑）")
    p.add_argument("--sentences", required=True, help="命中句文件（xlsx/csv）")
    p.add_argument("--out", required=True, help="输出文件（xlsx）")
    p.add_argument("--cache", default="dual_cache.db", help="SQLite 缓存路径（断点续跑依赖）")
    p.add_argument("--rubric", default=str(DEFAULT_RUBRIC), help="评分标准 Markdown")
    p.add_argument("--model", default="", help="模型名，留空用内置默认")
    p.add_argument("--workers", type=int, default=8, help="并发线程数")
    p.add_argument("--col-year", default="年份", help="年份列名")
    p.add_argument("--col-keyword", default="命中关键词", help="关键词列名")
    p.add_argument("--col-category", default="分类", help="关键词分类列名")
    p.add_argument("--col-sentence", default="命中句子", help="句子列名")
    return p.parse_args()


def load_sentences(path: str) -> pd.DataFrame:
    """股票代码固定为字符串，避免 "000001" 被推断成整数而丢掉前导零。"""
    kw = {"dtype": {"公司代码": str}}
    return (pd.read_csv(path, **kw) if path.lower().endswith(".csv")
            else pd.read_excel(path, **kw))


def main() -> None:
    args = parse_args()
    if not os.environ.get("DASHSCOPE_API_KEY"):
        sys.exit("缺少环境变量 DASHSCOPE_API_KEY")

    df = load_sentences(args.sentences)
    for col in (args.col_keyword, args.col_sentence):
        if col not in df.columns:
            sys.exit(f"输入文件缺少必需列：{col}（现有列：{list(df.columns)}）")
    if args.col_category not in df.columns:
        df[args.col_category] = ""

    # 按年份分片：每片独立落盘，中途可取用已完成年份
    if args.col_year in df.columns:
        groups = [(str(y), g) for y, g in df.groupby(args.col_year, sort=True)]
    else:
        groups = [("全部", df)]

    analyzer = QwenSentenceAnalyzer(
        LLMAnalyzerConfig.from_inputs(
            api_key=os.environ["DASHSCOPE_API_KEY"],
            model=args.model,
            rubric_path=args.rubric,
            cache_path=args.cache,
            max_workers=args.workers,
            max_sentences=0,          # 0 = 无上限
        ),
        log_cb=lambda m: print("    " + m[:120], flush=True),
    )
    if not analyzer.dual_scoring:
        sys.exit(f"未能加载评分标准：{args.rubric}")

    ckpt_path = args.out + ".partial.csv"
    print(f"共 {len(df):,} 句，分 {len(groups)} 片；缓存 {args.cache}", flush=True)
    print(f"模型 {analyzer.config.model}，并发 {args.workers}", flush=True)
    print(f"检查点 {ckpt_path}（每片更新，原子替换）\n", flush=True)

    done_parts: list[pd.DataFrame] = []
    t_start = time.time()
    finished = 0

    for i, (label, part) in enumerate(groups, 1):
        t0 = time.time()
        recs = [{"命中关键词": r[args.col_keyword],
                 "分类": r[args.col_category],
                 "命中句子": r[args.col_sentence]}
                for _, r in part.iterrows()]
        out = analyzer.analyze_records(recs)

        part = part.copy()
        part["LLM物理分数"] = [o.get("LLM物理分数") for o in out]
        part["LLM转型分数"] = [o.get("LLM转型分数") for o in out]
        # 语气语调用于面板的「语调积极/消极/中性」分组计数，必须落盘，
        # 否则模型算了、也计了费，结果却被丢弃。
        part["LLM语气语调"] = [o.get("LLM语气语调") for o in out]
        part["LLM置信度"] = [o.get("LLM置信度") for o in out]
        part["LLM分析状态"] = [o.get("LLM分析状态") for o in out]
        part["LLM分析错误"] = [o.get("LLM分析错误") for o in out]
        done_parts.append(part)

        # 每片结束落一次检查点。用 CSV 而非 Excel：同样 12 万行，Excel 写一次
        # 需 44 秒、CSV 仅 1 秒，25 片累计可省约 18 分钟；更要紧的是那 44 秒里
        # 若进程被杀，Excel 文件会残缺不可读。
        # 先写临时文件再 os.replace 原子替换，确保检查点任何时刻都是完整的。
        tmp = ckpt_path + ".tmp"
        pd.concat(done_parts, ignore_index=True).to_csv(tmp, index=False)
        os.replace(tmp, ckpt_path)

        finished += len(part)
        ok = (part["LLM分析状态"] == "成功").sum()
        el = time.time() - t0
        rate = finished / max(time.time() - t_start, 1e-9)
        eta = (len(df) - finished) / rate / 3600 if rate else float("nan")
        print(f"[{i}/{len(groups)}] {label}：{ok}/{len(part)} 成功，"
              f"耗时 {el / 60:.1f} 分钟｜累计 {finished:,}/{len(df):,}，"
              f"预计剩余 {eta:.1f} 小时\n", flush=True)

    final = pd.concat(done_parts, ignore_index=True)
    # Excel 只在全部跑完后写一次；同样先写临时文件再原子替换
    tmp_xlsx = args.out + ".tmp.xlsx"
    final.to_excel(tmp_xlsx, index=False)
    os.replace(tmp_xlsx, args.out)
    ok = (final["LLM分析状态"] == "成功").sum()
    print(f"全部完成：{ok:,}/{len(final):,} 成功，"
          f"总耗时 {(time.time() - t_start) / 3600:.1f} 小时")
    if ok < len(final):
        print(f"注意：{len(final) - ok:,} 条未成功。直接重跑本命令即可，"
              f"已完成的句子会命中缓存跳过，只补失败部分。")

    for dim, lbl in [("LLM物理分数", "物理"), ("LLM转型分数", "转型")]:
        v = pd.to_numeric(final[dim], errors="coerce")
        print(f"  {lbl}：均值 {v.mean():.2f}，>0 占比 {(v > 0).mean() * 100:.1f}%，"
              f"分布 {'/'.join(str(int((v == i).sum())) for i in range(5))}")
    print(f"\n结果已写出：{args.out}")


if __name__ == "__main__":
    main()
