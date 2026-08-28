#!/usr/bin/env python3
"""由双轨打分结果生成「公司-年份」面板（Stata .dta + Excel + 变量说明）。

口径约定（经研究团队确认）：
- 相关性阈值统一为 **得分 >= 1**（不是 > 1）；
- 物理与转型**各做一套**关键词层变量，互不合并；
- 「出现次数」按关键词在句中出现的次数累加；
  「句子数量」同一句只计一次。

变量命名：关键词层一律 kwNNN_ 前缀（NNN 为词典序号，与词表对照见说明文档），
避免中文变量名在 Stata 中的兼容问题。
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd

TONE_NAME = {0: "neg", 1: "neu", 2: "pos"}
TONE_LABEL = {0: "消极", 1: "中性", 2: "积极"}


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="生成公司-年份面板")
    p.add_argument("--scores", required=True, help="双轨打分结果（xlsx/csv）")
    p.add_argument("--wordfreq-panel", required=True, help="词频面板（含 mda_sent/total_freq）")
    p.add_argument("--dict", required=True, help="词典 xlsx")
    p.add_argument("--outdir", required=True)
    p.add_argument("--stata-version", type=int, default=15)
    return p.parse_args()


def load(path: str) -> pd.DataFrame:
    """读取时必须把股票代码固定为字符串。

    pandas 会把 "000001" 推断成整数 1，前导零丢失后再也无法与 CSMAR 等
    外部数据库按代码匹配；而 600000 这类沪市代码看不出异常，问题极易漏检。
    """
    kw = {"dtype": {"公司代码": str}}
    return (pd.read_csv(path, **kw) if path.lower().endswith(".csv")
            else pd.read_excel(path, **kw))


def main() -> None:
    args = parse_args()
    out = Path(args.outdir)
    out.mkdir(parents=True, exist_ok=True)

    sc = load(args.scores)
    sc = sc[sc["LLM分析状态"] == "成功"].copy()
    for c in ("LLM物理分数", "LLM转型分数", "LLM语气语调"):
        sc[c] = pd.to_numeric(sc[c], errors="coerce")
    sc = sc.dropna(subset=["LLM物理分数", "LLM转型分数"])
    sc["LLM语气语调"] = sc["LLM语气语调"].fillna(1).astype(int)  # 缺失按中性
    sc["phys"] = sc["LLM物理分数"].astype(int)
    sc["trans"] = sc["LLM转型分数"].astype(int)
    sc["code"] = sc["公司代码"].astype(str).str.zfill(6)
    sc["year"] = sc["年份"].astype(int)

    # 同一句中关键词的出现次数（句子数量则每句计一次）
    sc["occ"] = [
        max(1, str(s).lower().count(str(k).lower()))
        for s, k in zip(sc["命中句子"], sc["命中关键词"])
    ]

    dic = pd.read_excel(args.dict, sheet_name="词典")[["term", "tier", "tier_name"]]
    dic["term"] = dic["term"].astype(str).str.strip()
    dic = dic[dic["term"] != ""].reset_index(drop=True)
    dic["var"] = [f"kw{i + 1:03d}" for i in range(len(dic))]
    term2var = dict(zip(dic["term"], dic["var"]))

    wf = load(args.wordfreq_panel)
    wf["code"] = wf["公司代码"].astype(str).str.zfill(6)
    wf["year"] = wf["年份"].astype(int)

    labels: dict[str, str] = {}
    panel = wf[["code", "year", "mda_sent", "total_freq", "total_ratio"]].copy()
    labels.update({
        "code": "股票代码（6位）", "year": "报告年份",
        "mda_sent": "该年度该公司 MD&A 文本句子总数",
        "total_freq": "词典190词在 MD&A 中出现的总次数（绝对值，与 v26 口径一致）",
        "total_ratio": "total_freq / mda_sent（与 v26 口径一致）",
    })

    g = sc.groupby(["code", "year"])

    # ---------- 公司-年度层 ----------
    firm = pd.DataFrame(index=g.size().index)
    firm["hit_sent"] = g["命中句子"].nunique()
    firm["phys_sum"] = g["phys"].sum()
    firm["trans_sum"] = g["trans"].sum()
    labels.update({
        "hit_sent": "词典190词命中的句子数（去重，不论得分）",
        "phys_sum": "物理实质性得分合计",
        "trans_sum": "转型实质性得分合计",
    })
    for dim, zh in (("phys", "物理"), ("trans", "转型")):
        rel = sc[sc[dim] >= 1].groupby(["code", "year"])["命中句子"].nunique()
        zero = sc[sc[dim] == 0].groupby(["code", "year"])["命中句子"].nunique()
        firm[f"{dim}_rel_sent"] = rel
        firm[f"{dim}_zero_sent"] = zero
        labels[f"{dim}_rel_sent"] = f"{zh}得分>=1 的句子数（去重）"
        labels[f"{dim}_zero_sent"] = f"{zh}得分=0 的句子数（去重）"
    firm = firm.fillna(0).astype(int).reset_index()

    # ---------- 关键词层 ----------
    sc["var"] = sc["命中关键词"].map(term2var)
    blocks: list[pd.DataFrame] = []

    def add(series: pd.Series, suffix: str, label_tpl: str) -> None:
        """把 (code,year,var) 的统计量摊平成 kwNNN_suffix 列。

        必须按全部 190 词补齐：某个词在本次语料中一次都没命中时，
        unstack 不会为它生成列，面板结构就会缺失（应为全 0 而非无此列）。
        """
        w = series.unstack("var").reindex(columns=all_vars)
        w.columns = [f"{v}_{suffix}" for v in all_vars]
        blocks.append(w)
        for v in w.columns:
            labels[v] = label_tpl.format(term=var2term[v.split("_")[0]])

    var2term = {v: t for t, v in term2var.items()}
    all_vars = list(dic["var"])   # 全部 190 词，含本次语料中从未命中的
    base = sc.groupby(["code", "year", "var"])

    add(base["occ"].sum(), "occ_all", "「{term}」出现总次数（不论得分）")
    add(base["命中句子"].nunique(), "sent_all", "「{term}」命中句子数（不论得分）")

    for dim, zh in (("phys", "物理"), ("trans", "转型")):
        sub = sc[sc[dim] >= 1].groupby(["code", "year", "var"])
        add(sub["occ"].sum(), f"{dim}_occ", f"「{{term}}」{zh}得分>=1 的出现次数")
        add(sub["命中句子"].nunique(), f"{dim}_sent", f"「{{term}}」{zh}得分>=1 的句子数")
        add(sc.groupby(["code", "year", "var"])[dim].sum(), f"{dim}_score",
            f"「{{term}}」命中句的{zh}得分合计")
        z = sc[sc[dim] == 0].groupby(["code", "year", "var"])
        add(z["命中句子"].nunique(), f"{dim}_zero_sent", f"「{{term}}」{zh}得分=0 的句子数")
        for tv, tn in TONE_NAME.items():
            t = sc[(sc[dim] >= 1) & (sc["LLM语气语调"] == tv)].groupby(["code", "year", "var"])
            add(t["occ"].sum(), f"{dim}_{tn}_occ",
                f"「{{term}}」{zh}得分>=1 且语调{TONE_LABEL[tv]} 的出现次数")
            add(t["命中句子"].nunique(), f"{dim}_{tn}_sent",
                f"「{{term}}」{zh}得分>=1 且语调{TONE_LABEL[tv]} 的句子数")

    # int32 足以容纳任何计数（上限 21 亿），较 int64 省一半内存：
    # 6.8 万行 × 3 千列在 int64 下约 1.6 GB，int32 约 0.8 GB。
    kw = pd.concat(blocks, axis=1).fillna(0).astype("int32").reset_index()

    panel = (panel.merge(firm, on=["code", "year"], how="left")
                  .merge(kw, on=["code", "year"], how="left"))
    num = [c for c in panel.columns if c not in ("code", "year", "total_ratio")]
    panel[num] = panel[num].fillna(0)
    for c in num:
        if c != "mda_sent":
            panel[c] = panel[c].astype("int32")
    # 股票代码必须始终保持 6 位字符串。合并与排序过程中可能被推断成整数，
    # 一旦丢掉前导零（000002 → 2），就无法与 CSMAR 等外部库按代码匹配。
    panel["code"] = panel["code"].astype(str).str.replace(r"\.0$", "", regex=True).str.zfill(6)
    panel = panel.sort_values(["code", "year"]).reset_index(drop=True)

    # ---------- 输出 ----------
    # 不写整表 Excel：6.8 万行 × 3 千列 = 2 亿单元格，openpyxl 需数小时且
    # 内存翻数倍（实测 12 万行 × 10 列已需 44 秒）。改以 .dta 为主、CSV 为备，
    # Excel 仅用于人可读的变量说明与词表对照。
    panel.to_csv(out / "面板_公司年份.csv", index=False)

    safe = {c: re.sub(r"\W", "_", c)[:32] for c in panel.columns}
    dup = [c for c, n in pd.Series(list(safe.values())).value_counts().items() if n > 1]
    if dup:
        raise SystemExit(f"变量名截断后重名，请缩短命名：{dup[:5]}")
    try:
        import pyreadstat
        pyreadstat.write_dta(
            panel.rename(columns=safe), str(out / "面板_公司年份.dta"),
            column_labels={safe[c]: labels.get(c, c)[:80] for c in panel.columns},
            version=args.stata_version)
    except ImportError:
        print("提示：未安装 pyreadstat，跳过 .dta 输出（pip install pyreadstat）")

    pd.DataFrame({"变量名": list(panel.columns),
                  "含义": [labels.get(c, "") for c in panel.columns]}
                 ).to_excel(out / "变量说明.xlsx", index=False)
    dic[["var", "term", "tier", "tier_name"]].to_excel(out / "词表对照.xlsx", index=False)

    n_kw_cols = len(blocks)   # 每个 add() 产出一组「190 词 × 1 指标」
    print(f"面板 {len(panel):,} 行 × {len(panel.columns):,} 列 → {out}")
    print(f"  公司层 {len(panel.columns) - len(dic) * n_kw_cols} 个变量，"
          f"关键词层 {len(dic)} 词 × {n_kw_cols} 个指标 = {len(dic) * n_kw_cols:,} 列")


if __name__ == "__main__":
    main()
