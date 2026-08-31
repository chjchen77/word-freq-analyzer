#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小样本测试：随机抽取 20 份年报，跑气候相关词典，验证词频和命中句导出。
"""
import random
import sys
import os
from pathlib import Path

# 把程序目录加入 path
sys.path.insert(0, str(Path(__file__).parent))
import word_freq_analyzer as m

# ── 配置 ──────────────────────────────────────────────────
ANNUAL_REPORT_DIR = os.getenv(
    "ANNUAL_REPORT_DIR",
    "/Users/chj/Desktop/Work/气候适应性政策/01_raw/年报/解压/2022_5203份",
)
OUTPUT_PATH = str(Path(__file__).parent / "test_output.xlsx")
SAMPLE_N = 20
RANDOM_SEED = 42
ENABLE_LLM = bool(os.getenv("DASHSCOPE_API_KEY", "").strip())
LLM_MODEL = os.getenv("QWEN_MODEL", "qwen-plus")
LLM_BASE_URL = os.getenv("QWEN_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1")
LLM_MAX_SENTENCES = 50
LLM_MAX_WORKERS = 4

# 气候适应性相关词典（可自行修改）
KEYWORD_DICT = {
    "气候风险": ["气候变化", "气候风险", "气候灾害", "气候异常", "极端天气", "洪涝", "干旱", "台风"],
    "低碳转型": ["碳排放", "碳中和", "碳达峰", "低碳", "减碳", "脱碳", "碳足迹", "碳交易"],
    "绿色发展": ["绿色发展", "绿色金融", "绿色债券", "ESG", "可持续发展", "环境保护", "生态文明"],
    "适应措施": ["气候适应", "应对气候", "节能减排", "清洁能源", "新能源", "可再生能源"],
}
# ─────────────────────────────────────────────────────────


def main():
    report_dir = Path(ANNUAL_REPORT_DIR)
    if not report_dir.exists():
        print(f"样本目录不存在：{report_dir}")
        print("可先设置环境变量 ANNUAL_REPORT_DIR 指向你的年报 txt 样本目录。")
        return

    # 1. 随机抽样
    all_files = list(report_dir.rglob("*.txt"))
    if not all_files:
        print(f"样本目录下未找到任何 .txt 文件：{report_dir}")
        print("请确认样本目录是否正确，或目录里是否已经解压出年报 txt 文件。")
        return

    random.seed(RANDOM_SEED)
    sample_files = [str(f) for f in random.sample(all_files, min(SAMPLE_N, len(all_files)))]
    print(f"抽取 {len(sample_files)} 份年报：")
    for f in sorted(sample_files):
        print(f"  {Path(f).name}")

    # 2. 构建词典
    dict_mgr = m.DictionaryManager()
    for cat, words in KEYWORD_DICT.items():
        dict_mgr.add_category(cat)
        for w in words:
            dict_mgr.add_word(cat, w)

    # 3. 运行分析
    print(f"\n正在分析，输出至：{OUTPUT_PATH}")
    if ENABLE_LLM:
        print(f"已启用 LLM 句子分析：model={LLM_MODEL}, max_sentences={LLM_MAX_SENTENCES}")
    else:
        print("未检测到 DASHSCOPE_API_KEY，本次仅提取命中句子，不调用 LLM。")

    log_lines = []
    def log(msg):
        print(f"  [LOG] {msg}")
        log_lines.append(msg)

    m.run_analysis(
        files=sample_files,
        dict_mgr=dict_mgr,
        col_stkcd="公司代码",
        col_year="年份",
        text_columns=["文本内容"],
        output_path=OUTPUT_PATH,
        use_regex=True,
        use_tf=False,
        export_sentences=True,
        analyze_llm=ENABLE_LLM,
        llm_model=LLM_MODEL,
        llm_base_url=LLM_BASE_URL,
        llm_max_sentences=LLM_MAX_SENTENCES,
        llm_max_workers=LLM_MAX_WORKERS,
        log_cb=log,
    )

    # 4. 读取结果并打印摘要
    import pandas as pd
    print("\n" + "="*60)
    sentences_path = str(Path(os.path.splitext(OUTPUT_PATH)[0] + "_sentences.xlsx"))
    print("【命中句子文件 — 前 30 条】")
    print("="*60)
    try:
        df = pd.read_excel(sentences_path, sheet_name="命中句子")
        print(f"共 {len(df)} 条命中句子\n")

        if "LLM分析状态" in df.columns:
            print("LLM 分析状态：")
            print(df["LLM分析状态"].value_counts(dropna=False).to_string())
            print()

        # 打印前 30 条明细
        pd.set_option("display.max_colwidth", 60)
        pd.set_option("display.width", 200)
        cols = [
            "公司代码", "年份", "分类", "命中关键词",
            "LLM时间指向", "LLM语态", "LLM句子类型", "LLM确定性", "LLM量化属性", "LLM语气语调",
            "命中句子",
        ]
        cols = [c for c in cols if c in df.columns]
        print(df[cols].head(30).to_string(index=False))

    except Exception as e:
        print(f"读取结果失败：{e}")


if __name__ == "__main__":
    main()
