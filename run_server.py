#!/usr/bin/env python3
"""
中文文本词频统计分析工具 — 服务器命令行入口
================================================
无 GUI，适合在 Linux 服务器上通过 nohup 长时间挂机运行。

典型用法
--------
# 基础运行（仅词频统计）
python3 run_server.py \
    --data-dir /data/annual_reports \
    --dict /data/dict/keywords.xlsx \
    --output /data/output/result.xlsx

# 启用 LLM 句子分析（无上限，16 并发）
python3 run_server.py \
    --data-dir /data/annual_reports \
    --dict /data/dict/keywords.xlsx \
    --output /data/output/result.xlsx \
    --llm-api-key sk-xxxx \
    --llm-model qwen-turbo \
    --llm-max-sentences 0 \
    --llm-workers 16 \
    --log-file /data/output/run.log

# 后台挂机（不挂断）
nohup python3 run_server.py [参数...] > /data/output/nohup.log 2>&1 &
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import threading
from pathlib import Path

# ── 确保脚本所在目录在 Python 路径中 ────────────────────────────────────────
_SCRIPT_DIR = Path(__file__).parent.resolve()
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))


def _setup_logger(log_file: str | None) -> logging.Logger:
    """配置同时输出到控制台和文件的 logger。"""
    logger = logging.getLogger("run_server")
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter(
        fmt="%(asctime)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # 控制台 handler（unbuffered，适合 nohup 实时监控）
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    # 文件 handler（可选）
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setFormatter(fmt)
        logger.addHandler(fh)

    return logger


def _make_log_cb(logger: logging.Logger):
    """返回线程安全的日志回调，同时写控制台和文件。"""
    _lock = threading.Lock()

    def log_cb(msg: str) -> None:
        with _lock:
            for line in str(msg).splitlines():
                logger.info(line)

    return log_cb


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="run_server.py",
        description="中文文本词频统计分析工具 — 服务器命令行模式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )

    # ── 必填参数 ────────────────────────────────────────────────────────────
    p.add_argument(
        "--data-dir",
        dest="data_dirs",
        metavar="DIR",
        action="append",
        required=True,
        help="数据文件夹路径（可重复指定多个）",
    )
    p.add_argument(
        "--dict",
        dest="dict_files",
        metavar="FILE",
        action="append",
        required=True,
        help="词典文件路径（.xlsx / .xls / .txt，可重复指定多个）",
    )
    p.add_argument(
        "--output",
        dest="output",
        metavar="FILE",
        required=True,
        help="输出 Excel 文件路径（.xlsx）",
    )

    # ── 列名配置 ────────────────────────────────────────────────────────────
    p.add_argument("--col-stkcd", default="公司代码", help="公司代码列名（默认：公司代码）")
    p.add_argument("--col-year",  default="年份",     help="年份列名（默认：年份）")
    p.add_argument(
        "--col-text",
        dest="col_texts",
        metavar="COL",
        action="append",
        default=None,
        help="文本列名（可重复，默认：文本内容）",
    )

    # ── 分析设置 ────────────────────────────────────────────────────────────
    p.add_argument(
        "--workers",
        type=int,
        default=4,
        metavar="N",
        help="并发文件处理工作进程数（默认：4，建议不超过 CPU 核数）",
    )
    p.add_argument(
        "--no-regex",
        dest="use_regex",
        action="store_false",
        default=True,
        help="使用 jieba 分词模式（默认：正则匹配）",
    )
    p.add_argument(
        "--export-sentences",
        action="store_true",
        default=False,
        help="导出命中句子 Sheet（启用 LLM 时自动开启）",
    )
    p.add_argument(
        "--use-tf",
        action="store_true",
        default=False,
        help="输出词频标准化（TF）占比列",
    )
    p.add_argument(
        "--jieba-userdict",
        default="",
        metavar="FILE",
        help="jieba 用户自定义词典路径（仅 jieba 模式有效）",
    )

    # ── LLM 设置 ────────────────────────────────────────────────────────────
    g = p.add_argument_group("LLM 句子分析（不传 --llm-api-key 则禁用）")
    g.add_argument(
        "--llm-api-key",
        default="",
        metavar="KEY",
        help="API Key（也可通过环境变量 DASHSCOPE_API_KEY 设置）",
    )
    g.add_argument(
        "--llm-model",
        default="qwen-turbo",
        metavar="MODEL",
        help="模型名称（默认：qwen-turbo；服务器长跑推荐 turbo，便宜 7x）",
    )
    g.add_argument(
        "--llm-base-url",
        default="https://dashscope.aliyuncs.com/compatible-mode/v1",
        metavar="URL",
        help="API 基础 URL（默认：DashScope 兼容接口）",
    )
    g.add_argument(
        "--llm-max-sentences",
        type=int,
        default=0,
        metavar="N",
        help="每次运行最多调用 LLM 分析的唯一句子数（0=无上限，默认：0）",
    )
    g.add_argument(
        "--llm-workers",
        type=int,
        default=16,
        metavar="N",
        help="LLM 并发请求线程数（默认：16）",
    )
    g.add_argument(
        "--llm-max-retries",
        type=int,
        default=2,
        metavar="N",
        help="LLM 普通错误最大重试次数（默认：2；429 限速独立最多重试 10 次）",
    )
    g.add_argument(
        "--llm-cache",
        default="",
        metavar="FILE",
        help="LLM 缓存文件路径（默认：与 --output 同目录下的 llm_cache.json）",
    )
    g.add_argument(
        "--llm-system-prompt-file",
        default="",
        metavar="FILE",
        help=(
            "自定义 LLM 系统提示词文件路径（UTF-8 纯文本）。\n"
            "不传则使用内置默认提示词（推荐大多数情况）。\n"
            "注意：修改提示词后，已缓存的结果不会被重新分析；\n"
            "如需全量重跑请删除或重命名 llm_cache.json。"
        ),
    )

    # ── LLM 续跑模式（词频已完成，只跑 LLM）────────────────────────────────
    p.add_argument(
        "--resume-sentences",
        default="",
        metavar="FILE",
        help=(
            "【LLM 续跑模式】跳过词频分析，直接对指定命中句子文件运行 LLM 分析。\n"
            "词频已完成但 LLM 因限流/崩溃失败时使用。\n"
            "优先传 result_sentences_full.csv（完整），次选 result_sentences.xlsx。\n"
            "已在 llm_cache.json 中的句子会自动跳过，无需重复分析。\n"
            "示例：--resume-sentences /mnt/fastdata/output/result_sentences.xlsx"
        ),
    )

    # ── 日志 ────────────────────────────────────────────────────────────────
    p.add_argument(
        "--log-file",
        default="",
        metavar="FILE",
        help="同时将日志写入指定文件（为空则仅输出到控制台）",
    )

    return p.parse_args()


def main() -> int:
    # 确保 stdout/stderr 实时刷新（nohup 下默认全缓冲）
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(line_buffering=True)  # type: ignore[attr-defined]

    args = _parse_args()

    # ── 读取自定义系统提示词（可选）────────────────────────────────────────────
    llm_system_prompt: str = ""
    _prompt_file = getattr(args, "llm_system_prompt_file", "").strip()
    if _prompt_file:
        try:
            llm_system_prompt = Path(_prompt_file).read_text(encoding="utf-8").strip()
        except Exception as _e:
            print(f"[错误] 读取自定义提示词文件失败：{_e}", flush=True)
            return 2

    # ── 推断 LLM 缓存路径 ────────────────────────────────────────────────────
    llm_cache_path: str | None = None
    resolved_api_key = (
        args.llm_api_key.strip()
        or os.getenv("DASHSCOPE_API_KEY", "").strip()
        or os.getenv("QWEN_API_KEY", "").strip()
    )
    analyze_llm = bool(resolved_api_key)

    if analyze_llm:
        if args.llm_cache.strip():
            llm_cache_path = args.llm_cache.strip()
        else:
            output_dir = Path(args.output).parent
            llm_cache_path = str(output_dir / "llm_cache.json")

    # ── Logger ───────────────────────────────────────────────────────────────
    log_file = args.log_file.strip() or None
    logger = _setup_logger(log_file)
    log_cb = _make_log_cb(logger)

    # ── LLM 续跑模式：跳过词频，只跑 LLM ─────────────────────────────────────
    resume_sentences = args.resume_sentences.strip()
    if resume_sentences:
        if not resolved_api_key:
            log_cb("[错误] LLM 续跑模式需要提供 --llm-api-key。")
            return 2
        if not os.path.isfile(resume_sentences):
            log_cb(f"[错误] 命中句子文件不存在：{resume_sentences}")
            return 2

        log_cb("=" * 60)
        log_cb("中文文本词频统计分析工具 — LLM 续跑模式")
        log_cb("（跳过词频分析，直接对已有命中句子续跑 LLM）")
        log_cb("=" * 60)
        log_cb(f"命中句子来源：{resume_sentences}")
        log_cb(f"输出路径：{args.output}")
        log_cb(f"LLM 模型：{args.llm_model} | 并发：{args.llm_workers} | 缓存：{llm_cache_path}")
        log_cb("-" * 60)

        cancel_event = threading.Event()
        import signal
        def _sig(sig, frame):
            if not cancel_event.is_set():
                log_cb("收到中断信号，正在优雅退出…")
                cancel_event.set()
        signal.signal(signal.SIGINT, _sig)
        signal.signal(signal.SIGTERM, _sig)

        try:
            from word_freq_analyzer import run_llm_analysis_only
        except ImportError as exc:
            log_cb(f"[错误] 导入模块失败：{exc}")
            return 2

        Path(args.output).parent.mkdir(parents=True, exist_ok=True)
        try:
            run_llm_analysis_only(
                sentences_source=resume_sentences,
                output_path=args.output,
                llm_api_key=resolved_api_key,
                llm_model=args.llm_model,
                llm_base_url=args.llm_base_url,
                llm_max_sentences=args.llm_max_sentences,
                llm_max_workers=args.llm_workers,
                llm_max_retries=args.llm_max_retries,
                llm_cache_path=llm_cache_path,
                llm_system_prompt=llm_system_prompt,
                log_cb=log_cb,
                cancel_event=cancel_event,
            )
        except Exception as exc:
            log_cb(f"[错误] LLM 续跑失败：{exc}")
            import traceback
            log_cb(traceback.format_exc())
            return 1
        log_cb("=" * 60)
        log_cb("LLM 续跑完成！")
        log_cb("=" * 60)
        return 0

    log_cb("=" * 60)
    log_cb("中文文本词频统计分析工具 — 服务器模式")
    log_cb("=" * 60)
    log_cb(f"数据目录：{args.data_dirs}")
    log_cb(f"词典文件：{args.dict_files}")
    log_cb(f"输出路径：{args.output}")
    log_cb(f"公司代码列：{args.col_stkcd} | 年份列：{args.col_year} | 文本列：{args.col_texts or ['文本内容']}")
    log_cb(f"匹配模式：{'正则' if args.use_regex else 'jieba 分词'} | 文件处理进程：{args.workers}")
    if analyze_llm:
        limit_desc = "无上限" if args.llm_max_sentences <= 0 else str(args.llm_max_sentences)
        log_cb(
            f"LLM：{args.llm_model} | 并发 {args.llm_workers} | "
            f"句子上限 {limit_desc} | 缓存 {llm_cache_path}"
        )
    else:
        log_cb("LLM：未启用（未提供 API Key）")
    log_cb("-" * 60)

    # ── 导入业务模块（放在参数校验之后，避免启动太慢）────────────────────────
    try:
        from word_freq_analyzer import collect_data_files, run_analysis, DictionaryManager
    except ImportError as exc:
        log_cb(f"[错误] 导入业务模块失败：{exc}")
        log_cb("请确认 word_freq_analyzer.py 与本脚本在同一目录，且依赖已安装：")
        log_cb("  pip install pandas openpyxl xlrd jieba openai")
        return 2

    # ── 加载词典 ─────────────────────────────────────────────────────────────
    dict_mgr = DictionaryManager()
    for dict_file in args.dict_files:
        if not os.path.isfile(dict_file):
            log_cb(f"[错误] 词典文件不存在：{dict_file}")
            return 2
        try:
            dict_mgr.import_file(dict_file)
            log_cb(f"已加载词典：{dict_file}（{dict_mgr.total_word_count()} 词）")
        except Exception as exc:
            log_cb(f"[错误] 词典加载失败（{dict_file}）：{exc}")
            return 2

    if dict_mgr.total_word_count() == 0:
        log_cb("[错误] 词典为空，请检查词典文件内容。")
        return 2

    log_cb(f"词典汇总：\n{dict_mgr.summary_text()}")

    # ── 扫描数据文件 ──────────────────────────────────────────────────────────
    for d in args.data_dirs:
        if not os.path.isdir(d):
            log_cb(f"[错误] 数据目录不存在：{d}")
            return 2

    log_cb("正在扫描数据文件…")
    files, ext_counts, scan_warnings = collect_data_files(args.data_dirs)
    for w in scan_warnings:
        log_cb(f"[警告] {w}")
    if not files:
        log_cb("[错误] 未找到任何数据文件（.xlsx / .xls / .csv / .txt）。")
        log_cb(f"已扫描目录：{args.data_dirs}")
        return 2

    ext_summary = "  ".join(f"{ext}: {n}" for ext, n in sorted(ext_counts.items()))
    log_cb(f"共找到 {len(files)} 个文件（{ext_summary}）")

    # ── 确保输出目录存在 ──────────────────────────────────────────────────────
    output_path = args.output
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    # ── 取消事件（Ctrl+C） ────────────────────────────────────────────────────
    cancel_event = threading.Event()

    import signal

    def _sigint_handler(sig, frame):
        if not cancel_event.is_set():
            log_cb("\n收到 Ctrl+C，正在优雅取消（等待当前任务完成）…")
            cancel_event.set()

    signal.signal(signal.SIGINT, _sigint_handler)
    signal.signal(signal.SIGTERM, _sigint_handler)

    # ── 运行分析 ──────────────────────────────────────────────────────────────
    col_texts = args.col_texts if args.col_texts else ["文本内容"]

    try:
        run_analysis(
            files=files,
            dict_mgr=dict_mgr,
            col_stkcd=args.col_stkcd,
            col_year=args.col_year,
            text_columns=col_texts,
            output_path=output_path,
            use_regex=args.use_regex,
            use_stopwords=False,
            use_tf=args.use_tf,
            export_sentences=(args.export_sentences or analyze_llm),
            analyze_llm=analyze_llm,
            llm_api_key=resolved_api_key,
            llm_model=args.llm_model,
            llm_base_url=args.llm_base_url,
            llm_max_sentences=args.llm_max_sentences,
            llm_max_workers=args.llm_workers,
            llm_max_retries=args.llm_max_retries,
            llm_cache_path=llm_cache_path,
            llm_system_prompt=llm_system_prompt,
            analysis_workers=args.workers,
            jieba_userdict=args.jieba_userdict,
            log_cb=log_cb,
            cancel_event=cancel_event,
        )
    except KeyboardInterrupt:
        cancel_event.set()
        log_cb("已中断。")
        return 130
    except Exception as exc:
        log_cb(f"[错误] 分析失败：{exc}")
        import traceback
        log_cb(traceback.format_exc())
        return 1

    if cancel_event.is_set():
        log_cb("分析已取消（部分结果可能已写入输出文件）。")
        return 130

    log_cb("=" * 60)
    log_cb(f"分析完成！结果已保存到：{output_path}")
    log_cb("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
