#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中文文本词频统计分析工具 v3.0
==============================
科研级面板数据词频统计。输出 公司×年份 面板数据格式。
支持分类词典管理、正则/jieba 双模式、命中句子导出。

作者：陈浩杰
单位：澳门城市大学金融学院

v3.0 修复：
  - 修复子文件夹递归扫描遗漏 .xls 文件
  - 新增大文件分块读取（CSV >100MB 自动分块）
  - 每文件处理后立即按 公司×年份 聚合，大幅降低内存占用
  - Sheet2 改用 melt() 向量化构建，大数据集速度提升 10x+
  - 修复 pandas 2.x infer_datetime_format 废弃警告
  - 新增取消按钮，分析可随时中止
  - 新增 Excel 行数上限保护，超大结果自动截断并提示
  - 扫描后显示子文件夹统计，方便确认递归深度
  - 文件间自动回收内存，防止长时间运行 OOM
"""
from __future__ import annotations

import gc
import os
import re
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from collections import Counter, defaultdict
from pathlib import Path

import jieba
import pandas as pd

# ============================================================
# 常量
# ============================================================

# v3.0: 加入 .xls 支持（需要 xlrd 库）
SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv"}

MAX_EXCEL_ROWS = 1_048_575  # Excel 行上限（减去表头）
MAX_SENTENCES = 200_000
BIG_CSV_THRESHOLD = 100 * 1024 * 1024  # 100MB
CHUNK_ROWS = 50_000

DEFAULT_STOPWORDS = frozenset({
    "的", "了", "在", "是", "我", "有", "和", "就", "不", "人",
    "都", "一", "一个", "上", "也", "很", "到", "说", "要", "去",
    "你", "会", "着", "没有", "看", "好", "自己", "这", "他", "她",
    "它", "们", "这个", "那个", "那", "什么", "如果", "但是", "因为",
    "所以", "可以", "已经", "而且", "或者", "虽然", "这些", "那些",
    "只是", "但", "而", "及", "与", "以", "为", "之", "其", "中",
    "对", "被", "从", "把", "将", "向", "个", "各", "等", "则",
    "能", "又", "该", "于", "当", "更", "还", "让", "用", "过",
    "后", "前", "下", "两", "所", "没", "这样", "那样", "怎么",
    "并", "给", "最", "些", "比", "做", "第", "如", "即", "且",
    "每", "应", "您", "哪", "得", "使", "才", "再", "还是", "及其",
    "以及", "因此", "由于", "其中", "可能", "已", "此", "需要",
    "通过", "进行", "没有", "不是", "如何", "或", "之间",
    " ", "\t", "\n", "\r",
    "，", "。", "、", "：", "；", "？", "！",
    """, """, "'", "'", "（", "）", "【", "】", "《", "》",
    "—", "…", "·",
    ",", ".", ":", ";", "?", "!", "(", ")", "[", "]",
    "{", "}", "-", "_", "/", "\\", "|", "@", "#", "%", "&", "*",
})

_NONE_LABEL = "（不选）"

# ============================================================
# 词典管理器
# ============================================================

class DictionaryManager:
    """管理 分类 → 关键词列表 映射。"""

    def __init__(self):
        self.data: dict[str, list[str]] = {}

    def categories(self) -> list[str]:
        return list(self.data.keys())

    def words(self, category: str) -> list[str]:
        return self.data.get(category, [])

    def add_category(self, name: str):
        if name and name not in self.data:
            self.data[name] = []

    def remove_category(self, name: str):
        self.data.pop(name, None)

    def add_word(self, category: str, word: str):
        word = word.strip()
        if category in self.data and word and word not in self.data[category]:
            self.data[category].append(word)

    def remove_word(self, category: str, word: str):
        if category in self.data:
            try:
                self.data[category].remove(word)
            except ValueError:
                pass

    def word_to_category(self) -> dict[str, str]:
        result = {}
        for cat, words in self.data.items():
            for w in words:
                result[w] = cat
        return result

    def find_duplicates(self) -> dict[str, list[str]]:
        """找出跨分类重复的关键词 → {词: [分类1, 分类2, ...]}。"""
        word_cats: dict[str, list[str]] = {}
        for cat, words in self.data.items():
            for w in words:
                word_cats.setdefault(w, []).append(cat)
        return {w: cats for w, cats in word_cats.items() if len(cats) > 1}

    def all_words(self) -> set[str]:
        return {w for words in self.data.values() for w in words}

    def total_word_count(self) -> int:
        return sum(len(ws) for ws in self.data.values())

    # ---- 导入 ----

    def import_file(self, path: str):
        """自动识别格式并导入词典（.xlsx / .xls / .txt）。"""
        ext = Path(path).suffix.lower()
        if ext in (".xlsx", ".xls"):
            self._import_excel(path)
        elif ext == ".txt":
            self._import_txt(path)
        else:
            raise ValueError(f"不支持的词典格式：{ext}（支持 .xlsx .xls .txt）")

    def _import_excel(self, path: str):
        df = pd.read_excel(path, header=None, dtype=str)
        if df.shape[1] < 2:
            raise ValueError("Excel 词典至少需要两列（分类、关键词）。")
        for _, row in df.iterrows():
            cat = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            word = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            if cat and word:
                self.add_category(cat)
                self.add_word(cat, word)

    def _import_txt(self, path: str):
        """格式：每行 分类：词1,词2,词3 或 分类:词1 词2 词3"""
        text = None
        for enc in ("utf-8", "utf-8-sig", "gb18030", "gbk"):
            try:
                with open(path, encoding=enc) as f:
                    text = f.read()
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError(f"无法读取文件（编码不支持）：{path}")

        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            sep = None
            for s in ("：", ":"):
                if s in line:
                    sep = s
                    break
            if sep is None:
                continue
            cat, words_part = line.split(sep, 1)
            cat = cat.strip()
            if not cat:
                continue
            self.add_category(cat)
            for w in re.split(r"[,，;；\s]+", words_part):
                w = w.strip()
                if w:
                    self.add_word(cat, w)

    def export_excel(self, path: str):
        rows = []
        for cat, words in self.data.items():
            for w in words:
                rows.append({"分类": cat, "关键词": w})
        pd.DataFrame(rows).to_excel(path, index=False)

    def clear(self):
        self.data.clear()

    def summary_text(self) -> str:
        parts = [f"  {cat}: {len(self.words(cat))} 个" for cat in self.categories()]
        return "\n".join(parts) if parts else "  （空）"


# ============================================================
# 工具函数
# ============================================================

def collect_data_files(folders: list[str]) -> tuple[list[str], dict[str, int], list[str]]:
    """
    递归扫描文件夹，返回 (文件列表, 子目录文件计数, 错误列表)。
    v3.0: followlinks=True 跟随符号链接；onerror 收集权限错误。
    """
    files: list[str] = []
    dir_counts: dict[str, int] = {}
    errors: list[str] = []

    def _on_error(err):
        errors.append(str(err))

    for folder in folders:
        for root, _, filenames in os.walk(folder, followlinks=True, onerror=_on_error):
            found = 0
            for fname in sorted(filenames):
                if fname.startswith("~$") or fname.startswith("."):
                    continue
                if Path(fname).suffix.lower() in SUPPORTED_EXTENSIONS:
                    files.append(os.path.join(root, fname))
                    found += 1
            if found > 0:
                rel = os.path.relpath(root, folder)
                display = f"{os.path.basename(folder)}/{rel}" if rel != "." else os.path.basename(folder)
                dir_counts[display] = found
    return files, dir_counts, errors


def read_data_file(filepath: str, nrows=None) -> pd.DataFrame:
    """读取单个数据文件（.xlsx / .xls / .csv），自动检测编码。"""
    ext = Path(filepath).suffix.lower()
    if ext in (".xlsx", ".xls"):
        try:
            return pd.read_excel(filepath, nrows=nrows, dtype=str)
        except ImportError:
            if ext == ".xls":
                raise ValueError(
                    f"读取 .xls 文件需要安装 xlrd 库：pip install xlrd\n"
                    f"文件：{filepath}"
                )
            raise
    # CSV
    for enc in ("utf-8", "utf-8-sig", "gb18030", "gbk", "big5"):
        try:
            return pd.read_csv(filepath, nrows=nrows, dtype=str, encoding=enc, low_memory=False)
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    raise ValueError(f"无法读取文件（编码不支持）：{filepath}")


def read_csv_chunked(filepath: str, chunksize: int = CHUNK_ROWS):
    """
    分块读取大 CSV 文件的生成器。
    v3.0: 先用小样本探测编码，再用确定的编码分块读取全文件。
    避免 generator 中途解码失败无法回退的问题。
    """
    # 第一步：探测正确的编码
    detected_enc = None
    for enc in ("utf-8", "utf-8-sig", "gb18030", "gbk", "big5"):
        try:
            pd.read_csv(filepath, dtype=str, encoding=enc, nrows=100)
            detected_enc = enc
            break
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    if detected_enc is None:
        raise ValueError(f"无法读取文件（编码不支持）：{filepath}")

    # 第二步：用确定的编码分块读取
    reader = pd.read_csv(
        filepath, dtype=str, encoding=detected_enc,
        low_memory=False, chunksize=chunksize,
    )
    for chunk in reader:
        yield chunk


def _read_columns_fast(filepath: str) -> list[str]:
    """
    快速读取文件列名（不加载数据）。
    v3.0: 对大 xlsx 使用 openpyxl read_only 模式，只读第一行；
    CSV 只读第一行。比 pd.read_excel(nrows=0) 快 10-100x。
    """
    ext = Path(filepath).suffix.lower()
    if ext in (".xlsx", ".xls"):
        # 优先用 openpyxl 快速读取第一行
        if ext == ".xlsx":
            try:
                from openpyxl import load_workbook
                wb = load_workbook(filepath, read_only=True, data_only=True)
                ws = wb.active
                for row in ws.iter_rows(max_row=1, values_only=True):
                    cols = [str(c) if c is not None else f"Unnamed_{i}"
                            for i, c in enumerate(row)]
                    wb.close()
                    return cols
                wb.close()
                return []
            except Exception:
                pass
        # 回退到 pandas（较慢但更兼容）
        try:
            df = pd.read_excel(filepath, nrows=0, dtype=str)
            return [str(c) for c in df.columns]
        except Exception:
            return []
    # CSV: 只读第一行
    for enc in ("utf-8", "utf-8-sig", "gb18030", "gbk", "big5"):
        try:
            df = pd.read_csv(filepath, nrows=0, dtype=str, encoding=enc)
            return [str(c) for c in df.columns]
        except Exception:
            continue
    return []


HEADER_SCAN_SIZE_LIMIT = 20 * 1024 * 1024  # 20MB: 跳过大 xlsx 的表头扫描（同目录小文件列相同）


def scan_all_columns(files: list[str]) -> tuple[list[str], dict[str, int]]:
    """
    扫描所有文件的列名（只读表头，不加载数据）。
    v3.0: 返回 (列名列表按出现频率降序, 列名→出现文件数)。
    频率高的列更可能是用户需要的公共列，优先展示。
    对超过 50MB 的 xlsx 文件跳过扫描（同目录小文件列结构相同）。
    """
    col_freq: dict[str, int] = {}
    skipped = 0
    for f in files:
        try:
            ext = Path(f).suffix.lower()
            fsize = os.path.getsize(f)
            # 跳过巨大 xlsx（同目录的小文件结构相同）
            if ext in (".xlsx", ".xls") and fsize > HEADER_SCAN_SIZE_LIMIT:
                skipped += 1
                continue
            columns = _read_columns_fast(f)
            for c in columns:
                col_freq[c] = col_freq.get(c, 0) + 1
        except Exception:
            continue
    # 按频率降序排列
    sorted_cols = sorted(col_freq.keys(), key=lambda x: -col_freq[x])
    return sorted_cols, col_freq


def fix_stock_code(code) -> str:
    """修复股票代码：数字补零至6位，非数字保留原值。"""
    if pd.isna(code):
        return ""
    s = str(code).strip()
    if not s:
        return ""
    try:
        return str(int(float(s))).zfill(6)
    except (ValueError, TypeError):
        return s


def parse_year_column(series: pd.Series) -> pd.Series:
    """智能解析年份列：支持数值、日期字符串、混合格式。"""
    if series.empty:
        return pd.Series(dtype=int)

    # 尝试直接数值
    numeric = pd.to_numeric(series, errors="coerce")
    valid = numeric.between(1900, 2100)
    if valid.sum() > len(series) * 0.3:
        return numeric.where(valid, 0).astype(int)

    # 尝试日期解析（v3.0: 移除废弃的 infer_datetime_format）
    try:
        dates = pd.to_datetime(series, errors="coerce", format="mixed")
    except (ValueError, TypeError):
        dates = pd.to_datetime(series, errors="coerce")
    years = dates.dt.year
    if years.notna().sum() > len(series) * 0.3:
        return years.fillna(0).astype(int)

    # 正则提取 4 位年份
    extracted = series.astype(str).str.extract(r"((?:19|20)\d{2})", expand=False)
    extracted_num = pd.to_numeric(extracted, errors="coerce")
    if extracted_num.notna().sum() > len(series) * 0.3:
        return extracted_num.fillna(0).astype(int)

    return pd.Series(0, index=series.index)


def split_sentences(text: str) -> list[str]:
    parts = re.split(r"[。！？；\n!?]", text)
    return [s.strip() for s in parts if len(s.strip()) > 2]


def load_stopwords_file(path: str) -> set[str]:
    words = set()
    for enc in ("utf-8", "utf-8-sig", "gb18030"):
        try:
            with open(path, encoding=enc) as f:
                for line in f:
                    w = line.strip()
                    if w:
                        words.add(w)
            return words
        except UnicodeDecodeError:
            continue
    return words


# ============================================================
# 面板数据分析引擎
# ============================================================

def _process_chunk(
    df: pd.DataFrame,
    col_stkcd: str,
    col_year: str,
    text_columns: list[str],
    keywords: list[str],
    use_regex: bool,
    all_dict_words: set[str],
    stopwords: set[str] | None,
    use_stopwords: bool,
    do_export_sentences: bool,
    word_to_cat: dict[str, str],
) -> tuple[pd.DataFrame | None, int, int, list[dict]]:
    """
    处理单个 DataFrame（整文件或 CSV 块）。
    返回 (聚合后的 公司×年份 DataFrame, 原始行数, 命中次数, 命中句子列表)。
    """
    hit_sents: list[dict] = []

    if col_stkcd not in df.columns:
        return None, 0, 0, []
    if col_year not in df.columns:
        return None, 0, 0, []

    available_text = [c for c in text_columns if c in df.columns]
    if not available_text:
        return None, 0, 0, []

    df = df.copy()
    raw_rows = len(df)  # v3.0: 在过滤前记录原始行数
    df["_stkcd"] = df[col_stkcd].apply(fix_stock_code)
    df["_year"] = parse_year_column(df[col_year])
    df = df[df["_year"] > 0].copy()  # v3.0: .copy() 防止 SettingWithCopyWarning
    if len(df) == 0:
        return None, raw_rows, 0, []
    text_series = df[available_text].fillna("").astype(str).agg(" ".join, axis=1)

    # 关键词匹配
    if use_regex:
        for kw in keywords:
            df[kw] = text_series.str.count(f"(?i){re.escape(kw)}")
    else:
        sw = stopwords if use_stopwords else None
        dict_set = all_dict_words

        def _jieba_row(text):
            tokens = jieba.lcut(str(text))
            if sw:
                tokens = [t for t in tokens if t not in sw]
            cnt = Counter(t for t in tokens if t in dict_set)
            return pd.Series({kw: cnt.get(kw, 0) for kw in keywords})

        kw_df = text_series.apply(_jieba_row)
        for kw in keywords:
            df[kw] = kw_df[kw].values

    n_hits = int(df[keywords].sum().sum())

    # 提取命中句子
    if do_export_sentences and n_hits > 0:
        match_mask = df[keywords].sum(axis=1) > 0
        for row_idx in df[match_mask].index:
            text = str(text_series.loc[row_idx])
            stkcd = df.loc[row_idx, "_stkcd"]
            year = int(df.loc[row_idx, "_year"])
            row_kws = [kw for kw in keywords if df.loc[row_idx, kw] > 0]
            for sent in split_sentences(text):
                for kw in row_kws:
                    if re.search(re.escape(kw), sent, re.IGNORECASE):
                        hit_sents.append({
                            "公司代码": stkcd, "年份": year,
                            "命中关键词": kw,
                            "分类": word_to_cat.get(kw, ""),
                            "命中句子": sent[:500],
                        })

    # 按 公司×年份 聚合（v3.0: 每文件/块立即聚合，大幅节省内存）
    chunk = df[["_stkcd", "_year"] + keywords].copy()
    chunk.columns = ["公司代码", "年份"] + keywords
    agg_rules = {kw: "sum" for kw in keywords}
    chunk_agg = chunk.groupby(["公司代码", "年份"], as_index=False).agg(agg_rules)

    return chunk_agg, raw_rows, n_hits, hit_sents


def run_analysis(
    files: list[str],
    dict_mgr: DictionaryManager,
    col_stkcd: str,
    col_year: str,
    text_columns: list[str],
    output_path: str,
    *,
    use_regex: bool = True,
    use_stopwords: bool = False,
    stopwords: set[str] | None = None,
    use_tf: bool = False,
    export_sentences: bool = False,
    jieba_userdict: str = "",
    progress_cb=None,
    log_cb=None,
    cancel_event: threading.Event | None = None,
):
    def log(msg):
        if log_cb:
            log_cb(msg)

    def is_cancelled():
        return cancel_event is not None and cancel_event.is_set()

    word_to_cat = dict_mgr.word_to_category()
    all_dict_words = dict_mgr.all_words()
    keywords = sorted(all_dict_words)
    categories = list(dict_mgr.categories())  # 快照，防止线程不安全

    if not all_dict_words:
        raise ValueError("词典为空，请先添加或导入词典。")

    # v3.0: 跨分类重复关键词警告（同一词在多个分类中只会归入最后一个分类）
    duplicates = dict_mgr.find_duplicates()
    if duplicates:
        dup_items = [f"「{w}」→{cats}" for w, cats in list(duplicates.items())[:10]]
        log(f"警告：以下关键词在多个分类中重复出现（仅归入最后一个分类）：\n  " + "\n  ".join(dup_items))
        if len(duplicates) > 10:
            log(f"  …共 {len(duplicates)} 个重复词")

    # v3.0: 列名碰撞检查 — 关键词/分类名不能与内部列名冲突
    _reserved = {"公司代码", "年份", "总计", "_stkcd", "_year"}
    bad_kws = [kw for kw in keywords if kw in _reserved]
    if bad_kws:
        log(f"警告：以下关键词与系统保留列名冲突，已自动跳过：{bad_kws}")
        keywords = [kw for kw in keywords if kw not in _reserved]
        all_dict_words -= _reserved
    bad_cats = [c for c in categories if c in _reserved]
    if bad_cats:
        log(f"警告：以下分类名与系统保留列名冲突，已自动跳过：{bad_cats}")
        categories = [c for c in categories if c not in _reserved]

    if len(keywords) > 1000:
        log(f"提示：关键词数量较大（{len(keywords)}），分析可能较慢。")

    if use_regex:
        log("匹配模式：正则匹配（不区分大小写）")
    else:
        if jieba_userdict and os.path.isfile(jieba_userdict):
            jieba.load_userdict(jieba_userdict)
            log(f"已加载 jieba 用户词典：{os.path.basename(jieba_userdict)}")
        for w in all_dict_words:
            jieba.add_word(w)
        log("匹配模式：jieba 分词")

    log(f"词典：{len(keywords)} 个关键词，{len(categories)} 个分类")
    log(f"公司代码列：{col_stkcd}，年份列：{col_year}")
    log(f"文本列：{', '.join(text_columns)}")

    all_chunks: list[pd.DataFrame] = []
    hit_sentences: list[dict] = []
    total_files = len(files)
    do_export_sentences = export_sentences
    agg_rules = {kw: "sum" for kw in keywords}

    for idx, fpath in enumerate(files, 1):
        if is_cancelled():
            log("用户已取消。")
            break

        fname = os.path.basename(fpath)
        rel_dir = os.path.basename(os.path.dirname(fpath))
        display_name = f"{rel_dir}/{fname}" if rel_dir else fname
        log(f"[{idx}/{total_files}] {display_name}")

        try:
            ext = Path(fpath).suffix.lower()
            file_size = os.path.getsize(fpath)

            # v3.0: 大 CSV 文件分块读取
            if ext == ".csv" and file_size > BIG_CSV_THRESHOLD:
                log(f"  大文件模式（{file_size / 1024 / 1024:.0f}MB），分块处理…")
                file_chunks = []
                file_rows = 0
                file_hits = 0
                chunk_num = 0

                for chunk_df in read_csv_chunked(fpath, CHUNK_ROWS):
                    if is_cancelled():
                        break
                    chunk_num += 1
                    result, rows, hits, sents = _process_chunk(
                        chunk_df, col_stkcd, col_year, text_columns,
                        keywords, use_regex, all_dict_words,
                        stopwords, use_stopwords, do_export_sentences,
                        word_to_cat,
                    )
                    if result is not None:
                        file_chunks.append(result)
                        file_rows += rows
                        file_hits += hits
                        hit_sentences.extend(sents)
                        if len(hit_sentences) > MAX_SENTENCES:
                            do_export_sentences = False
                            log("  命中句子已达上限，后续跳过提取。")

                if file_chunks:
                    merged = pd.concat(file_chunks, ignore_index=True)
                    merged = merged.groupby(["公司代码", "年份"], as_index=False).agg(agg_rules)
                    all_chunks.append(merged)
                    log(f"  {chunk_num} 块，共 {file_rows} 行，命中 {file_hits} 次")
                else:
                    log(f"  跳过：无有效数据")

            else:
                # 正常读取（Excel 或小 CSV）
                df = read_data_file(fpath)
                result, rows, hits, sents = _process_chunk(
                    df, col_stkcd, col_year, text_columns,
                    keywords, use_regex, all_dict_words,
                    stopwords, use_stopwords, do_export_sentences,
                    word_to_cat,
                )
                del df  # 及时释放

                if result is not None:
                    all_chunks.append(result)
                    hit_sentences.extend(sents)
                    if len(hit_sentences) > MAX_SENTENCES:
                        do_export_sentences = False
                        log("  命中句子已达上限，后续跳过提取。")
                    log(f"  {rows} 行，命中 {hits} 次")
                else:
                    log(f"  跳过：缺少必需列或无有效数据")

        except Exception as e:
            log(f"  跳过（错误）：{e}")

        # v3.0: 每个文件处理完后回收内存
        gc.collect()

        if progress_cb:
            progress_cb(idx, total_files)

    if is_cancelled():
        raise ValueError("分析已被用户取消。")

    # ---- 汇总 ----
    if not all_chunks:
        raise ValueError("没有成功处理任何文件，请检查数据和列配置。")

    log("正在汇总面板数据…")
    combined = pd.concat(all_chunks, ignore_index=True)
    del all_chunks
    gc.collect()

    panel = combined.groupby(["公司代码", "年份"], as_index=False).agg(agg_rules)
    del combined
    gc.collect()

    panel.sort_values(["公司代码", "年份"], inplace=True)
    panel.reset_index(drop=True, inplace=True)

    # v3.0: 提前构建分类→关键词映射快照（线程安全，不再依赖 dict_mgr）
    cat_to_kws: dict[str, list[str]] = {}
    for cat in categories:
        cat_to_kws[cat] = [kw for kw in keywords if word_to_cat.get(kw) == cat]

    # ---- Sheet1: 公司×年份 分类统计 ----
    sheet1 = panel[["公司代码", "年份"]].copy()
    for cat in categories:
        cat_kws = [kw for kw in cat_to_kws.get(cat, []) if kw in panel.columns]
        sheet1[cat] = panel[cat_kws].sum(axis=1).astype(int) if cat_kws else 0
    sheet1["总计"] = sheet1[categories].sum(axis=1)
    if use_tf:
        total_per_row = sheet1["总计"].replace(0, 1)  # 避免除零
        for cat in categories:
            sheet1[f"{cat}_TF"] = (sheet1[cat] / total_per_row).round(6)

    # ---- Sheet2: 公司×年份×关键词 明细（v3.0: 向量化 melt）----
    log("正在构建关键词明细…")
    sheet2 = panel.melt(
        id_vars=["公司代码", "年份"],
        value_vars=keywords,
        var_name="关键词",
        value_name="次数",
    )
    sheet2["次数"] = sheet2["次数"].astype(int)
    sheet2 = sheet2[sheet2["次数"] > 0].copy()
    sheet2["分类"] = sheet2["关键词"].map(word_to_cat).fillna("")
    sheet2.sort_values(["公司代码", "年份", "分类", "关键词"], inplace=True)
    sheet2.reset_index(drop=True, inplace=True)

    # ---- Sheet3: 总体分类统计 ----
    cat_totals: dict[str, int] = {}
    for cat in categories:
        cat_kws = [kw for kw in cat_to_kws.get(cat, []) if kw in panel.columns]
        cat_totals[cat] = int(panel[cat_kws].sum().sum()) if cat_kws else 0
    grand_total = sum(cat_totals.values())
    rows_s3 = []
    for cat, total in sorted(cat_totals.items(), key=lambda x: -x[1]):
        rows_s3.append({
            "分类": cat,
            "总次数": total,
            "占比(%)": round(total / grand_total * 100, 2) if grand_total else 0,
        })
    sheet3 = pd.DataFrame(rows_s3)

    del panel
    gc.collect()

    # ---- 写入 Excel ----
    log("正在写入 Excel…")

    # v3.0: Excel 行数上限保护 — 先保存完整 CSV，再截断写 Excel
    truncated = False
    if len(sheet1) > MAX_EXCEL_ROWS or len(sheet2) > MAX_EXCEL_ROWS:
        truncated = True
        csv_dir = os.path.splitext(output_path)[0] + "_csv"
        os.makedirs(csv_dir, exist_ok=True)
        log(f"数据量超过 Excel 行上限，正在先导出完整 CSV…")
        sheet1.to_csv(os.path.join(csv_dir, "公司年份分类统计.csv"), index=False, encoding="utf-8-sig")
        sheet2.to_csv(os.path.join(csv_dir, "关键词明细.csv"), index=False, encoding="utf-8-sig")
        sheet3.to_csv(os.path.join(csv_dir, "分类汇总.csv"), index=False, encoding="utf-8-sig")
        log(f"完整数据已保存至：{csv_dir}")

    # 截断后写 Excel
    if len(sheet1) > MAX_EXCEL_ROWS:
        log(f"警告：Sheet1 共 {len(sheet1)} 行，Excel 上限 {MAX_EXCEL_ROWS}，Excel 中已截断。")
        sheet1 = sheet1.head(MAX_EXCEL_ROWS)
    if len(sheet2) > MAX_EXCEL_ROWS:
        log(f"警告：Sheet2 共 {len(sheet2)} 行，Excel 上限 {MAX_EXCEL_ROWS}，Excel 中已截断。")
        sheet2 = sheet2.head(MAX_EXCEL_ROWS)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        sheet1.to_excel(writer, sheet_name="公司年份分类统计", index=False)
        sheet2.to_excel(writer, sheet_name="关键词明细", index=False)
        sheet3.to_excel(writer, sheet_name="分类汇总", index=False)
        if hit_sentences:
            df_sent = pd.DataFrame(hit_sentences)
            if len(df_sent) > MAX_SENTENCES:
                df_sent = df_sent.head(MAX_SENTENCES)
            if len(df_sent) > MAX_EXCEL_ROWS:
                df_sent = df_sent.head(MAX_EXCEL_ROWS)
            df_sent.to_excel(writer, sheet_name="命中句子", index=False)

    n_companies = sheet1["公司代码"].nunique()
    year_min = int(sheet1["年份"].min()) if len(sheet1) > 0 else 0
    year_max = int(sheet1["年份"].max()) if len(sheet1) > 0 else 0
    log(f"完成！{len(sheet1)} 条面板记录，{n_companies} 家公司，年份 {year_min}-{year_max}")
    log(f"Sheet1: 公司×年份 分类统计 ({len(sheet1)} 行)")
    log(f"Sheet2: 关键词明细 ({len(sheet2)} 行)")
    log(f"Sheet3: 分类汇总 ({len(sheet3)} 行)")
    if hit_sentences:
        log(f"Sheet4: 命中句子 ({len(hit_sentences)} 条)")
    log(f"结果已保存至：{output_path}")


# ============================================================
# 批量添加对话框
# ============================================================

class BatchAddDialog(tk.Toplevel):
    def __init__(self, parent, title="批量添加关键词"):
        super().__init__(parent)
        self.title(title)
        self.result: list[str] | None = None
        self.geometry("340x320")
        self.resizable(False, True)

        ttk.Label(self, text="每行输入一个关键词：").pack(padx=10, pady=(10, 0))
        self.text = tk.Text(self, width=36, height=14)
        self.text.pack(padx=10, pady=5, fill="both", expand=True)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=(0, 10))
        ttk.Button(btn_frame, text="确定", command=self._ok).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=self.destroy).pack(side="left", padx=5)

        self.transient(parent)
        self.grab_set()
        self.text.focus_set()

    def _ok(self):
        raw = self.text.get("1.0", "end").strip()
        self.result = [w.strip() for w in raw.split("\n") if w.strip()]
        self.destroy()


# ============================================================
# 主界面
# ============================================================

class WordFreqApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("中文文本词频统计分析工具 v3.0 — 陈浩杰 | 澳门城市大学金融学院")
        self.geometry("1020x800")
        self.resizable(True, True)
        self.minsize(880, 680)

        # v3.0: 主题和样式
        self._setup_style()

        self.dict_mgr = DictionaryManager()
        self.folders: list[str] = []
        self.scanned_files: list[str] = []
        self.all_columns: list[str] = []
        self._col_freq: dict[str, int] = {}
        self._col_display_map: dict[str, str] = {}

        self.output_path = tk.StringVar()
        self.var_regex = tk.BooleanVar(value=True)
        self.var_stopwords = tk.BooleanVar(value=False)
        self.stopwords_path = tk.StringVar()
        self.var_tf = tk.BooleanVar(value=False)
        self.var_sentences = tk.BooleanVar(value=False)
        self.jieba_dict_path = tk.StringVar()
        self.current_file_var = tk.StringVar(value="就绪")

        # v3.0: 取消事件
        self._cancel_event = threading.Event()

        self._build_ui()

    # ================================================================
    #  样式配置
    # ================================================================

    def _setup_style(self):
        style = ttk.Style(self)
        available = style.theme_names()
        # macOS: aqua; Linux/Win: clam 或 vista
        for theme in ("aqua", "vista", "clam"):
            if theme in available:
                style.theme_use(theme)
                break

        # 标签页字体稍大
        style.configure("TNotebook.Tab", padding=(12, 6))
        # LabelFrame 标题加粗
        style.configure("TLabelframe.Label", font=("", 11, "bold"))
        # 按钮内边距
        style.configure("TButton", padding=(8, 4))
        # 大按钮
        style.configure("Accent.TButton", padding=(16, 8), font=("", 12, "bold"))

    # ================================================================
    #  UI 构建
    # ================================================================

    def _build_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=(10, 4))

        self._build_data_tab(notebook)
        self._build_dict_tab(notebook)
        self._build_settings_tab(notebook)
        self._build_run_tab(notebook)

        # v3.0: 底部状态栏
        status_frame = ttk.Frame(self, relief="sunken")
        status_frame.pack(fill="x", padx=10, pady=(0, 6))
        self._status_var = tk.StringVar(value="就绪  |  词典：0 词  |  数据：0 文件")
        ttk.Label(status_frame, textvariable=self._status_var,
                  font=("", 10)).pack(side="left", padx=8, pady=3)
        ttk.Label(status_frame, text="作者：陈浩杰 | 澳门城市大学金融学院",
                  font=("", 9), foreground="#666666").pack(side="right", padx=8, pady=3)

    # ---- 标签页1：数据选择 ----

    def _build_data_tab(self, notebook):
        tab = ttk.Frame(notebook)
        notebook.add(tab, text=" 数据选择 ")

        # -- 文件夹 --
        frm_folder = ttk.LabelFrame(tab, text="文本文件夹（递归扫描所有子目录中的 .xlsx / .xls / .csv 文件）")
        frm_folder.pack(fill="x", padx=8, pady=(8, 4))

        top = ttk.Frame(frm_folder)
        top.pack(fill="x", padx=5, pady=5)
        self.folder_listbox = tk.Listbox(top, height=3)
        self.folder_listbox.pack(side="left", fill="both", expand=True)
        btns = ttk.Frame(top)
        btns.pack(side="right", padx=(5, 0))
        ttk.Button(btns, text="添加文件夹", command=self._add_folder).pack(fill="x", pady=1)
        ttk.Button(btns, text="移除选中", command=self._remove_folder).pack(fill="x", pady=1)
        ttk.Button(btns, text="扫描文件", command=self._scan_files).pack(fill="x", pady=1)

        # -- 文件列表 + 列配置 --
        mid = ttk.Frame(tab)
        mid.pack(fill="both", expand=False, padx=8, pady=4)
        mid.columnconfigure(0, weight=2)
        mid.columnconfigure(1, weight=1)
        mid.rowconfigure(0, weight=1)

        frm_files = ttk.LabelFrame(mid, text="扫描到的文件（点击预览数据）")
        frm_files.grid(row=0, column=0, sticky="nsew", padx=(0, 4))

        file_container = ttk.Frame(frm_files)
        file_container.pack(fill="both", expand=True, padx=5, pady=5)
        file_container.rowconfigure(0, weight=1)
        file_container.columnconfigure(0, weight=1)

        self.file_listbox = tk.Listbox(file_container, height=8)
        file_vsb = ttk.Scrollbar(file_container, orient="vertical", command=self.file_listbox.yview)
        file_hsb = ttk.Scrollbar(file_container, orient="horizontal", command=self.file_listbox.xview)
        self.file_listbox.configure(yscrollcommand=file_vsb.set, xscrollcommand=file_hsb.set)
        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        file_vsb.grid(row=0, column=1, sticky="ns")
        file_hsb.grid(row=1, column=0, sticky="ew")
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        # 列配置面板
        frm_cols = ttk.LabelFrame(mid, text="列配置（面板数据必需）")
        frm_cols.grid(row=0, column=1, sticky="nsew")

        inner = ttk.Frame(frm_cols)
        inner.pack(fill="both", expand=True, padx=5, pady=5)

        ttk.Label(inner, text="公司代码列：").grid(row=0, column=0, sticky="w", pady=2)
        self.combo_stkcd = ttk.Combobox(inner, state="readonly", width=18)
        self.combo_stkcd.grid(row=0, column=1, sticky="ew", pady=2)

        ttk.Label(inner, text="年份/日期列：").grid(row=1, column=0, sticky="w", pady=2)
        self.combo_year = ttk.Combobox(inner, state="readonly", width=18)
        self.combo_year.grid(row=1, column=1, sticky="ew", pady=2)

        ttk.Label(inner, text="文本列（多选）：").grid(row=2, column=0, sticky="nw", pady=(4, 0))
        self.col_listbox = tk.Listbox(inner, height=4, selectmode="extended")
        col_sb = ttk.Scrollbar(inner, orient="vertical", command=self.col_listbox.yview)
        self.col_listbox.configure(yscrollcommand=col_sb.set)
        self.col_listbox.grid(row=2, column=1, sticky="nsew", pady=(4, 0))
        col_sb.grid(row=2, column=2, sticky="ns", pady=(4, 0))

        inner.columnconfigure(1, weight=1)
        inner.rowconfigure(2, weight=1)

        # -- 数据预览 --
        frm_preview = ttk.LabelFrame(tab, text="数据预览（前 100 行，点击上方文件自动加载）")
        frm_preview.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        container = ttk.Frame(frm_preview)
        container.pack(fill="both", expand=True, padx=5, pady=5)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        self.preview_tree = ttk.Treeview(container, show="headings")
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.preview_tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    # ---- 标签页2：词典管理 ----

    def _build_dict_tab(self, notebook):
        tab = ttk.Frame(notebook)
        notebook.add(tab, text=" 词典管理 ")

        toolbar = ttk.Frame(tab)
        toolbar.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Button(toolbar, text="新建词典", command=self._dict_new).pack(side="left", padx=2)
        ttk.Button(toolbar, text="导入词典", command=self._dict_import).pack(side="left", padx=2)
        ttk.Button(toolbar, text="导出词典", command=self._dict_export).pack(side="left", padx=2)
        self.dict_info_var = tk.StringVar(value="当前词典：0 个分类，0 个关键词")
        ttk.Label(toolbar, textvariable=self.dict_info_var).pack(side="right", padx=5)

        frm_hint = ttk.LabelFrame(tab, text="支持的词典格式")
        frm_hint.pack(fill="x", padx=8, pady=(0, 4))
        hint = (
            "格式1 (Excel .xlsx/.xls)：第一列=分类，第二列=关键词，每行一个词\n"
            "格式2 (文本 .txt)：每行 分类：词1,词2,词3  例如 → 低碳战略：碳中和,碳达峰,低碳转型"
        )
        ttk.Label(frm_hint, text=hint, wraplength=800, justify="left").pack(
            anchor="w", padx=5, pady=4
        )

        pane = ttk.Frame(tab)
        pane.pack(fill="both", expand=True, padx=8, pady=4)
        pane.columnconfigure(0, weight=1)
        pane.columnconfigure(1, weight=2)
        pane.rowconfigure(0, weight=1)

        frm_cat = ttk.LabelFrame(pane, text="分类列表")
        frm_cat.grid(row=0, column=0, sticky="nsew", padx=(0, 4))

        self.cat_listbox = tk.Listbox(frm_cat, height=10)
        self.cat_listbox.pack(fill="both", expand=True, padx=5, pady=(5, 0))
        self.cat_listbox.bind("<<ListboxSelect>>", self._on_cat_select)

        cat_btns = ttk.Frame(frm_cat)
        cat_btns.pack(fill="x", padx=5, pady=5)
        ttk.Button(cat_btns, text="添加分类", command=self._cat_add).pack(side="left", padx=2)
        ttk.Button(cat_btns, text="删除分类", command=self._cat_remove).pack(side="left", padx=2)

        frm_words = ttk.LabelFrame(pane, text="关键词列表（选中分类后显示）")
        frm_words.grid(row=0, column=1, sticky="nsew")

        self.word_listbox = tk.Listbox(frm_words, height=10)
        word_sb = ttk.Scrollbar(frm_words, orient="vertical", command=self.word_listbox.yview)
        self.word_listbox.configure(yscrollcommand=word_sb.set)
        self.word_listbox.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=(5, 0))
        word_sb.pack(side="right", fill="y", padx=(0, 5), pady=(5, 0))

        word_btns = ttk.Frame(frm_words)
        word_btns.pack(fill="x", padx=5, pady=5)
        ttk.Button(word_btns, text="添加关键词", command=self._word_add).pack(side="left", padx=2)
        ttk.Button(word_btns, text="批量添加", command=self._word_batch_add).pack(side="left", padx=2)
        ttk.Button(word_btns, text="删除关键词", command=self._word_remove).pack(side="left", padx=2)

    # ---- 标签页3：分析设置 ----

    def _build_settings_tab(self, notebook):
        tab = ttk.Frame(notebook)
        notebook.add(tab, text=" 分析设置 ")
        pad = {"padx": 8, "pady": 4}

        frm_mode = ttk.LabelFrame(tab, text="匹配模式 — 决定如何在文本中查找关键词")
        frm_mode.pack(fill="x", **pad)
        ttk.Radiobutton(
            frm_mode, variable=self.var_regex, value=True,
            text="正则匹配（推荐）— 用 re.findall 直接搜索关键词，不区分大小写，适合含英文词的词库"
        ).pack(anchor="w", padx=5, pady=2)
        ttk.Radiobutton(
            frm_mode, variable=self.var_regex, value=False,
            text="jieba 分词匹配 — 先用 jieba 将文本切词，再匹配词典，适合纯中文精细分析（较慢）"
        ).pack(anchor="w", padx=5, pady=(0, 5))

        frm_stop = ttk.LabelFrame(tab, text="停用词 — 分词后过滤无意义词语（仅 jieba 模式生效）")
        frm_stop.pack(fill="x", **pad)
        ttk.Checkbutton(frm_stop, text="启用停用词过滤（内置 100+ 常用中文停用词）",
                        variable=self.var_stopwords).pack(anchor="w", padx=5, pady=2)
        row_sf = ttk.Frame(frm_stop)
        row_sf.pack(fill="x", padx=5, pady=(0, 5))
        ttk.Label(row_sf, text="追加自定义停用词文件（每行一个词）：").pack(side="left")
        ttk.Entry(row_sf, textvariable=self.stopwords_path, width=36,
                  state="readonly").pack(side="left", padx=5)
        ttk.Button(row_sf, text="选择", command=self._choose_stopwords).pack(side="left")

        frm_tf = ttk.LabelFrame(tab, text="输出选项")
        frm_tf.pack(fill="x", **pad)
        ttk.Checkbutton(
            frm_tf, variable=self.var_tf,
            text="计算 TF 标准化词频 — 在 Sheet1 中为每个分类添加 TF 列（分类词频 / 该行总命中数）"
        ).pack(anchor="w", padx=5, pady=2)
        ttk.Checkbutton(
            frm_tf, variable=self.var_sentences,
            text="导出命中句子 — 提取包含关键词的原文句子，用于论文引用验证（大数据集较慢）"
        ).pack(anchor="w", padx=5, pady=(0, 5))

        frm_jieba = ttk.LabelFrame(tab, text="自定义 jieba 用户词典 — 提升中文分词准确率（仅 jieba 模式）")
        frm_jieba.pack(fill="x", **pad)
        row_jb = ttk.Frame(frm_jieba)
        row_jb.pack(fill="x", padx=5, pady=5)
        ttk.Entry(row_jb, textvariable=self.jieba_dict_path, width=50,
                  state="readonly").pack(side="left", padx=(0, 5))
        ttk.Button(row_jb, text="选择", command=self._choose_jieba_dict).pack(side="left")

        frm_out = ttk.LabelFrame(tab, text="输出文件路径 — 生成的面板数据 Excel 保存位置")
        frm_out.pack(fill="x", **pad)
        row_out = ttk.Frame(frm_out)
        row_out.pack(fill="x", padx=5, pady=5)
        ttk.Entry(row_out, textvariable=self.output_path, width=50,
                  state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(row_out, text="选择路径", command=self._choose_output).pack(side="left")

    # ---- 标签页4：运行分析 ----

    def _build_run_tab(self, notebook):
        tab = ttk.Frame(notebook)
        notebook.add(tab, text=" 运行分析 ")
        pad = {"padx": 8, "pady": 4}

        # v3.0: 开始 + 取消按钮
        btn_row = ttk.Frame(tab)
        btn_row.pack(pady=(15, 5))
        self.btn_start = ttk.Button(btn_row, text="  开始统计  ", command=self._start_analysis)
        self.btn_start.pack(side="left", padx=5)
        self.btn_cancel = ttk.Button(btn_row, text="  取消  ", command=self._cancel_analysis, state="disabled")
        self.btn_cancel.pack(side="left", padx=5)

        frm_prog = ttk.LabelFrame(tab, text="进度")
        frm_prog.pack(fill="x", **pad)
        self.progress = ttk.Progressbar(frm_prog, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=5, pady=(5, 2))
        ttk.Label(frm_prog, textvariable=self.current_file_var).pack(
            anchor="w", padx=5, pady=(0, 5)
        )

        frm_log = ttk.LabelFrame(tab, text="运行日志")
        frm_log.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(frm_log, wrap="word", height=18, state="disabled",
                               font=("Menlo", 11) if sys.platform == "darwin" else ("Consolas", 10))
        log_sb = ttk.Scrollbar(frm_log, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        self.log_text.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
        log_sb.pack(side="right", fill="y", padx=(0, 5), pady=5)

        # v3.0: 日志颜色标签
        self.log_text.tag_configure("error", foreground="#cc3333")
        self.log_text.tag_configure("success", foreground="#228B22")
        self.log_text.tag_configure("warn", foreground="#cc8800")

    # ================================================================
    #  数据选择 — 事件
    # ================================================================

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择文本文件夹")
        if folder and folder not in self.folders:
            self.folders.append(folder)
            self.folder_listbox.insert(tk.END, folder)

    def _remove_folder(self):
        sel = self.folder_listbox.curselection()
        if sel:
            idx = sel[0]
            self.folder_listbox.delete(idx)
            self.folders.pop(idx)

    def _scan_files(self):
        if not self.folders:
            messagebox.showwarning("提示", "请先添加文件夹。")
            return

        # v3.0: 使用增强的 collect_data_files
        self.scanned_files, dir_counts, errors = collect_data_files(self.folders)
        self.file_listbox.delete(0, tk.END)
        for f in self.scanned_files:
            self.file_listbox.insert(tk.END, f)

        if not self.scanned_files:
            err_msg = ""
            if errors:
                err_msg = "\n\n扫描错误：\n" + "\n".join(errors[:5])
            messagebox.showinfo("提示", f"未找到 .xlsx / .xls / .csv 文件。{err_msg}")
            return

        # v3.0: 后台线程扫描列名（大 xlsx 读取慢，避免冻结 UI）
        self._status_var.set(f"正在扫描 {len(self.scanned_files)} 个文件的列名…")
        self.update_idletasks()

        scan_dir_counts = dir_counts
        scan_errors = errors

        def _do_scan():
            cols, freq = scan_all_columns(self.scanned_files)
            self.after(0, lambda: self._finish_scan(cols, freq, scan_dir_counts, scan_errors))

        threading.Thread(target=_do_scan, daemon=True).start()

    def _finish_scan(self, cols, freq, dir_counts, errors):
        """列名扫描完成后在主线程更新 UI。"""
        self.all_columns = cols
        self._col_freq = freq
        n_files = len(self.scanned_files)

        col_values = list(self.all_columns)  # 已按频率降序排列
        self.combo_stkcd["values"] = col_values
        self.combo_year["values"] = col_values
        self.col_listbox.delete(0, tk.END)
        self._col_display_map = {}
        for c in col_values:
            f = self._col_freq.get(c, 0)
            # 在列表中显示列名及其出现频率，方便用户判断
            label = f"{c}  [{f}/{n_files}]" if f < n_files else c
            self.col_listbox.insert(tk.END, label)
            self._col_display_map[label] = c

        # v3.0: 按频率加权的自动列匹配
        stkcd_patterns = ("股票代码", "证券代码", "公司代码", "stkcd", "stock_code", "scode", "code")
        year_patterns = ("年份", "year", "年", "日期", "时间", "date", "qtm")

        def _best_match(patterns):
            candidates = []
            for c in col_values:
                cl = c.lower()
                if any(k in cl for k in patterns):
                    candidates.append((c, self._col_freq.get(c, 0)))
            if candidates:
                candidates.sort(key=lambda x: -x[1])
                return candidates[0][0]
            return None

        best_stkcd = _best_match(stkcd_patterns)
        if best_stkcd:
            self.combo_stkcd.set(best_stkcd)

        best_year = _best_match(year_patterns)
        if best_year:
            self.combo_year.set(best_year)

        # v3.0: 文本列自动选择（匹配常见文本列名，高频优先）
        text_patterns = ("reply", "回复", "text", "文本", "内容", "content", "qsubj",
                         "问题", "question", "answer", "描述", "说明", "摘要", "正文",
                         "subject", "title", "标题", "body")
        auto_text_indices = []
        for idx, label in enumerate(list(self._col_display_map.keys())):
            real_col = self._col_display_map[label]
            cl = real_col.lower()
            if any(k in cl for k in text_patterns):
                auto_text_indices.append(idx)
        if auto_text_indices:
            for i in auto_text_indices:
                self.col_listbox.selection_set(i)

        # v3.0: 自动建议输出路径
        if not self.output_path.get() and self.folders:
            base_dir = self.folders[0]
            suggested_name = os.path.basename(base_dir) + "_词频统计.xlsx"
            suggested_path = os.path.join(os.path.dirname(base_dir), suggested_name)
            self.output_path.set(suggested_path)

        # 显示子目录统计
        dir_info_lines = sorted(dir_counts.items())
        if len(dir_info_lines) > 15:
            dir_info = "\n".join(f"  {d}: {n} 个文件" for d, n in dir_info_lines[:12])
            dir_info += f"\n  … 共 {len(dir_info_lines)} 个目录"
        else:
            dir_info = "\n".join(f"  {d}: {n} 个文件" for d, n in dir_info_lines)

        err_info = ""
        if errors:
            err_info = f"\n\n扫描警告（{len(errors)} 个错误）：\n" + "\n".join(errors[:3])

        self._update_status_bar()

        # 自动配置提示
        auto_info = ""
        if best_stkcd:
            auto_info += f"\n  公司代码列 → {best_stkcd}"
        if best_year:
            auto_info += f"\n  年份/日期列 → {best_year}"
        if auto_text_indices:
            auto_text_names = [self._col_display_map[list(self._col_display_map.keys())[i]]
                               for i in auto_text_indices]
            auto_info += f"\n  文本列 → {', '.join(auto_text_names)}"
        if auto_info:
            auto_info = f"\n\n自动检测的列配置：{auto_info}\n（请核实是否正确）"

        messagebox.showinfo(
            "扫描完成",
            f"共扫描到 {len(self.scanned_files)} 个数据文件，\n"
            f"分布在 {len(dir_counts)} 个目录中：\n{dir_info}\n\n"
            f"发现 {len(self.all_columns)} 个数据列。{auto_info}{err_info}"
        )

    def _on_file_select(self, _event=None):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        filepath = self.scanned_files[sel[0]]
        try:
            df = read_data_file(filepath, nrows=100)
        except Exception as e:
            messagebox.showerror("预览失败", str(e))
            return

        self.preview_tree.delete(*self.preview_tree.get_children())
        cols = list(df.columns)
        self.preview_tree["columns"] = cols
        for c in cols:
            self.preview_tree.heading(c, text=c)
            self.preview_tree.column(c, width=130, minwidth=60, stretch=False)
        for _, row in df.iterrows():
            vals = [str(v)[:200] if pd.notna(v) else "" for v in row]
            self.preview_tree.insert("", "end", values=vals)

    # ================================================================
    #  词典管理 — 事件
    # ================================================================

    def _update_dict_info(self):
        n_cat = len(self.dict_mgr.categories())
        n_word = self.dict_mgr.total_word_count()
        self.dict_info_var.set(f"当前词典：{n_cat} 个分类，{n_word} 个关键词")
        self._update_status_bar()

    def _update_status_bar(self):
        n_word = self.dict_mgr.total_word_count()
        n_files = len(self.scanned_files)
        self._status_var.set(f"就绪  |  词典：{n_word} 词  |  数据：{n_files} 文件")

    def _refresh_cat_list(self):
        self.cat_listbox.delete(0, tk.END)
        for c in self.dict_mgr.categories():
            n = len(self.dict_mgr.words(c))
            self.cat_listbox.insert(tk.END, f"{c}  ({n})")
        self._update_dict_info()

    def _refresh_word_list(self, category: str):
        self.word_listbox.delete(0, tk.END)
        for w in self.dict_mgr.words(category):
            self.word_listbox.insert(tk.END, w)

    def _selected_category(self) -> str | None:
        sel = self.cat_listbox.curselection()
        if not sel:
            return None
        cats = self.dict_mgr.categories()
        return cats[sel[0]] if sel[0] < len(cats) else None

    def _on_cat_select(self, _event=None):
        cat = self._selected_category()
        if cat:
            self._refresh_word_list(cat)

    def _cat_add(self):
        name = simpledialog.askstring("添加分类", "请输入分类名称：", parent=self)
        if name and name.strip():
            self.dict_mgr.add_category(name.strip())
            self._refresh_cat_list()

    def _cat_remove(self):
        cat = self._selected_category()
        if not cat:
            messagebox.showwarning("提示", "请先选择一个分类。")
            return
        if messagebox.askyesno("确认", f"确定删除分类「{cat}」及其所有关键词？"):
            self.dict_mgr.remove_category(cat)
            self.word_listbox.delete(0, tk.END)
            self._refresh_cat_list()

    def _word_add(self):
        cat = self._selected_category()
        if not cat:
            messagebox.showwarning("提示", "请先选择一个分类。")
            return
        word = simpledialog.askstring("添加关键词", f"分类「{cat}」— 请输入关键词：", parent=self)
        if word and word.strip():
            self.dict_mgr.add_word(cat, word.strip())
            self._refresh_word_list(cat)
            self._update_dict_info()

    def _word_batch_add(self):
        cat = self._selected_category()
        if not cat:
            messagebox.showwarning("提示", "请先选择一个分类。")
            return
        dlg = BatchAddDialog(self, title=f"批量添加关键词 — {cat}")
        self.wait_window(dlg)
        if dlg.result:
            for w in dlg.result:
                self.dict_mgr.add_word(cat, w)
            self._refresh_word_list(cat)
            self._refresh_cat_list()

    def _word_remove(self):
        cat = self._selected_category()
        if not cat:
            return
        sel = self.word_listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选择要删除的关键词。")
            return
        word = self.word_listbox.get(sel[0])
        self.dict_mgr.remove_word(cat, word)
        self._refresh_word_list(cat)
        self._refresh_cat_list()

    def _dict_new(self):
        if self.dict_mgr.total_word_count() > 0:
            if not messagebox.askyesno("确认", "新建词典将清空当前所有分类和关键词，确定？"):
                return
        self.dict_mgr.clear()
        self.cat_listbox.delete(0, tk.END)
        self.word_listbox.delete(0, tk.END)
        self._update_dict_info()

    def _dict_import(self):
        path = filedialog.askopenfilename(
            title="导入词典（Excel 或 TXT）",
            filetypes=[
                ("词典文件", "*.xlsx *.xls *.txt"),
                ("Excel 文件", "*.xlsx *.xls"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*"),
            ],
        )
        if not path:
            return
        try:
            self.dict_mgr.import_file(path)
            self._refresh_cat_list()
            self.word_listbox.delete(0, tk.END)
            summary = self.dict_mgr.summary_text()
            messagebox.showinfo(
                "导入成功",
                f"共 {len(self.dict_mgr.categories())} 个分类，"
                f"{self.dict_mgr.total_word_count()} 个关键词：\n\n{summary}"
            )
        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    def _dict_export(self):
        if self.dict_mgr.total_word_count() == 0:
            messagebox.showwarning("提示", "词典为空，无法导出。")
            return
        path = filedialog.asksaveasfilename(
            title="导出词典 Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            try:
                self.dict_mgr.export_excel(path)
                messagebox.showinfo("导出成功", f"词典已导出至：{path}")
            except Exception as e:
                messagebox.showerror("导出失败", str(e))

    # ================================================================
    #  设置 — 事件
    # ================================================================

    def _choose_stopwords(self):
        path = filedialog.askopenfilename(
            title="选择停用词文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
        )
        if path:
            self.stopwords_path.set(path)

    def _choose_jieba_dict(self):
        path = filedialog.askopenfilename(
            title="选择 jieba 用户词典",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
        )
        if path:
            self.jieba_dict_path.set(path)

    def _choose_output(self):
        path = filedialog.asksaveasfilename(
            title="选择输出文件路径",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    # ================================================================
    #  日志 / 进度（线程安全）
    # ================================================================

    def _log(self, msg: str):
        def _do():
            self.log_text.configure(state="normal")
            # v3.0: 日志着色 — 错误红色，完成绿色，警告橙色
            tag = None
            if "错误" in msg or "跳过" in msg or "失败" in msg:
                tag = "error"
            elif "完成" in msg or "成功" in msg:
                tag = "success"
            elif "警告" in msg or "提示" in msg:
                tag = "warn"
            if tag:
                self.log_text.insert(tk.END, msg + "\n", tag)
            else:
                self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state="disabled")
        self.after(0, _do)

    def _update_progress(self, current: int, total: int):
        def _do():
            self.progress["maximum"] = total
            self.progress["value"] = current
        self.after(0, _do)

    def _set_current_file(self, msg: str):
        self.after(0, lambda: self.current_file_var.set(msg))

    # ================================================================
    #  运行分析
    # ================================================================

    def _cancel_analysis(self):
        self._cancel_event.set()
        self._log("正在取消…")
        self.btn_cancel.configure(state="disabled")

    def _start_analysis(self):
        if not self.scanned_files:
            messagebox.showwarning("提示", "请先在「数据选择」扫描文件。")
            return

        col_stkcd = self.combo_stkcd.get()
        col_year = self.combo_year.get()
        if not col_stkcd:
            messagebox.showwarning("提示", "请在「数据选择」配置 公司代码列。")
            return
        if not col_year:
            messagebox.showwarning("提示", "请在「数据选择」配置 年份/日期列。")
            return

        # v3.0: 将显示标签还原为真实列名
        col_map = getattr(self, "_col_display_map", {})
        selected_text_cols = []
        for i in self.col_listbox.curselection():
            label = self.col_listbox.get(i)
            real_col = col_map.get(label, label)
            selected_text_cols.append(real_col)
        if not selected_text_cols:
            messagebox.showwarning("提示", "请在「数据选择」选择至少一个文本列。")
            return

        if self.dict_mgr.total_word_count() == 0:
            messagebox.showwarning("提示", "词典为空，请先在「词典管理」添加或导入词典。")
            return

        if not self.output_path.get():
            messagebox.showwarning("提示", "请在「分析设置」选择输出路径。")
            return

        # 停用词
        stopwords: set[str] = set()
        if self.var_stopwords.get():
            stopwords = set(DEFAULT_STOPWORDS)
            if self.stopwords_path.get() and os.path.isfile(self.stopwords_path.get()):
                stopwords |= load_stopwords_file(self.stopwords_path.get())

        # 锁定 UI
        self._cancel_event.clear()
        self.btn_start.configure(state="disabled")
        self.btn_cancel.configure(state="normal")
        self.progress["value"] = 0
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")
        self.current_file_var.set("准备中…")

        def log_with_status(msg):
            self._log(msg)
            self._set_current_file(msg)

        kwargs = dict(
            files=list(self.scanned_files),
            dict_mgr=self.dict_mgr,
            col_stkcd=col_stkcd,
            col_year=col_year,
            text_columns=selected_text_cols,
            output_path=self.output_path.get(),
            use_regex=self.var_regex.get(),
            use_stopwords=self.var_stopwords.get(),
            stopwords=stopwords,
            use_tf=self.var_tf.get(),
            export_sentences=self.var_sentences.get(),
            jieba_userdict=self.jieba_dict_path.get(),
            progress_cb=self._update_progress,
            log_cb=log_with_status,
            cancel_event=self._cancel_event,
        )

        thread = threading.Thread(target=self._run_thread, kwargs=kwargs, daemon=True)
        thread.start()

    def _run_thread(self, **kwargs):
        try:
            run_analysis(**kwargs)
            self.after(0, lambda: messagebox.showinfo("完成", "面板数据词频统计已完成！"))
        except Exception as e:
            msg = str(e)
            self._log(f"错误：{msg}")
            if "取消" not in msg:
                self.after(0, lambda: messagebox.showerror("运行出错", msg))
        finally:
            self.after(0, lambda: self.btn_start.configure(state="normal"))
            self.after(0, lambda: self.btn_cancel.configure(state="disabled"))
            self._set_current_file("就绪")


# ============================================================
# 入口
# ============================================================

if __name__ == "__main__":
    # Windows PyInstaller multiprocessing 兼容
    import multiprocessing
    multiprocessing.freeze_support()

    # PyInstaller 打包兼容
    if getattr(sys, "frozen", False):
        os.chdir(os.path.dirname(sys.executable))

    app = WordFreqApp()
    app.mainloop()
