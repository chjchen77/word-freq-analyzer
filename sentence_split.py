"""中文年报分句 —— 全项目唯一权威实现。

`word_freq_analyzer`（产出命中句，即比值的分子）与 `count_mda_sentences`
（产出 MD&A 句子总数，即分母）必须使用完全相同的切分口径，否则 ratio 类
变量的分子分母不可比。历史上两处各存一份副本并逐渐漂移，实测 40/40 份年报
句数不一致，故统一收敛到本模块。

仅依赖标准库 re，可在无 pandas / tkinter 的服务器环境下独立使用。
"""
from __future__ import annotations

import re

MAX_SENT_LEN = 200
MIN_SENT_LEN = 8

# 页面装饰（页码行、运行页眉），连同其上下空行整块塌缩为单个换行。
# 年报跨页处形如 "环境规\n \n35 \n \n划署"，若只删装饰文字而留下空行，
# 反而凭空造出段落边界，把一句话拦腰截断，故必须整块消化。
# 页眉形如 "××股份有限公司 2021 年年度报告全文"，混入正文时会同时造成
# 语义断裂与词语截断（"国际通用"→"际通用"），进而影响关键词命中。
# 判别关键：整行以「报告/报告全文」结尾。正文中提及年报的句子（如
# "…的年度报告，检查了公司的会计政策…"）不以其结尾，因此不会被误删。
# 注意：内部一律用 [ \t　]* 而非 \s*，避免跨越换行误吞正文。
_PAGE_NUM = r"第?[ \t　]*\d{1,4}[ \t　]*页?(?:[ \t　]*/[ \t　]*\d{1,4}[ \t　]*)?"
_RUN_HEADER = r".{0,30}?\d{4}[ \t　]*年?[ \t　]*(?:年度|半年度|度)?[ \t　]*报告[ \t　]*(?:全文)?"
_PAGE_BLOCK_RE = re.compile(
    r"\n(?:[ \t　]*\n)*"
    rf"(?:[ \t　]*(?:{_PAGE_NUM}|{_RUN_HEADER})[ \t　]*\n(?:[ \t　]*\n)*)+"
)
# 段落边界：两个及以上换行，其间允许只有空白。年报里普遍是 "\n \n"（夹空格），
# 旧写法 \n{2,} 要求换行严格相邻，匹配不到，导致整段被粘成数百字巨块。
_PARA_BREAK_RE = re.compile(r"\n[ \t　]*\n\s*")
# 行内页眉：PDF 抽取时页眉与正文并入同一行，行尾不再是「报告」，
# 上面按整行匹配的规则抓不到。这里要求「完整公司名 + 年份 + 报告全文」，
# 其中「全文」是页眉的强特征——正文提及年报（如"审议通过了公司 2007 年
# 年度报告及其摘要"）不会带此后缀，故不会误删。
_INLINE_HEADER_RE = re.compile(
    r"[一-龥]{2,20}(?:股份)?有限公司[ \t　]*\d{4}[ \t　]*年[ \t　]*年?度?报告全文"
)
_SENT_END_RE = re.compile(r"[。！？!?]")
_TRIM_CHARS = "；;：:，,、 　"


def normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\x00", "")


def is_table_like(s: str) -> bool:
    """疑似表格/数字块：数字与分隔符密度高，或中文字符过少。

    调用方必须传入同一形态的字符串（本模块内统一为 strip 后的原串，
    保留串内空格），否则空格计入非中文比率会导致两侧判定分叉。
    """
    if not s:
        return True
    non_cn = sum(c.isdigit() or c in ",.%-—、|/ " for c in s)
    cn = sum("一" <= c <= "龥" for c in s)
    return non_cn / len(s) > 0.45 or cn < 8


def _resplit_long(seg: str, max_len: int) -> list[str]:
    """超长段落按次级标点逐级细分，避免整块表格/长段落成为单个"句子"。"""
    for punct in ("；", ";", "，", ","):
        if punct not in seg:
            continue
        buf, out = "", []
        for piece in seg.split(punct):
            cand = f"{buf}{punct}{piece}" if buf else piece
            if len(cand) > max_len and buf:
                out.append(buf)
                buf = piece
            else:
                buf = cand
        if buf:
            out.append(buf)
        if all(len(x) <= max_len for x in out):
            return out
        deeper: list[str] = []
        for x in out:
            deeper.extend(_resplit_long(x, max_len) if len(x) > max_len else [x])
        return deeper
    # 无标点可切（典型为财务表格数字块），按长度硬切兜底
    return [seg[i:i + max_len] for i in range(0, len(seg), max_len)]


def split_sentences(text: str, max_len: int = MAX_SENT_LEN,
                    drop_tables: bool = True, compact: bool = False) -> list[str]:
    """切分句子，专门处理年报 PDF 转文字后大量断行的问题。

    1. 统一换行符；
    2. 纯页码行连同周边空行塌缩为单换行（跨页续行得以正确粘合）；
    3. 段落边界（连续换行，容忍其间空白）替换为句末标点；
    4. 剩余单个换行 = PDF 排版断行，直接删除，两边文字合并；
    5. 以 。！？!? 为句子边界；
    6. 超过 max_len 的长段按次级标点继续细分；
    7. 过滤 < MIN_SENT_LEN 的碎片，并剔除疑似表格块。

    去留判定一律基于 strip 后的原串，`compact` 只影响最终返回形态，
    不影响句子数量，从而保证分子与分母口径严格一致。
    """
    text = normalize_newlines(text)
    text = _PAGE_BLOCK_RE.sub("\n", text)
    text = _INLINE_HEADER_RE.sub("", text)
    text = _PARA_BREAK_RE.sub("。", text)
    text = text.replace("\n", "")

    segments: list[str] = []
    for part in _SENT_END_RE.split(text):
        part = part.strip()
        if not part:
            continue
        segments.extend(
            _resplit_long(part, max_len) if len(part) > max_len else [part]
        )

    out: list[str] = []
    for seg in segments:
        seg = seg.strip()
        if len(seg) < MIN_SENT_LEN:
            continue
        if drop_tables and is_table_like(seg):
            continue
        out.append(re.sub(r"\s+", "", seg).strip(_TRIM_CHARS) if compact else seg)
    return out
