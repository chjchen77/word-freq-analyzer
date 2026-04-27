from __future__ import annotations

import hashlib
import json
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import pandas as pd

DEFAULT_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
DEFAULT_MODEL = "qwen-plus"

# 单句最大字符数，超过则截断后再送给 LLM（节省 token，防止超出上下文）
_MAX_SENTENCE_CHARS = 500

# 单次 API 请求超时秒数
_REQUEST_TIMEOUT = 60.0

# 重试退避基础秒数（第 n 次重试等待 n * _RETRY_BACKOFF_BASE 秒）
_RETRY_BACKOFF_BASE = 1.0

# 429 限速专项重试上限（独立于普通错误重试，指数退避最长 60s）
_MAX_RATE_LIMIT_RETRIES = 10

# 增量缓存：每成功 N 条写一次盘，防止长时间运行中途崩溃丢失进度
_CACHE_SAVE_INTERVAL = 50

TIME_ORIENTATION_VALUES = ("过去", "当前", "未来", "混合", "不明显")
VOICE_VALUES = ("主动", "被动", "不明显")
SENTENCE_TYPE_VALUES = ("回顾总结", "现状描述", "未来规划", "风险提示", "措施行动", "结果成效", "其他")
CERTAINTY_VALUES = ("高", "中", "低")
QUANT_VALUES = ("定量", "半定量", "定性")
TONE_VALUES = ("积极", "消极", "其他")
CONFIDENCE_VALUES = ("高", "中", "低")

RESULT_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "time_orientation": {"type": "string", "enum": list(TIME_ORIENTATION_VALUES)},
        "voice": {"type": "string", "enum": list(VOICE_VALUES)},
        "sentence_type": {"type": "string", "enum": list(SENTENCE_TYPE_VALUES)},
        "certainty": {"type": "string", "enum": list(CERTAINTY_VALUES)},
        "quantitativeness": {"type": "string", "enum": list(QUANT_VALUES)},
        "tone": {"type": "string", "enum": list(TONE_VALUES)},
        "confidence": {"type": "string", "enum": list(CONFIDENCE_VALUES)},
        "reason": {"type": "string"},
    },
    "required": [
        "time_orientation",
        "voice",
        "sentence_type",
        "certainty",
        "quantitativeness",
        "tone",
        "confidence",
        "reason",
    ],
}

# 400 BadRequest 时才触发 json_schema → json_object 降级，
# 这些关键词用于识别"模型不支持该 response_format"类错误。
_FORMAT_UNSUPPORTED_KEYWORDS = (
    "json_schema",
    "response_format",
    "invalid_request",
    "unsupported",
    "not support",
    "not supported",
)

# _normalize_result：LLM 返回的 JSON 中至少要包含其中一个维度字段才视为有效响应
_EXPECTED_RESULT_KEYS = frozenset(
    {"time_orientation", "voice", "sentence_type", "certainty", "quantitativeness", "tone"}
)

# 缓存命中校验：所有字段都存在时才使用缓存，防止格式不完整的旧缓存污染结果
_REQUIRED_CACHE_KEYS = frozenset({
    "LLM时间指向", "LLM语态", "LLM句子类型",
    "LLM确定性", "LLM量化属性", "LLM语气语调", "LLM分析状态",
})


def _clean_json_string(text: str) -> str:
    text = re.sub(r"```json\s*", "", text)
    text = re.sub(r"```\s*", "", text)
    return text.strip()


def _safe_int(value: Any, default: int) -> int:
    """解析正整数，最小值为 1。用于 max_workers、max_retries 等只能为正数的参数。"""
    try:
        return max(1, int(value))
    except (TypeError, ValueError):
        return default


def _safe_nonneg_int(value: Any, default: int) -> int:
    """解析非负整数，允许 0。max_sentences=0 表示「无上限」，必须用此函数而非 _safe_int。"""
    try:
        return max(0, int(value))
    except (TypeError, ValueError):
        return default


def _is_format_unsupported_error(exc: Exception) -> bool:
    """判断异常是否属于"模型不支持 json_schema response_format"类 400 错误。
    避免把网络超时、限速（429）等错误误判为格式不支持而触发降级。
    """
    exc_type = type(exc).__name__
    exc_msg = str(exc).lower()

    # openai 库的 BadRequestError / InvalidRequestError 对应 HTTP 400
    if exc_type in ("BadRequestError", "InvalidRequestError"):
        return True

    # 某些 SDK 版本用 APIStatusError，通过状态码 + 关键词双重判断
    if "400" in exc_msg and any(kw in exc_msg for kw in _FORMAT_UNSUPPORTED_KEYWORDS):
        return True

    return False


@dataclass
class LLMAnalyzerConfig:
    api_key: str
    model: str = DEFAULT_MODEL
    base_url: str = DEFAULT_BASE_URL
    max_workers: int = 4
    max_sentences: int = 500
    max_retries: int = 2
    cache_path: str | None = None
    system_prompt: str | None = None  # None = 使用内置默认提示词

    @classmethod
    def from_inputs(
        cls,
        *,
        api_key: str = "",
        model: str = DEFAULT_MODEL,
        base_url: str = DEFAULT_BASE_URL,
        max_workers: int = 4,
        max_sentences: int = 500,
        max_retries: int = 2,
        cache_path: str | None = None,
        system_prompt: str | None = None,
    ) -> "LLMAnalyzerConfig":
        resolved_key = (
            api_key.strip()
            or os.getenv("DASHSCOPE_API_KEY", "").strip()
            or os.getenv("QWEN_API_KEY", "").strip()
        )
        resolved_base_url = (
            (base_url or os.getenv("QWEN_BASE_URL", DEFAULT_BASE_URL)).strip() or DEFAULT_BASE_URL
        )
        resolved_model = (
            (model or os.getenv("QWEN_MODEL", DEFAULT_MODEL)).strip() or DEFAULT_MODEL
        )
        return cls(
            api_key=resolved_key,
            model=resolved_model,
            base_url=resolved_base_url.rstrip("/"),
            max_workers=_safe_int(max_workers, 4),
            max_sentences=_safe_nonneg_int(max_sentences, 500),  # 0 = 无上限，不能用 _safe_int
            max_retries=_safe_int(max_retries, 2),
            cache_path=cache_path,
            system_prompt=system_prompt or None,
        )


class QwenSentenceAnalyzer:
    SYSTEM_PROMPT = """你是中文财经文本分析助手。请分析"包含目标关键词的句子"。

请严格按照给定 JSON 结构返回，不要输出任何解释性文字。
分类规则：
1. time_orientation：
- 过去：主要描述已经发生、已经完成、历史回顾。
- 当前：主要描述现状、持续状态、截至目前情况。
- 未来：主要描述计划、将来安排、目标、预期。
- 混合：同一句同时包含明显的过去/当前/未来信息。
- 不明显：时间指向无法可靠判断。
2. voice：
- 主动：主体主动实施、推进、开展、完成某事项。
- 被动：出现"被""获""受到""由…完成"等受动表达。
- 不明显：无法可靠判断。
3. sentence_type：
- 回顾总结、现状描述、未来规划、风险提示、措施行动、结果成效、其他。
4. certainty：
- 高：措辞明确、事实性强。
- 中：有一定判断但不完全绝对。
- 低：明显存在预计、可能、或许、拟、计划等不确定表达。
5. quantitativeness：
- 定量：有明确数字、比例、金额、数量。
- 半定量：没有精确数字，但有"显著""大幅""较快"等程度量化表达。
- 定性：纯描述性表达。
6. tone：
- 积极：整体语气偏正向、乐观、肯定、表扬、成效明显。
- 消极：整体语气偏负向、风险、压力、困难、亏损、处罚、下滑。
- 其他：中性陈述、难以判断，或正负都不明显。
7. confidence（你对本次所有维度判断的综合置信度）：
- 高：句子含义明确，各维度判断依据充分，无歧义。
- 中：部分维度存在一定模糊性，但整体判断有把握。
- 低：句子语义模糊、截断或缺乏上下文，判断较难确定。
8. reason：
- 用一句简短中文说明最关键的判断依据，尽量引用句中触发词，不超过80字。

半定量与定性的区分规则（重要）：
- 定量：含明确数字/百分比/金额/数量，如"同比增长15%"、"投入3亿元"。
- 半定量：无精确数字，但含程度副词表达相对量，如"大幅下降"、"显著提升"、"较快增长"、"明显改善"。
- 定性：纯描述性陈述，如"积极推进"、"持续关注"、"高度重视"，不含任何量化表达。"""

    def __init__(
        self,
        config: LLMAnalyzerConfig,
        *,
        log_cb: Callable[[str], None] | None = None,
        cancel_event: threading.Event | None = None,
    ):
        if not config.api_key:
            raise ValueError("未提供 API Key。请在界面中填写，或设置环境变量 DASHSCOPE_API_KEY。")
        self.config = config
        self.log_cb = log_cb
        self.cancel_event = cancel_event
        self._local = threading.local()
        self._cache_lock = threading.Lock()
        self._cache = self._load_cache(config.cache_path)
        # 记录本次运行是否已触发过 json_schema → json_object 降级，
        # 降级后后续请求直接使用 json_object，避免重复触发 400 错误
        self._use_json_object_fallback = False
        self._fallback_lock = threading.Lock()
        # 实际使用的系统提示词：优先用用户自定义，否则用内置默认
        self._system_prompt: str = config.system_prompt or self.SYSTEM_PROMPT
        # 当使用自定义提示词时，在缓存 key 中混入其哈希，避免与默认提示词的缓存冲突
        self._prompt_hash_suffix: str = (
            ""
            if config.system_prompt is None
            else (":" + hashlib.sha1(config.system_prompt.encode("utf-8")).hexdigest()[:8])
        )

    def _log(self, msg: str) -> None:
        if self.log_cb:
            self.log_cb(msg)

    def _is_cancelled(self) -> bool:
        return self.cancel_event is not None and self.cancel_event.is_set()

    def _get_client(self):
        """每个线程维护独立的 OpenAI client（线程安全），并设置请求超时。"""
        try:
            from openai import OpenAI
        except ImportError as exc:
            raise ValueError(
                "缺少 openai 依赖。请先执行 `pip install openai`。"
            ) from exc

        client = getattr(self._local, "client", None)
        if client is None:
            client = OpenAI(
                api_key=self.config.api_key,
                base_url=self.config.base_url,
                timeout=_REQUEST_TIMEOUT,  # 修复：防止线程因 API 挂起而永远等待
            )
            self._local.client = client
        return client

    def _record_key(self, record: dict[str, Any], model: str) -> str:
        payload = {
            "model": model + self._prompt_hash_suffix,
            "keyword": str(record.get("命中关键词", "")).strip(),
            "category": str(record.get("分类", "")).strip(),
            "sentence": str(record.get("命中句子", "")).strip(),
        }
        raw = json.dumps(payload, ensure_ascii=False, sort_keys=True)
        return hashlib.sha1(raw.encode("utf-8")).hexdigest()

    def _load_cache(self, cache_path: str | None) -> dict[str, dict[str, Any]]:
        if not cache_path:
            return {}
        path = Path(cache_path)
        try:
            with path.open("r", encoding="utf-8") as f:
                data = json.load(f)
            # JSON 合法但不是 dict（如 list），返回空缓存
            return data if isinstance(data, dict) else {}
        except FileNotFoundError:
            # 文件尚不存在属于预期情况（首次运行），静默返回
            return {}
        except Exception as exc:
            # JSON 损坏、权限不足等非预期情况，给出明确警告
            self._log(f"警告：LLM 缓存文件读取失败（{exc}），将以空缓存启动。")
            return {}

    def _save_cache(self) -> None:
        if not self.config.cache_path:
            return
        path = Path(self.config.cache_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(path.suffix + ".tmp")
        # 先在锁内做快照，再在锁外写盘，避免写盘期间阻塞其他线程
        with self._cache_lock:
            cache_snapshot = dict(self._cache)
        try:
            with temp_path.open("w", encoding="utf-8") as f:
                json.dump(cache_snapshot, f, ensure_ascii=False, indent=2)
            temp_path.replace(path)
        except Exception as exc:
            self._log(f"警告：LLM 缓存保存失败：{exc}")

    @staticmethod
    def _normalize_value(value: Any, allowed: tuple[str, ...], default: str) -> str:
        value_str = str(value or "").strip()
        if value_str in allowed:
            return value_str
        return default

    def _normalize_result(self, data: dict[str, Any]) -> dict[str, str]:
        # 修复：空响应或缺失所有维度字段时返回失败，避免将全默认值误写为"成功"
        if not isinstance(data, dict) or not any(k in data for k in _EXPECTED_RESULT_KEYS):
            return _failed_result("LLM返回了空响应或格式不符合预期（缺少所有维度字段）")
        return {
            "LLM时间指向": self._normalize_value(data.get("time_orientation"), TIME_ORIENTATION_VALUES, "不明显"),
            "LLM语态": self._normalize_value(data.get("voice"), VOICE_VALUES, "不明显"),
            "LLM句子类型": self._normalize_value(data.get("sentence_type"), SENTENCE_TYPE_VALUES, "其他"),
            "LLM确定性": self._normalize_value(data.get("certainty"), CERTAINTY_VALUES, "中"),
            "LLM量化属性": self._normalize_value(data.get("quantitativeness"), QUANT_VALUES, "定性"),
            "LLM语气语调": self._normalize_value(data.get("tone"), TONE_VALUES, "其他"),
            "LLM置信度": self._normalize_value(data.get("confidence"), CONFIDENCE_VALUES, "中"),
            "LLM判定依据": str(data.get("reason", "") or "").strip()[:80],
            "LLM分析状态": "成功",
            "LLM分析错误": "",
        }

    def _build_prompt(self, record: dict[str, Any]) -> str:
        keyword = str(record.get("命中关键词", "")).strip()
        category = str(record.get("分类", "")).strip()
        sentence = str(record.get("命中句子", "")).strip()
        # 修复：超长句子截断，避免浪费 token 并防止超出模型上下文
        if len(sentence) > _MAX_SENTENCE_CHARS:
            sentence = sentence[:_MAX_SENTENCE_CHARS] + "…（已截断）"
        return (
            "请分析下列中文年报句子，并按给定 JSON 结构返回结果。\n"
            f"目标关键词：{keyword or '未提供'}\n"
            f"关键词分类：{category or '未提供'}\n"
            f"句子：{sentence}\n\n"
            "只返回 JSON，不要输出任何额外文字。"
        )

    def _request_once(self, record: dict[str, Any]) -> dict[str, str]:
        """发起一次 API 请求并返回标准化结果。

        优先使用 json_schema 严格模式（确保字段完整）；
        仅当模型明确返回 400 BadRequest（格式不支持）时，才降级为 json_object 模式，
        并记录降级状态避免后续请求重复触发 400。
        网络超时、限速（429）等错误直接抛出，由上层重试逻辑处理。
        """
        client = self._get_client()
        request_kwargs: dict[str, Any] = {
            "model": self.config.model,
            "messages": [
                {"role": "system", "content": self._system_prompt},
                {"role": "user", "content": self._build_prompt(record)},
            ],
            "temperature": 0,
            "max_tokens": 300,
        }

        # 如果本次会话已确认模型不支持 json_schema，直接使用 json_object
        with self._fallback_lock:
            already_fallback = self._use_json_object_fallback

        if not already_fallback:
            try:
                response = client.chat.completions.create(
                    response_format={
                        "type": "json_schema",
                        "json_schema": {
                            "name": "annual_report_sentence_analysis",
                            "strict": True,
                            "schema": RESULT_SCHEMA,
                        },
                    },
                    **request_kwargs,
                )
            except Exception as exc:
                # 修复：只有明确是"格式不支持"的 400 错误才降级；
                # 其他异常（超时、限速等）直接抛出，由重试逻辑处理
                if _is_format_unsupported_error(exc):
                    # 修复：仅在首次触发降级时记录日志，防止多线程重复打印
                    with self._fallback_lock:
                        if not self._use_json_object_fallback:
                            self._use_json_object_fallback = True
                            self._log(
                                f"提示：模型 {self.config.model} 不支持 json_schema 格式，"
                                "已自动降级为 json_object 模式（后续请求同步切换）。"
                            )
                    response = client.chat.completions.create(
                        response_format={"type": "json_object"},
                        **request_kwargs,
                    )
                else:
                    raise  # 网络错误/限速等，交由上层重试
        else:
            response = client.chat.completions.create(
                response_format={"type": "json_object"},
                **request_kwargs,
            )

        raw_content = response.choices[0].message.content or "{}"
        parsed = json.loads(_clean_json_string(raw_content))
        return self._normalize_result(parsed)

    def _analyze_unique_record(self, record: dict[str, Any]) -> dict[str, str]:
        """带退避重试的单条记录分析。

        - 普通错误：最多 max_retries 次，线性退避（n × _RETRY_BACKOFF_BASE 秒）
        - 429 限速：最多 _MAX_RATE_LIMIT_RETRIES 次，指数退避（5→10→20→40→60s）
          429 计数与普通错误计数独立，确保限速重试不占用普通错误重试次数。
        """
        last_error = ""
        generic_attempts = 0
        rate_limit_attempts = 0

        while True:
            if self._is_cancelled():
                return _cancelled_result()
            try:
                return self._request_once(record)
            except Exception as exc:
                last_error = str(exc)
                exc_lower = last_error.lower()
                is_rate_limit = (
                    "429" in last_error
                    or "rate_limit" in exc_lower
                    or "ratelimit" in exc_lower
                    or ("rate" in exc_lower and "limit" in exc_lower)
                )

                if is_rate_limit:
                    rate_limit_attempts += 1
                    if rate_limit_attempts > _MAX_RATE_LIMIT_RETRIES:
                        # 超过限速重试上限，直接返回失败
                        break
                    backoff = min(60.0, 5.0 * (2 ** (rate_limit_attempts - 1)))
                    # 只在首次触发时记录日志，避免淹没进度信息
                    if rate_limit_attempts == 1:
                        self._log(
                            f"  API 限速（429），将等待 {backoff:.0f}s 后重试"
                            f"（最多还可重试 {_MAX_RATE_LIMIT_RETRIES - rate_limit_attempts} 次）…"
                        )
                    time.sleep(backoff)
                else:
                    generic_attempts += 1
                    if generic_attempts >= self.config.max_retries:
                        break
                    backoff = generic_attempts * _RETRY_BACKOFF_BASE
                    time.sleep(backoff)

        return _failed_result(last_error[:200])

    def analyze_records(self, records: list[dict[str, Any]]) -> list[dict[str, str]]:
        if not records:
            return []

        # 去重：相同（模型+关键词+分类+句子）的记录只调用一次 API
        key_to_indices: dict[str, list[int]] = {}
        key_to_record: dict[str, dict[str, Any]] = {}
        ordered_keys: list[str] = []
        for idx, record in enumerate(records):
            key = self._record_key(record, self.config.model)
            if key not in key_to_indices:
                ordered_keys.append(key)
                key_to_indices[key] = []
                key_to_record[key] = record
            key_to_indices[key].append(idx)

        total_unique = len(ordered_keys)
        result_by_key: dict[str, dict[str, str]] = {}
        pending_keys: list[str] = []
        skipped_keys: list[str] = []
        cache_hits = 0

        for key in ordered_keys:
            cached = self._cache.get(key)
            if (isinstance(cached, dict)
                    and cached.get("LLM分析状态") == "成功"
                    and _REQUIRED_CACHE_KEYS.issubset(cached.keys())):
                result_by_key[key] = cached
                cache_hits += 1
                continue
            # max_sentences <= 0 表示无上限
            no_limit = self.config.max_sentences <= 0
            if no_limit or len(pending_keys) < self.config.max_sentences:
                pending_keys.append(key)
            else:
                skipped_keys.append(key)

        if cache_hits:
            self._log(f"LLM 缓存命中 {cache_hits} 条，跳过 API 调用。")

        if skipped_keys:
            self._log(
                f"提示：唯一命中句子共 {total_unique} 条（含缓存 {cache_hits} 条），"
                f"按上限仅新增调用 {self.config.max_sentences} 条，其余 {len(skipped_keys)} 条标记为未分析。"
            )
            for key in skipped_keys:
                result_by_key[key] = _skipped_result()

        limit_desc = "无上限" if self.config.max_sentences <= 0 else str(self.config.max_sentences)
        if pending_keys:
            self._log(
                f"正在调用 LLM 分析句子：新增 {len(pending_keys)} 条（上限：{limit_desc}），"
                f"线程数 {self.config.max_workers}，模型 {self.config.model}。"
            )
            # 修复：进度日志颗粒度改为每 10 条或总量的 10%，以较小者为准
            log_interval = max(1, min(10, len(pending_keys) // 10))
            with ThreadPoolExecutor(max_workers=self.config.max_workers) as executor:
                future_to_key = {
                    executor.submit(self._analyze_unique_record, key_to_record[key]): key
                    for key in pending_keys
                }
                completed = 0
                failed = 0
                successes = 0  # 用于触发增量缓存写盘
                for future in as_completed(future_to_key):
                    key = future_to_key[future]
                    completed += 1
                    result = future.result()
                    result_by_key[key] = result
                    if result.get("LLM分析状态") == "成功":
                        with self._cache_lock:
                            self._cache[key] = result
                        successes += 1
                        # 增量保存：每 _CACHE_SAVE_INTERVAL 条成功结果写一次盘，
                        # 防止长时间运行中途崩溃导致所有进度丢失
                        if successes % _CACHE_SAVE_INTERVAL == 0:
                            self._save_cache()
                    elif result.get("LLM分析状态") == "失败":
                        failed += 1
                    # 修复：更细粒度的进度日志，小批量也能及时反馈
                    if completed % log_interval == 0 or completed == len(pending_keys):
                        self._log(
                            f"LLM 分析进度：{completed}/{len(pending_keys)}"
                            + (f"（失败 {failed} 条）" if failed else "")
                        )

        self._save_cache()

        final_results: list[dict[str, str]] = []
        for record in records:
            key = self._record_key(record, self.config.model)
            final_results.append(result_by_key[key])
        return final_results


# ── 辅助函数：统一构造各种终态结果，避免散落在各处的字典字面量 ──────────────

def _base_result() -> dict[str, str]:
    return {
        "LLM时间指向": "",
        "LLM语态": "",
        "LLM句子类型": "",
        "LLM确定性": "",
        "LLM量化属性": "",
        "LLM语气语调": "",
        "LLM置信度": "",
        "LLM判定依据": "",
        "LLM分析状态": "",
        "LLM分析错误": "",
    }


def _cancelled_result() -> dict[str, str]:
    r = _base_result()
    r["LLM分析状态"] = "已取消"
    r["LLM分析错误"] = "用户已取消"
    return r


def _failed_result(error: str) -> dict[str, str]:
    r = _base_result()
    r["LLM分析状态"] = "失败"
    r["LLM分析错误"] = error
    return r


def _skipped_result() -> dict[str, str]:
    r = _base_result()
    r["LLM分析状态"] = "未分析"
    r["LLM分析错误"] = "超过 LLM 分析上限"
    return r


# ── 公共入口函数 ─────────────────────────────────────────────────────────────

def apply_llm_sentence_analysis(
    df_sentences: pd.DataFrame,
    config: LLMAnalyzerConfig,
    *,
    log_cb: Callable[[str], None] | None = None,
    cancel_event: threading.Event | None = None,
) -> pd.DataFrame:
    """对命中句子 DataFrame 追加 LLM 分析列，返回新 DataFrame（原始不修改）。"""
    if df_sentences.empty:
        return df_sentences

    analyzer = QwenSentenceAnalyzer(config, log_cb=log_cb, cancel_event=cancel_event)

    # 只提取 LLM 实际需要的 3 列转 dict，内存峰值从 ~6GB 降到 ~2.5GB
    _llm_input_cols = [c for c in ["命中关键词", "分类", "命中句子"] if c in df_sentences.columns]
    records = df_sentences[_llm_input_cols].to_dict(orient="records")
    import gc as _gc
    _gc.collect()

    llm_results = analyzer.analyze_records(records)
    del records
    _gc.collect()

    llm_columns = [
        "LLM时间指向",
        "LLM语态",
        "LLM句子类型",
        "LLM确定性",
        "LLM量化属性",
        "LLM语气语调",
        "LLM置信度",
        "LLM判定依据",
        "LLM分析状态",
        "LLM分析错误",
    ]
    llm_df = pd.DataFrame(llm_results, columns=llm_columns)
    del llm_results
    _gc.collect()
    return pd.concat([df_sentences.reset_index(drop=True), llm_df], axis=1)
