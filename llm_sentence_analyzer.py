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
# 1000 句七模型横评结果：qwen3.7-max 与 qwen3.7-plus 的加权 Kappa 并列最高
# （物理 0.77 / 转型 0.74），但前者输出 token 少 19%、且 1000/1000 无缺失，
# 故定为默认。旧默认 qwen-turbo/qwen-plus 系统性偏保守，是 rel=0 过多的部分成因。
DEFAULT_MODEL = "qwen3.7-max"

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

# 整数编码对照表（文档用途，供提示词和后处理参考）
# rel:   不相关=0, 相关=1
# time:  不明显=0, 过去=1, 当前=2, 未来=3, 混合=4
# voice: 不明显=0, 被动=1, 主动=2
# type:  其他=0, 回顾总结=1, 现状描述=2, 未来规划=3, 风险提示=4, 措施行动=5, 结果成效=6
# cert:  低=0, 中=1, 高=2
# quant: 定性=0, 半定量=1, 定量=2
# tone:  消极=0, 其他=1, 积极=2
# conf:  低=0, 中=1, 高=2

RESULT_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "rel":   {"type": "integer", "enum": [0, 1]},
        "time":  {"type": "integer", "enum": [0, 1, 2, 3, 4]},
        "voice": {"type": "integer", "enum": [0, 1, 2]},
        "type":  {"type": "integer", "enum": [0, 1, 2, 3, 4, 5, 6]},
        "cert":  {"type": "integer", "enum": [0, 1, 2]},
        "quant": {"type": "integer", "enum": [0, 1, 2]},
        "tone":  {"type": "integer", "enum": [0, 1, 2]},
        "conf":  {"type": "integer", "enum": [0, 1, 2]},
    },
    "required": ["rel", "time", "voice", "type", "cert", "quant", "tone", "conf"],
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

# 历史默认提示词（v26 全量缓存所对应）的 sha1[:8]。生效提示词哈希等于它时，
# 缓存 key 沿用空后缀，保证既有 14.6 万条标注缓存继续命中；一旦提示词被修改，
# 哈希改变 → key 改变 → 缓存自动失效并重新请求。切勿随提示词改动而更新此常量。
_LEGACY_PROMPT_SHA = "fea337e2"

# _normalize_result：LLM 返回的 JSON 中至少要包含其中一个字段才视为有效响应
_EXPECTED_RESULT_KEYS = frozenset({"rel", "time", "voice", "type", "cert", "quant", "tone", "conf"})

# 缓存命中校验：所有字段都存在时才使用缓存，防止格式不完整的旧缓存污染结果
_REQUIRED_CACHE_KEYS = frozenset({
    "LLM相关性", "LLM时间指向", "LLM语态", "LLM句子类型",
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


class _SQLiteCacheProxy:
    """SQLite 后端的 dict-like 缓存代理。

    不把数据全量加载到内存，彻底解决大缓存导致的 OOM 问题。
    接口与 dict 兼容（__contains__ / __getitem__ / __setitem__ / __len__）。
    写入时每 50 条批量 commit 一次，减少 fsync 频率。
    """

    _COMMIT_BATCH = 50

    def __init__(self, conn, lock: threading.Lock) -> None:
        self._conn = conn
        self._lock = lock
        self._pending = 0

    def get(self, key: str, default: Any = None) -> Any:
        with self._lock:
            row = self._conn.execute(
                "SELECT value FROM llm_cache WHERE key = ?", (key,)
            ).fetchone()
        return json.loads(row[0]) if row else default

    def __contains__(self, key: object) -> bool:
        with self._lock:
            row = self._conn.execute(
                "SELECT 1 FROM llm_cache WHERE key = ?", (key,)
            ).fetchone()
        return row is not None

    def __getitem__(self, key: str) -> dict[str, Any]:
        with self._lock:
            row = self._conn.execute(
                "SELECT value FROM llm_cache WHERE key = ?", (key,)
            ).fetchone()
        if row is None:
            raise KeyError(key)
        return json.loads(row[0])

    def __setitem__(self, key: str, value: dict[str, Any]) -> None:
        encoded = json.dumps(value, ensure_ascii=False)
        with self._lock:
            self._conn.execute(
                "INSERT OR REPLACE INTO llm_cache (key, value) VALUES (?, ?)",
                (key, encoded),
            )
            self._pending += 1
            if self._pending >= self._COMMIT_BATCH:
                self._conn.commit()
                self._pending = 0

    def __len__(self) -> int:
        with self._lock:
            return self._conn.execute(
                "SELECT COUNT(*) FROM llm_cache"
            ).fetchone()[0]

    def flush(self) -> None:
        with self._lock:
            if self._pending > 0:
                self._conn.commit()
                self._pending = 0


class QwenSentenceAnalyzer:
    SYSTEM_PROMPT = """你是中文财经文本分析助手，分析年报中命中生物多样性/环境关键词的句子。
只返回 JSON 整数，禁止输出任何额外文字。

字段与编码：
rel  相关性：不相关=0，相关=1
  相关=句子核心确实讨论关键词分类主题（生物多样性、自然生态系统、物种保护等）
  不相关=仅字面含关键词但实际讨论其他主题

═══ 通用排除规则（rel=0）═══
- 关键词出现在商业生态语境（生态圈/生态链/产业生态/云生态/生态体系等），句子核心讨论的是商业合作网络/IT平台/数字化转型而非自然生态系统
- "生态系统"出现在科技/商业语境，句子核心为技术架构或商业生态而非自然生态系统。商业/科技生态系统包括但不限于：Hadoop生态系统/物联网生态系统/支付生态系统/游戏生态系统/用户生态系统/产业生态系统/数字生态系统/大屏生态系统/工业生态系统/移动互联网生态系统/营销生态系统/商业生态系统/制造生态系统/渠道生态系统/平台生态系统/金融生态系统/软件生态系统/智慧城市生态系统/教育生态系统/医疗生态系统/汽车生态系统/零售生态系统/内容生态系统/创新生态系统（指商业创新体系）；判断标准：修饰词为非自然实体（商业/工业/数字/科技/用户/营销/渠道/产业/平台），则标注 rel=0
- "生态系统"仅在以下自然语境中标注 rel=1：句子核心讨论碳汇/碳储量/生物多样性/物种多样性/栖息地/自然保护/生态修复/植被恢复/水文/土壤健康/食物链/生态退化，或明确涉及森林/海洋/湿地/河流/草地/高山等自然地理单元的生态系统功能
- 关键词出现在句子中仅作为公司名称一部分，且句子内容为财务数据/营业收入/利润等，无具体生态活动描述

═══ 关键词专项规则（v16新增）═══

【物种】rel=1情境（农业生物多样性/agrobiodiversity属于生物多样性范畴）：
  相关=1：①珍稀/濒危/保护物种；②物种多样性/遗传多样性/种质资源保护；③农作物/畜禽地方品种/野生近缘种保育（遗传多样性保存目的）；④入侵物种/外来物种；⑤物种栖息地保护；⑥野生植物/动物物种调查监测
  不相关=0：①纯粹的品种选育/育种目标（增产/抗病）而无多样性保育语境；②"新品种/优良品种上市"等商业种子描述；③药品名称含"物种"字眼

【淡水】rel=1情境（淡水生态系统依存关系属于生物多样性暴露）：
  相关=1：①淡水生态系统保护/修复；②淡水鱼类/水生生物保护；③淡水渔业对河流/湖泊生态的依存关系披露；④淡水生物多样性
  不相关=0：①纯粹的淡水养殖产量/营收/市场描述（无生态系统语境）；②淡水处理设备/净水技术

【碳汇】rel=1情境（生态系统碳汇功能属于生物多样性服务范畴）：
  相关=1：①森林/草地/湿地/海洋碳汇作为生态系统服务功能；②碳汇与生态修复/植被恢复结合语境；③自然碳汇保护/增汇
  不相关=0：①碳汇交易/碳汇项目CCER/碳市场（纯碳金融工具语境，无生态系统保护内容）；②"碳汇林"仅作为碳指标资产而无生态保护描述

【野生动物】rel=1情境：
  相关=1：①野生动物保护/救助/栖息地；②野生动物种群监测（保护目的）；③野生动物贸易监管/禁令；④野生动物与企业生产活动冲突风险
  不相关=0：①野生动物园/野生动物主题乐园的门票/旅游收入/景区介绍（纯商业旅游语境）；②"养殖/驯化野生动物"商业繁育无保护意义；③机场野生动物管理系统/机场驱鸟系统/鸟击防治平台（航空安全商业产品，目的是防鸟击而非生物多样性保护；句中常见"机场""鸟击""航空安全"等词）

【珊瑚】rel=1情境：
  相关=1：①珊瑚礁生态系统/珊瑚白化/珊瑚保护；②珊瑚礁周边作业的海洋企业生态影响
  不相关=0：①草珊瑚（中药材/OTC药品成分）；②以"珊瑚"命名的商品/颜色描述

【海洋】rel=1情境（海洋生物多样性保护语境）：
  相关=1：①海洋生物多样性/海洋生态系统保护/修复；②海洋物种/珊瑚礁/红树林/海草床；③企业海洋作业对海洋生态的影响评估；④海洋污染对生物多样性的影响
  不相关=0：①海洋工程/海洋运输/海洋能源/海洋旅游（无生态系统保护语境）；②"海洋强国战略"宏观政策描述；③海洋食品/水产加工（纯商业语境）

【动物/植物】rel=1情境：
  相关=1：①野生动物/植物保护/调查；②濒危/珍稀动物/植物；③动物/植物多样性；④生态修复中的动植物恢复；⑤动物/植物栖息地
  不相关=0：①食品加工中的动物原料/植物原料；②宠物行业/养殖业（纯商业）；③"植物工厂/植物提取物"商业语境；④药材中的动植物原料（无保育意义）

【树木/树林】rel=1情境：
  相关=1：①造林/植树/护林（生态修复目的）；②林业生物多样性；③树木砍伐/树林破坏的生态影响
  不相关=0：①木材加工/家具生产（纯林木商品）；②"种树绿化"仅作景观工程无生态意义描述

【热带】rel=1情境：
  相关=1：①热带雨林/热带生态系统；②热带物种/热带生物多样性；③热带森林砍伐
  不相关=0：①热带旅游/热带水果/热带气候旅游描述；②"热带鱼养殖"纯商业水产

【雨林】rel=1情境：
  相关=1：①热带雨林生态系统保护/毁林；②雨林生物多样性；③企业供应链对雨林的影响
  不相关=0：①房地产项目/楼盘以"雨林"命名（如"雨林澜山"）；②雨林主题景区旅游

【生态环境】注意：此词极度高频，须严格过滤：
  相关=1：①企业活动对自然生态环境（森林/湿地/栖息地/物种）的具体影响描述；②生态环境损害/赔偿（有具体生物多样性内容）；③自然生态环境保护/修复行动
  不相关=0：①泛化的"改善生态环境"政策宣言（无具体生物多样性行动）；②"生态环境部/监管部门"仅作机构名称；③"绿色发展/生态环境优化"作为企业整体ESG描述无具体内容

【森林】注意：此词常出现在品牌名/楼盘名/旅游景区中，须排除：
  相关=1：①天然林/原始森林保护/保育；②造林/植树/护林（生态修复/生物多样性目的，非商业绿化）；③森林生物多样性/物种调查；④毁林/森林砍伐对生态的影响；⑤森林碳汇作为生态系统服务；⑥森林退化/恢复语境
  不相关=0：①品牌/楼盘/产品名称含"森林"（如"元气森林""三湘森林海尚城""森林公馆""雨林/海上森林"等）；②森林公园/自然风景区旅游收入/门票（纯旅游商业）；③木材/林产品加工/销售（纯商业林业）；④"林业局/林业部"等机构名称（仅作监管机构提及）；⑤"森林城市/绿色城市"纯政策口号无具体生态行动

【生物安全】注意：此词在年报中常出现于养殖业疫病防控语境，与生物多样性无关：
  相关=1：①外来入侵物种的生物安全威胁（生物入侵/生态安全）；②生物多样性安全/基因资源安全保护；③转基因生物/GMO的生态安全评估（对自然生态系统影响）；④实验室/研究机构高风险病原体的生物安全防护（BSL实验室安全，非养殖）
  不相关=0：①畜禽/猪/鸡/水产养殖中的"生物安全防控"（疫病/疫情控制，如非洲猪瘟防控/禽流感防控）；②药物/化合物"生物安全性"（毒理/药理测试）；③食品/农产品"生物安全"（食品安全/污染物检测）；④"生物安全"仅在公司经营计划中作为管理指标无生态内容；⑤"全面生物安全防控能力"用于描述养殖场管理体系

【绿色矿山】注意：此词常出现在合规认证语境，无实质生态内容：
  相关=1：①矿山绿色化改造的具体生态措施（植被恢复/废水处理/生物多样性影响评估）；②矿山生态修复的实质内容（物种保护/栖息地恢复）
  不相关=0：①"国家级/省级绿色矿山"认定/申报/数量统计（仅作资质/荣誉列举）；②"推进绿色矿山建设"政策宣言（无具体生态内容）；③"绿色矿山建设标准/规范"泛政策描述；④含"绿色矿山"的营收/成本/采矿量等纯财务数据

【生物修复】注意：中文"生物修复"存在口腔医学同名词，须严格区分：
  相关=1：①土壤/水体/湿地/矿山生物修复（微生物/植物修复污染土壤/水体）；②生态系统生物修复（自然生境恢复中的生物手段）
  不相关=0：①生物修复膜（口腔软组织修复医疗器械/牙科材料）；②口腔/外科/骨科中的"生物性修复材料"；③医疗器械注册文件中的"生物修复"

【沼泽】注意：此词在非生态语境中常指"沼泽地形"，须排除：
  相关=1：①沼泽湿地生态系统保护/修复；②沼泽动植物/生物多样性；③企业活动对沼泽湿地的影响
  不相关=0：①游戏/影视/产品以"沼泽"命名（如"沼泽激战/Swamp Attack"）；②工程机械/车辆"适应沼泽地形"的描述；③探险/越野活动中"穿越沼泽"的环境描述

【生态学】注意：此词在年报中常出现于非宏观自然生态语境，须排除：
  相关=1：①宏观生态学研究（种群/群落/生态系统生态学）；②生态学视角的企业自然影响分析；③保护生态学相关表述
  不相关=0：①微生态学（肠道菌群/微生物组/益生菌）语境；②以"生态学"命名的公司/地产项目；③人类生态学/工业生态学在纺织/制造业中的认证描述

【遗传资源】注意：此词在年报中常出现于医药/农业商业语境，须区分：
  相关=1：①野生动植物遗传资源保护（就地/迁地保护）；②农业遗传多样性/种质资源保护；③生物多样性遗传资源获取与惠益分享（名古屋议定书相关）；④自然生态系统遗传资源
  不相关=0：①"中国人类遗传资源管理办公室"审批（医药临床试验监管）；②医疗/制药公司的"人类遗传资源"（DNA样本/血液样本/基因数据）；③人类遗传资源出境申请/许可（监管合规语境，无生物多样性内容）

【生物资源】注意：此词在年报中常出现于医疗/食品/工业语境，须区分：
  相关=1：①野生动植物生物资源保护/可持续利用；②生物多样性自然资本/生态系统服务语境中的生物资源；③濒危/稀缺生物资源保育
  不相关=0：①细胞存储/干细胞/脐带血存储（医疗生物样本库）；②微生物菌种库（发酵食品/医药工业用菌种，无保育意义）；③工业酶制剂/生物技术原料的"生物资源"

═══ 其他字段编码 ═══
neg  否定语态：否定=1，肯定=0
  否定=1：句子核心含"不存在/不涉及/不影响/未发现/无重大/没有发生"等否定词，表达的是"不存在该风险/影响/依赖"
    典型否定句：①"公司不涉及自然保护区"；②"我们不存在重大生物多样性风险"；③"本报告期内未发生生态环境损害事件"
  否定=0（肯定句）：正面陈述该主题，即使内容是负面的（如"公司生产导致栖息地破坏"仍为肯定句，rel=1,neg=0）
  注意：否定句在词频统计时仍被计入（公司提及了该话题），但研究者可用neg字段进行净值调整

time 时间指向：不明显=0，过去=1，当前=2，未来=3，混合=4
voice 语态：不明显=0，被动=1，主动=2
type 句子类型：其他=0，回顾总结=1，现状描述=2，未来规划=3，风险提示=4，措施行动=5，结果成效=6
cert 确定性：低=0，中=1，高=2
quant 量化：定性=0，半定量=1，定量=2
  定量=含明确数字/比例/金额；半定量=含"大幅""显著"等程度词；定性=纯描述
tone 语气：消极=0，其他=1，积极=2
conf 置信度：低=0，中=1，高=2"""

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
        # 缓存 key 必须绑定"实际生效的提示词内容"。
        # 旧实现只在使用自定义提示词时才混入哈希，于是直接修改内置 SYSTEM_PROMPT
        # 常量后重跑，key 完全不变 —— 缓存会原样吐回旧标注，改词无效且难以察觉。
        # 这里改为始终按生效提示词取哈希；_LEGACY_PROMPT_SHA 是历史默认提示词的
        # 哈希，命中它时沿用空后缀，使既有缓存继续可用。
        digest = hashlib.sha1(self._system_prompt.encode("utf-8")).hexdigest()[:8]
        self._prompt_hash_suffix: str = "" if digest == _LEGACY_PROMPT_SHA else f":{digest}"

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

    # ------------------------------------------------------------------
    # 缓存后端：JSON dict（旧） vs SQLite（新，内存占用低 10×+）
    # ------------------------------------------------------------------

    def _load_cache(self, cache_path: str | None):
        """返回缓存对象。

        优先使用 SQLite（.db）；若只有 JSON（.json），自动迁移后继续。
        SQLite 缓存不把全量数据加载到内存，彻底解决大缓存 OOM 问题。
        """
        if not cache_path:
            return {}

        path = Path(cache_path)

        # ── 判断使用哪种后端 ──────────────────────────────────────────
        if path.suffix.lower() == ".db":
            # 已明确指定 SQLite
            db_path = path
            json_path = path.with_suffix(".json")
        else:
            # 默认：优先升级到同目录 .db 文件，保留 .json 作备份
            json_path = path
            db_path = path.with_suffix(".db")

        # ── 如果 SQLite 文件不存在但 JSON 存在，先迁移 ───────────────
        if not db_path.exists() and json_path.exists():
            self._log("检测到旧版 JSON 缓存，正在迁移至 SQLite（仅首次，无需重复）…")
            try:
                import sqlite3 as _sq3
                db_path.parent.mkdir(parents=True, exist_ok=True)
                conn = _sq3.connect(str(db_path))
                conn.execute("PRAGMA journal_mode=WAL")
                conn.execute(
                    "CREATE TABLE IF NOT EXISTS llm_cache "
                    "(key TEXT PRIMARY KEY, value TEXT NOT NULL)"
                )
                conn.commit()
                with json_path.open("r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    batch: list[tuple[str, str]] = []
                    for k, v in data.items():
                        batch.append((k, json.dumps(v, ensure_ascii=False)))
                        if len(batch) >= 2000:
                            conn.executemany(
                                "INSERT OR IGNORE INTO llm_cache (key, value) VALUES (?, ?)",
                                batch,
                            )
                            conn.commit()
                            batch.clear()
                    if batch:
                        conn.executemany(
                            "INSERT OR IGNORE INTO llm_cache (key, value) VALUES (?, ?)",
                            batch,
                        )
                        conn.commit()
                    n = conn.execute("SELECT COUNT(*) FROM llm_cache").fetchone()[0]
                    self._log(f"✅ SQLite 缓存迁移完成：{n} 条 → {db_path}")
                conn.close()
            except Exception as exc:
                self._log(f"⚠️  JSON→SQLite 迁移失败（{exc}），将继续使用 JSON 缓存。")
                db_path = None  # type: ignore[assignment]

        # ── 打开 SQLite ───────────────────────────────────────────────
        if db_path is not None and db_path.exists():
            try:
                import sqlite3 as _sq3
                conn = _sq3.connect(str(db_path), check_same_thread=False)
                conn.execute("PRAGMA journal_mode=WAL")
                conn.execute("PRAGMA synchronous=NORMAL")
                conn.execute("PRAGMA cache_size=8000")   # 32 MB 页缓存
                conn.execute(
                    "CREATE TABLE IF NOT EXISTS llm_cache "
                    "(key TEXT PRIMARY KEY, value TEXT NOT NULL)"
                )
                conn.commit()
                n = conn.execute("SELECT COUNT(*) FROM llm_cache").fetchone()[0]
                self._log(f"SQLite 缓存已加载：{n} 条（{db_path.name}），内存占用极低。")
                # 存到实例上，供 _save_cache / 查询使用
                self._sqlite_conn = conn
                self._sqlite_db_path = db_path
                return _SQLiteCacheProxy(conn, self._cache_lock)
            except Exception as exc:
                self._log(f"⚠️  SQLite 缓存打开失败（{exc}），降级到 JSON 模式。")

        # ── 降级：JSON dict（旧行为）────────────────────────────────
        if json_path.exists():
            try:
                with json_path.open("r", encoding="utf-8") as f:
                    data = json.load(f)
                return data if isinstance(data, dict) else {}
            except Exception as exc:
                self._log(f"警告：LLM 缓存文件读取失败（{exc}），将以空缓存启动。")
        return {}

    def _save_cache(self) -> None:
        """持久化缓存。SQLite 已实时写盘，仅 JSON 模式需要此方法。"""
        # SQLite 模式：WAL 已保证实时持久化，flush 一次即可
        if isinstance(self._cache, _SQLiteCacheProxy):
            try:
                self._cache.flush()
            except Exception:
                pass
            return

        # JSON 模式（向下兼容）
        if not self.config.cache_path:
            return
        path = Path(self.config.cache_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(path.suffix + ".tmp")
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

    @staticmethod
    def _safe_code(value: Any, allowed: set[int], default: int) -> int:
        """将 LLM 返回值转为合法整数编码，超出范围则用默认值。"""
        try:
            v = int(value)
            return v if v in allowed else default
        except (TypeError, ValueError):
            return default

    def _normalize_result(self, data: dict[str, Any]) -> dict[str, Any]:
        # 空响应或缺失所有维度字段时返回失败，避免将全默认值误写为"成功"
        if not isinstance(data, dict) or not any(k in data for k in _EXPECTED_RESULT_KEYS):
            return _failed_result("LLM返回了空响应或格式不符合预期（缺少所有维度字段）")
        sc = self._safe_code
        return {
            "LLM相关性":   sc(data.get("rel"),   {0, 1},             1),
            "LLM时间指向": sc(data.get("time"),  {0, 1, 2, 3, 4},    0),
            "LLM语态":     sc(data.get("voice"), {0, 1, 2},           0),
            "LLM句子类型": sc(data.get("type"),  {0, 1, 2, 3, 4, 5, 6}, 0),
            "LLM确定性":   sc(data.get("cert"),  {0, 1, 2},           1),
            "LLM量化属性": sc(data.get("quant"), {0, 1, 2},           0),
            "LLM语气语调": sc(data.get("tone"),  {0, 1, 2},           1),
            "LLM置信度":   sc(data.get("conf"),  {0, 1, 2},           1),
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
                # 系统提示词在所有请求间完全相同，且体量远大于用户内容，
                # 打上 cache_control 走显式上下文缓存：命中部分按 10% 计费
                # （首次创建 125%，每次命中续期 5 分钟）。
                # 实测 qwen3.7-max：不加此标记仅隐式命中约 62%，加后达 99%。
                {"role": "system", "content": [
                    {"type": "text", "text": self._system_prompt,
                     "cache_control": {"type": "ephemeral"}},
                ]},
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

def _base_result() -> dict[str, Any]:
    return {
        "LLM相关性":   -1,
        "LLM时间指向": -1,
        "LLM语态":     -1,
        "LLM句子类型": -1,
        "LLM确定性":   -1,
        "LLM量化属性": -1,
        "LLM语气语调": -1,
        "LLM置信度":   -1,
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
        "LLM相关性",
        "LLM时间指向",
        "LLM语态",
        "LLM句子类型",
        "LLM确定性",
        "LLM量化属性",
        "LLM语气语调",
        "LLM置信度",
        "LLM分析状态",
        "LLM分析错误",
    ]
    llm_df = pd.DataFrame(llm_results, columns=llm_columns)
    del llm_results
    _gc.collect()
    return pd.concat([df_sentences.reset_index(drop=True), llm_df], axis=1)
