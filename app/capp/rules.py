"""CAPP 控制计划 → 质量策划转换规则。

规则来源：项目根目录《应用程序规则说明.md》。
"""

from __future__ import annotations

import re

KZJH_COLUMNS: tuple[str, ...] = (
    "工序号",
    "特性类型",
    "特征描述（产品特性+范围公差）",
    "特殊特性",
    "判定类型",
    "实验室检测",
    "自动采集",
    "目标值",
    "上限",
    "下限",
    "单位",
    "检验方式",
    "频次1（时间/）",
    "频次2（数量/件）",
    "抽样数",
    "评价测量技术",
    "责任人",
    "数据保存方法",
    "检查单位",
    "检验顺序",
    "对应物料",
    "备注",
)

TARGET_FIELDS: tuple[str, ...] = (
    "过程编号（工序号）",
    "过程名称",
    "机器、装置夹具、工装",
    "特性编号",
    "产品特性",
    "过程特性",
    "特殊特性分类（特殊特性）",
    "规范/公差",
    "评价测量技术（评价测量技术）",
    "样本数量",
    "样本频率",
    "控制方法",
    "反应计划",
)

# 控制计划输出列与目标字段一一对应
CONTROL_PLAN_COLUMNS: tuple[str, ...] = TARGET_FIELDS

SC_KEYS = ("SC", "Z", "YS")
CC_KEYS = ("CC", "G", "YC")

UNIT_COEF = {
    "分钟": 1 / 60,
    "min": 1 / 60,
    "秒": 1 / 3600,
    "s": 1 / 3600,
    "小时": 1,
    "h": 1,
    "天": 24,
    "日": 24,
    "月": 720,
    "年": 8760,
}

# 允许的单位白名单，避免括号、顿号、反斜杠等被误当成单位前后缀
_UNIT_TOKENS: tuple[str, ...] = (
    "N/0.1m/s", "N/0.3m/s", "N/0.5m/s", "N/0.6m/s", "N.m", "N·m", "N*m", "N/mm",
    "kgf", "kN", "KN", "N", "mm/s", "m/s", "mm", "cm", "μm", "um", "nm", "m",
    "mg", "kg", "g", "Lux", "lx", "Hz", "r/min", "rpm",
    "MPa", "kPa", "Pa", "bar", "mA", "A", "V", "W", "Ω", "dB", "J",
    "min", "s", "h", "ml", "mL", "L",
    "%", "°", "℃", "件", "次", "个", "套",
)
# 字典序会先命中短单位（如 min 被 m 截断），必须按长度倒序才能匹配完整单位
_UNIT_TOKENS_SORTED: tuple[str, ...] = tuple(
    sorted(_UNIT_TOKENS, key=len, reverse=True)
)
# 括号、顿号、反斜杠等只作分隔符，不能拼进单位
_UNIT_SPLIT_RE = re.compile(r"[（）()\[\]【】《》〈〉、，,；;：:。\\|／｜\"'“”‘’\s]+")


def _sanitize_unit(raw: str) -> str:
    """从尾缀文本中提取规范单位，括号、顿号、反斜杠等只作分隔。"""
    for segment in _UNIT_SPLIT_RE.split(clean(raw)):
        if not segment:
            continue
        lowered = segment.lower()
        for token in _UNIT_TOKENS_SORTED:
            if lowered.startswith(token.lower()):
                return token
    return ""

# 控制计划中常见的班次/容器简写（如“班”“框”），统一还原成模板写法
_FREQ_ALIASES = {
    "班": "每班",
    "框": "每框",
    "卷": "每卷",
}

# 无数量前缀的时间频率（如“每月”）按 1 次换算
_BARE_TIME_FREQ = {
    "分钟": 1 / 60,
    "秒": 1 / 3600,
    "小时": 1,
    "天": 24,
    "日": 24,
    "月": 720,
    "年": 8760,
    "每分钟": 1 / 60,
    "每秒": 1 / 3600,
    "每小时": 1,
    "每天": 24,
    "每日": 24,
    "每月": 720,
    "每年": 8760,
}

_BLOCK_WORDS = (
    "清洁度",
    "外观",
    "尺寸",
    "性能",
    "功能",
    "装配",
    "材料",
    "工艺",
    "检验",
    "扭矩",
    "尺寸链",
)
_BLOCK_RE = re.compile(r"(?=(?:" + "|".join(_BLOCK_WORDS) + r")[：:])")
_DESC_WORDS = ("不得", "标准", "要求", "允许", "说明", "下述", "颗粒", "污染物")
_RANGE_SYMBOL_RE = re.compile(r"[±≥≤>＞<＜]|[-—~～]\s*\d")


def clean(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\r", "\n").strip()


def _fmt(number: float | int | None) -> str:
    if number is None:
        return ""
    text = f"{float(number):.4f}".rstrip("0").rstrip(".")
    return text if text not in ("", "-0") else "0"


def compute_工序号_补位(value: str) -> str:
    text = clean(value)
    return f"{text}0" if text else ""


def compute_特性类型_new(product_feature: str, process_feature: str) -> str:
    return "产品特性" if clean(product_feature) else "过程特性"


def compute_特征描述_new(
    product_feature: str, process_feature: str, spec: str
) -> str:
    header = clean(product_feature) or clean(process_feature)
    body = _dedup_feature_desc(clean(spec))
    if not header:
        return f"：{body}" if body else ""
    return f"{header}：{body}" if body else header


def compute_特殊特性(value: str) -> str:
    text = clean(value).upper()
    if not text:
        return "/"
    if any(key in text for key in SC_KEYS):
        return "SC"
    if any(key in text for key in CC_KEYS):
        return "CC"
    return "/"


def _has_pseudo_range_remainder(text: str, end: int) -> bool:
    rest = text[end:].lstrip()
    if not rest:
        return False
    if rest[0].isdigit():
        return True
    return bool(re.match(r"[A-Za-z]\d", rest))


def _valid_range(
    match: re.Match, text: str, first: str, second: str
) -> tuple[float, float] | None:
    if _has_pseudo_range_remainder(text, match.end()):
        return None
    a, b = float(first), float(second)
    return (a, b) if a <= b else (b, a)


def compute_判定类型(value: str, process_feature: str = "") -> str:
    text = clean(value)
    if not text:
        return "合格/不合格"
    if "机种" in text:
        return "合格/不合格"
    # 基准值/公差：1.6/0.2（公差说明放在过程特性或本字段内）
    if "公差" in text or "公差" in clean(process_feature):
        if re.search(r"\d+(?:\.\d+)?\s*/\s*\d+(?:\.\d+)?", text):
            return "值判断"
    if re.search(r"[±≥≤>＞<＜]\s*\d", text):
        return "值判断"
    for match in re.finditer(
        r"\d+(?:\.\d+)?\s*(?:[A-Za-zμ℃°%]+)?\s*[-—~～]\s*\d+(?:\.\d+)?", text
    ):
        numbers = re.findall(r"\d+(?:\.\d+)?", match.group())
        pair = _valid_range(match, text, numbers[0], numbers[-1])
        if pair is not None:
            return "值判断"
    if ">=" in text or "<=" in text:
        return "值判断"
    return "合格/不合格"


def compute_实验室检测(value: str) -> str:
    text = clean(value)
    return "Y" if ("试验" in text or "MTS" in text.upper()) else "N"


def compute_自动采集(judge_type: str, measure: str) -> str:
    return "Y" if judge_type == "值判断" and "控制面板" in clean(measure) else "N"


def parse_single_range(
    value: str, process_feature: str = ""
) -> tuple[str, str, str, str]:
    """解析单段规范/公差 → (目标值, 上限, 下限, 单位)。"""
    text = clean(value)
    if not text:
        return "", "", "", ""

    # 1. 基准值/公差：1.6/0.2（公差说明可能在过程特性或本字段内）
    if "/公差" in clean(process_feature) or "公差" in text:
        match = re.search(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", text)
        if match:
            nominal, tolerance = float(match.group(1)), float(match.group(2))
            return (
                _fmt(nominal),
                _fmt(nominal + tolerance),
                _fmt(nominal - tolerance),
                "",
            )

    # 2.5 带符号公差：16%±2%、5.7°±2°
    match = re.search(
        r"(\d+(?:\.\d+)?)\s*([%％°])\s*±\s*(\d+(?:\.\d+)?)\s*([%％°])?", text
    )
    if match:
        nominal, tolerance = float(match.group(1)), float(match.group(3))
        unit = "%" if "%" in (match.group(2) + (match.group(4) or "")) else "°"
        return (
            _fmt(nominal),
            _fmt(nominal + tolerance),
            _fmt(nominal - tolerance),
            unit,
        )

    # 3. 括号容差：22.12（0,-0.02）
    match = re.search(
        r"(\d+(?:\.\d+)?)\s*[（(]\s*([+-]?\d+(?:\.\d+)?)\s*[,，]\s*([+-]?\d+(?:\.\d+)?)\s*[）)]"
        r"\s*([A-Za-zμ℃°%]*)",
        text,
    )
    if match:
        nominal = float(match.group(1))
        upper = nominal + float(match.group(2))
        lower = nominal + float(match.group(3))
        return _fmt(nominal), _fmt(upper), _fmt(lower), _sanitize_unit(match.group(4))

    # 3. ± + 括号单位
    match = re.search(
        r"(\d+(?:\.\d+)?)\s*±\s*(\d+(?:\.\d+)?)\s*[（(]([^）)]*)[）)]", text
    )
    if match:
        nominal, tolerance = float(match.group(1)), float(match.group(2))
        return (
            _fmt(nominal),
            _fmt(nominal + tolerance),
            _fmt(nominal - tolerance),
            _sanitize_unit(match.group(3)),
        )

    # 4. A-B + 括号单位
    match = re.search(
        r"(\d+(?:\.\d+)?)\s*[-—~～]\s*(\d+(?:\.\d+)?)\s*[（(]([^）)]*)[）)]", text
    )
    if match:
        pair = _valid_range(match, text, match.group(1), match.group(2))
        if pair:
            return (
                _fmt((pair[0] + pair[1]) / 2),
                _fmt(pair[1]),
                _fmt(pair[0]),
                _sanitize_unit(match.group(3)),
            )

    # 5. 比较符 + 括号单位
    match = re.search(
        r"([≥≤>＞<＜])\s*(\d+(?:\.\d+)?)\s*[（(]([^）)]*)[）)]", text
    )
    if match:
        value_num = _fmt(match.group(2))
        lower = value_num if match.group(1) in "≥＞" else ""
        upper = value_num if match.group(1) in "≤＜" else ""
        return "", upper, lower, _sanitize_unit(match.group(3))

    # 6. ± 常规
    match = re.search(
        r"(\d+(?:\.\d+)?)\s*±\s*(\d+(?:\.\d+)?)\s*"
        r"([A-Za-zμ℃°%][A-Za-zμ℃°%0-9./·*]*)?",
        text,
    )
    if match:
        nominal, tolerance = float(match.group(1)), float(match.group(2))
        return (
            _fmt(nominal),
            _fmt(nominal + tolerance),
            _fmt(nominal - tolerance),
            _sanitize_unit(match.group(3)),
        )

    # 7. A-B 常规（含伪范围保护）
    match = re.search(
        r"(\d+(?:\.\d+)?)\s*(?:[A-Za-zμ℃°%]+)?\s*[-—~～]\s*(\d+(?:\.\d+)?)\s*"
        r"([A-Za-zμ℃°%][A-Za-zμ℃°%./·*]*)?",
        text,
    )
    if match:
        pair = _valid_range(match, text, match.group(1), match.group(2))
        if pair:
            return (
                _fmt((pair[0] + pair[1]) / 2),
                _fmt(pair[1]),
                _fmt(pair[0]),
                _sanitize_unit(match.group(3)),
            )

    # 8. 比较符
    match = re.search(
        r"([≥≤>＞<＜])\s*(\d+(?:\.\d+)?)\s*([A-Za-zμ℃°%][A-Za-zμ℃°%./·*]*)?",
        text,
    )
    if match:
        value_num = _fmt(match.group(2))
        lower = value_num if match.group(1) in "≥＞" else ""
        upper = value_num if match.group(1) in "≤＜" else ""
        return "", upper, lower, _sanitize_unit(match.group(3))

    return "", "", "", ""


def _is_explanatory(segment: str) -> bool:
    text = segment.strip()
    if not text:
        return False
    if any(word in text for word in _DESC_WORDS):
        return True
    return len(text) >= 14 and not _RANGE_SYMBOL_RE.search(text) and "±" not in text


def _count_parseable(segments: list[str], process_feature: str) -> int:
    count = 0
    for segment in segments:
        target, upper, lower, unit = parse_single_range(segment, process_feature)
        if target or upper or lower or unit:
            count += 1
        elif re.fullmatch(r"\s*\d+\s*[件次个只套]\s*", segment):
            count += 1
    return count


def _try_split(segment: str, separators: str, process_feature: str) -> list[str]:
    parts = [part.strip() for part in re.split(separators, segment)]
    parts = [part for part in parts if part]
    if len(parts) < 2:
        return [segment]
    if any(_is_explanatory(part) for part in parts):
        return [segment]
    if _count_parseable(parts, process_feature) < 2:
        return [segment]
    return parts


def split_multi_ranges(value: str, process_feature: str = "") -> list[str]:
    text = clean(value)
    if not text:
        return []

    # 1. 换行
    segments = [seg.strip() for seg in text.split("\n") if seg.strip()]
    # 2. block 头 / Px: 前缀
    expanded: list[str] = []
    for segment in segments:
        pieces = [piece.strip() for piece in _BLOCK_RE.split(segment) if piece.strip()]
        for piece in pieces:
            if re.search(r"[，,、;；]\s*[A-Z]{1,3}[：:]", piece):
                expanded.extend(
                    part.strip()
                    for part in re.split(r"[，,、;；]\s*(?=[A-Z]{1,3}[：:])", piece)
                    if part.strip()
                )
            else:
                expanded.append(piece)
    segments = expanded or segments

    # 3~6. 分隔符试拆
    result: list[str] = []
    for segment in segments:
        normalized = segment.replace("、", ",").replace("，", ",")
        split_by_comma = _try_split(normalized, r",", process_feature)
        if len(split_by_comma) > 1:
            result.extend(part.strip() for part in split_by_comma if part.strip())
            continue
        if ";" in normalized and len(_RANGE_SYMBOL_RE.findall(segment)) >= 2:
            split_by_semicolon = _try_split(normalized, r";", process_feature)
            if len(split_by_semicolon) > 1:
                result.extend(
                    part.strip() for part in split_by_semicolon if part.strip()
                )
                continue
        result.append(segment)

    return result or [text]


def build_range_cell_values(
    value: str, process_feature: str = ""
) -> list[tuple[str, str, str, str, str]]:
    """返回 [(段文本, 目标值, 上限, 下限, 单位)]。"""
    text = clean(value)
    if not text:
        return [("", "", "", "", "")]
    segments = split_multi_ranges(text, process_feature)
    parsed: list[tuple[str, str, str, str, str]] = []
    for segment in segments:
        target, upper, lower, unit = parse_single_range(segment, process_feature)
        if target or upper or lower or unit:
            parsed.append((segment, target, upper, lower, unit))
    if "机种" in text:
        return [(text, "", "", "", "")]
    if not parsed:
        return [(text, "", "", "", "")]
    return parsed


def _dedup_feature_desc(value: str) -> str:
    text = clean(value)
    if not text:
        return ""
    blocks = [block.strip() for block in _BLOCK_RE.split(text) if block.strip()]
    if len(blocks) <= 1:
        return text
    seen: set[str] = set()
    result: list[str] = []
    for block in blocks:
        key = re.sub(r"\s+", "", block)
        if key in seen:
            continue
        seen.add(key)
        result.append(block)
    return "\n".join(result)


def split_long_feature_desc(value: str, threshold: int = 200) -> list[str]:
    text = clean(value)
    if not text or len(text) <= threshold:
        return [text]
    pattern = re.compile(
        r"(?m)(?=(?:^|\n)\s*(?:\d+[.、）)]|[①-⑳]|[一二三四五六七八九十]+[.、]))"
    )
    pieces = [piece.strip() for piece in pattern.split(text) if piece.strip()]
    if len(pieces) <= 1:
        pieces = [text]
    result: list[str] = []
    for piece in pieces:
        if len(piece) <= threshold:
            result.append(piece)
            continue
        start = 0
        minimum = max(16, threshold // 2)
        while start < len(piece):
            if len(piece) - start <= threshold:
                result.append(piece[start:].strip())
                break
            window_end = start + threshold
            cut = -1
            for index in range(window_end - 1, start + minimum - 1, -1):
                if piece[index] in "。；;！!？?，,\n":
                    cut = index + 1
                    break
            if cut <= start:
                cut = window_end
            result.append(piece[start:cut].strip())
            start = cut
    return [piece for piece in result if piece]


def _strip_suffix_keep_num(value: str) -> float | None:
    text = re.sub(r"\s+", "", clean(value))
    if not text:
        return None
    match = re.search(r"(\d+(?:\.\d+)?)\s*([万千百亿])?", text)
    if not match:
        return None
    number = float(match.group(1))
    scale = {"万": 10000, "千": 1000, "百": 100, "亿": 100000000}.get(
        match.group(2) or "", 1
    )
    return number * scale


def _cn_number(text: str) -> float | None:
    digits = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
              "六": 6, "七": 7, "八": 8, "九": 9}
    if not text:
        return None
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        return float(text)
    if text == "半":
        return 0.5
    total = 0
    section = 0
    current = 0
    for char in text:
        if char in digits:
            current = digits[char]
        elif char == "十":
            section += (current or 1) * 10
            current = 0
        elif char == "百":
            section += (current or 1) * 100
            current = 0
        elif char == "千":
            section += (current or 1) * 1000
            current = 0
        else:
            return None
    return float(total + section + current)


def _normalize_freq(value: str) -> str:
    """把“班 / 框 / 月”等简写频率还原成模板写法，避免频次规则失效。"""
    text = re.sub(r"\s+", "", clean(value))
    return _FREQ_ALIASES.get(text, text)


def compute_频次1_小时数(sample_qty: str, sample_freq: str) -> str:
    qty = re.sub(r"\s+", "", clean(sample_qty))
    freq = _normalize_freq(sample_freq)

    if ("每班" in freq or "补打条码时" in freq) and _is_every_class_series(qty):
        return "6" if "首中末" in qty else "12"
    if "开班" in freq or "换方向" in freq:
        return "12"
    # 每班抽检但未写首/末件时，按每班一次换算（连续抽检件数保持为空）
    if ("每班" in freq or "补打条码时" in freq) and "连续" not in qty:
        return "12"
    if freq in _BARE_TIME_FREQ:
        return _fmt(_BARE_TIME_FREQ[freq])

    for unit, coef in UNIT_COEF.items():
        pattern = rf"(半|\d+(?:\.\d+)?|[一二两三四五六七八九十百千]+)\s*{unit}"
        match = re.search(pattern, freq, re.IGNORECASE)
        if match:
            number = _cn_number(match.group(1))
            if number is None:
                number = 1
            return _fmt(number * coef)
    return ""


def _is_every_class_series(qty: str) -> bool:
    return any(key in qty for key in ("首3件", "首件", "首末件", "首中末"))


def _every_class_value(qty: str, freq: str) -> tuple[str, str, str] | None:
    if "每班" in freq or "补打条码时" in freq:
        if "首中末" in qty:
            match = re.search(r"首中末各?(?:连续)?(\d+)件", qty)
            sampling = _fmt(float(match.group(1))) if match else "1"
            return "首检;过程抽检;末检", "6", sampling
        if "首末件" in qty:
            return "首检;末检", "12", "1"
        if "首3件" in qty:
            return "首检", "12", "3"
        if "首件" in qty:
            if "补打条码时" in freq:
                return "首检;过程抽检", "12", "1"
            return "首检", "12", "1"
    return None


def compute_检验方式(sample_qty: str, sample_freq: str) -> str:
    qty = re.sub(r"\s+", "", clean(sample_qty))
    freq = _normalize_freq(sample_freq)

    table = _every_class_value(qty, freq)
    if table is not None:
        return table[0]
    # 100% 连续全检：无论“100%”写在样本容量还是样本频率，都按过程抽检处理
    if "100%" in qty + freq and "连续" in qty + freq:
        return "过程抽检"
    if ("每班" in freq or "补打条码时" in freq) and "连续" in qty and "首中末" not in qty:
        return "过程抽检"
    if any(
        key in freq
        for key in ("开班", "换方向", "作业前", "MSA分析计划及报告", "产品审核计划")
    ):
        return "首检"
    if any(
        key in freq
        for key in ("破损时", "使用前", "补料时", "每料盘", "每半车", "每框",
                    "打印时", "全尺寸检测计划及报告")
    ):
        return "过程抽检"
    if "防错" in freq:
        return "首检,过程抽检"
    if "型式试验" in freq:
        return "首检;过程抽检;末检"

    for prefix, mode in (
        ("首中末", "首检;过程抽检;末检"),
        ("首中", "首检;过程抽检"),
        ("首末", "首检;末检"),
        ("中末", "过程抽检;末检"),
    ):
        if prefix in qty and re.search(r"\d+件", qty):
            return mode
    if "首中末" in qty or "100%" in qty:
        return "首检;过程抽检;末检"
    if "首末件" in qty:
        return "首检;末检"
    # 样本频率为“连续”（如容量 1 / 100%）时按全检抽检处理，不得整行丢弃
    if "连续" in freq and not any(key in qty for key in ("连续", "首", "中", "末")):
        return "过程抽检"
    if "连续" in qty and not any(key in qty for key in ("首", "中", "末")):
        return "首检"
    # 频率写成“50件”这类数量时，表示每 50 件抽检一次
    if re.fullmatch(r"\d+件", freq):
        return "过程抽检"
    if any(key in qty for key in ("连续", "1次", "2件", "首件")):
        if "每" in freq or "作业前" in freq:
            return "首检"
        return "过程抽检"
    return ""


def compute_抽样数(sample_qty: str, sample_freq: str) -> str:
    qty = re.sub(r"\s+", "", clean(sample_qty))
    freq = _normalize_freq(sample_freq)
    if not qty:
        return ""

    table = _every_class_value(qty, freq)
    if table is not None:
        return table[2]
    # 样本容量或样本频率为 100%（含百分比格式的数值 1）时按全检 100 件
    if "100%" in qty or "100%" in freq:
        return "100"
    if any(key in qty for key in ("首件", "首末件", "首中末件", "一次", "MSA",
                                  "防错", "审核", "报告")):
        return "1"
    if "防错" in freq or "型式试验" in freq:
        return "1"

    match = re.search(r"首中末各(?:连续)?(\d+)件", qty)
    if match:
        return _fmt(float(match.group(1)))
    for prefix in ("首中末", "首中", "首末", "中末"):
        match = re.search(prefix + r"(\d+)件", qty)
        if match:
            number = float(match.group(1))
            return _fmt(number * 3) if prefix == "首中末" else _fmt(number * 2)
    if any(key in qty for key in ("1次", "1件", "首件", "末件")):
        return "1"
    if "首末件" in qty or "首中末件" in qty:
        return "1"
    if "首中末各连续3件" in qty or "首中末连续3件" in qty:
        return "3"
    match = re.search(r"连续(\d+)件", qty)
    if match and not any(key in qty for key in ("首", "中", "末")):
        return _fmt(float(match.group(1)))
    match = re.fullmatch(r"(\d+(?:\.\d+)?)件", qty)
    if match:
        return _fmt(float(match.group(1)))
    return ""


def compute_频次2(sample_qty: str, sample_freq: str, sampling: str) -> str:
    qty = re.sub(r"\s+", "", clean(sample_qty))
    freq = _normalize_freq(sample_freq)

    if "每班" in freq or "补打条码时" in freq:
        if _every_class_value(qty, freq) is not None:
            return ""
        if "连续" in qty and "首中末" not in qty:
            match = re.search(r"连续(\d+)件", qty)
            if match:
                return _fmt(float(match.group(1)))
    # 频率写成“50件”时，频次2 即该数量
    match = re.fullmatch(r"(\d+)件", freq)
    if match:
        return _fmt(float(match.group(1)))
    # 100% 全检：按每 1 件抽检 1 件填写（100% 写在频率列时同样适用）
    if "100%" in qty or "100%" in freq:
        return "1"
    if "防错" in freq:
        return "1"
    if "型式试验" in freq:
        return "1"
    if any(
        key in freq
        for key in ("破损时", "使用前", "补料时", "每料盘", "每半车", "每框",
                    "打印时", "全尺寸检测计划及报告")
    ):
        return clean(sampling)
    if any(
        key in freq
        for key in ("开班", "换方向", "作业前", "MSA分析计划及报告", "产品审核计划")
    ):
        return ""
    return ""


def compute_目标值(judge_type: str, target: str, upper: str, lower: str) -> str:
    if judge_type == "合格/不合格":
        return "合格"
    if target:
        return target
    if upper and not lower:
        return upper
    if lower and not upper:
        return lower
    return target


def build_quality_rows(source_rows: list[dict[str, str]]) -> list[dict[str, str]]:
    """把映射后的控制计划行转换为 22 列质量策划行。"""
    results: list[dict[str, str]] = []
    seen: set[tuple[str, ...]] = set()

    for source in source_rows:
        spec = clean(source.get("规范/公差", ""))
        process_feature = clean(source.get("过程特性", ""))
        ranges = build_range_cell_values(spec, process_feature)
        qty = source.get("样本数量", "")
        freq = source.get("样本频率", "")
        sampling = compute_抽样数(qty, freq)
        mode = compute_检验方式(qty, freq)
        freq1 = compute_频次1_小时数(qty, freq)
        freq2 = compute_频次2(qty, freq, sampling)
        if "过程抽检" not in mode and any(
            key in mode for key in ("首检", "首件", "末检")
        ):
            # §6.9：检验方式只含首检/首件/末检时，两个频次均为空
            freq1 = ""
            freq2 = ""
        elif "过程抽检" in mode and not freq1 and not freq2:
            # §6.5：过程抽检行频次1与频次2不得同时为空
            freq2 = sampling or "1"
        if not sampling:
            # §6.7：抽样数为空时取频次2，仍为空则默认 1
            sampling = freq2 or "1"
        for segment, target, upper, lower, unit in ranges:
            feature_desc = compute_特征描述_new(
                source.get("产品特性", ""), process_feature, segment
            )
            judge_type = compute_判定类型(segment, process_feature)
            if judge_type == "值判断":
                if not (target or upper or lower):
                    # 值判断必须至少有目标值/上限/下限之一，否则退回合格判定
                    judge_type = "合格/不合格"
                elif not unit:
                    unit = "N" if "阻尼力" in feature_desc else "mm"
            desc_parts = split_long_feature_desc(feature_desc)
            for desc_part in desc_parts:
                row = {
                    "工序号": compute_工序号_补位(source.get("过程编号（工序号）", "")),
                    "特性类型": compute_特性类型_new(
                        source.get("产品特性", ""), process_feature
                    ),
                    "特征描述（产品特性+范围公差）": desc_part,
                    "特殊特性": compute_特殊特性(
                        source.get("特殊特性分类（特殊特性）", "")
                    ),
                    "判定类型": judge_type,
                    "实验室检测": compute_实验室检测(
                        source.get("评价测量技术（评价测量技术）", "")
                    ),
                    "自动采集": compute_自动采集(
                        judge_type, source.get("评价测量技术（评价测量技术）", "")
                    ),
                    "目标值": compute_目标值(judge_type, target, upper, lower),
                    "上限": upper,
                    "下限": lower,
                    "单位": unit,
                    "检验方式": mode,
                    "频次1（时间/）": freq1,
                    "频次2（数量/件）": freq2,
                    "抽样数": sampling,
                    "评价测量技术": clean(
                        source.get("评价测量技术（评价测量技术）", "")
                    ),
                    "责任人": "操作者",
                    "数据保存方法": "/",
                    "检查单位": "现场",
                    "检验顺序": "",
                    "对应物料": "",
                    "备注": "",
                }
                if not row["检验方式"]:
                    continue
                key = (
                    row["工序号"],
                    clean(source.get("过程名称", "")),
                    row["特性类型"],
                    row["特征描述（产品特性+范围公差）"],
                    row["目标值"],
                    row["上限"],
                    row["下限"],
                    row["单位"],
                    row["检验方式"],
                    row["抽样数"],
                    row["频次1（时间/）"],
                    row["频次2（数量/件）"],
                )
                if key in seen:
                    continue
                seen.add(key)
                results.append(row)
    return results


def source_dedup_key(source: dict[str, str]) -> tuple[str, ...]:
    return (
        clean(source.get("过程编号（工序号）", "")),
        clean(source.get("产品特性", "")),
        clean(source.get("过程特性", "")),
        re.sub(r"\s+", "", clean(source.get("规范/公差", ""))),
        clean(source.get("样本频率", "")),
        clean(source.get("样本数量", "")),
        clean(source.get("机器、装置夹具、工装", "")),
    )
