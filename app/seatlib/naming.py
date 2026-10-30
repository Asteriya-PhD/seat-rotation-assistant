"""导出文件命名：识别「第N轮」与文件名里的日期，按轮次智能递进。

规则（与老师约定）：

- 原文件名含「第N轮」→ 第 N+1 轮（中文数字保持中文、阿拉伯数字保持数字/补零风格）；
- 原文件名含日期 → 换成新轮次日期（单日期用开始日；日期范围用开始～结束，保留原分隔符/补零风格）；
- 两者都有 → 都替换；
- 都没有 → 单轮保持原名，多轮追加 ``_第k轮`` 兜底避免重名；
- 所有返回值都不含 ``.xlsx`` 扩展名。
"""

from __future__ import annotations

import re

try:  # 包内导入
    from . import period as period_mod
except ImportError:  # 兼容脚本方式
    import period as period_mod  # type: ignore

ROUND_RE = re.compile(
    r"第\s*(?P<num>[0-9]{1,4}|[零〇一二两三四五六七八九十百]{1,5})\s*轮"
)
CN_DIGITS = "零一二三四五六七八九"
CN_MAP = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
          "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def strip_extension(stem: str) -> str:
    text = str(stem or "").strip()
    if text.lower().endswith(".xlsx"):
        text = text[:-5]
    return text


def parse_cn_number(text: str) -> int:
    total = 0
    section = 0
    number = 0
    for char in text:
        if char in CN_MAP:
            number = CN_MAP[char]
        elif char == "十":
            if number == 0:
                number = 1
            section += number * 10
            number = 0
        elif char == "百":
            if number == 0:
                number = 1
            section += number * 100
            number = 0
    return section + number


def format_cn_number(number: int) -> str:
    if number < 0:
        return str(number)
    if number < 10:
        return CN_DIGITS[number]
    if number == 10:
        return "十"
    if number < 20:
        return "十" + CN_DIGITS[number % 10]
    if number < 100:
        tail = CN_DIGITS[number % 10] if number % 10 else ""
        return CN_DIGITS[number // 10] + "十" + tail
    return str(number)


def format_round_number(raw: str, number: int) -> str:
    """按原数字风格格式化新的轮次数字。"""
    if raw.isdigit():
        if len(raw) > 1 and raw.startswith("0"):
            return str(number).zfill(len(raw))
        return str(number)
    return format_cn_number(number)


def analyze_filename(stem: str) -> dict:
    """解析文件名中的轮次与日期。"""
    stem = strip_extension(stem)
    round_match = ROUND_RE.search(stem)
    round_num = None
    round_style = None
    if round_match:
        raw = round_match.group("num")
        if raw.isdigit():
            round_num = int(raw)
            round_style = "arabic"
        else:
            round_num = parse_cn_number(raw)
            round_style = "cn"

    dates = []
    for match in period_mod.DATE_RE.finditer(stem):
        try:
            value, style, explicit_year = period_mod._parse_date_match(match)
        except ValueError:
            continue
        dates.append({
            "match": match,
            "date": value,
            "style": style,
            "explicit_year": explicit_year,
        })
        if len(dates) == 2:
            break

    if dates:
        # 补零风格以整段文件名里是否出现过前导零为准，避免起止各判各的
        pad_month = any(item["style"].get("pad_month") for item in dates)
        pad_day = any(item["style"].get("pad_day") for item in dates)
        for item in dates:
            item["style"]["pad_month"] = bool(pad_month)
            item["style"]["pad_day"] = bool(pad_day)

    if len(dates) >= 2:
        date_mode = "range"
    elif dates:
        date_mode = "single"
    else:
        date_mode = "none"

    return {
        "stem": stem,
        "round_match": round_match,
        "round": round_num,
        "round_style": round_style,
        "dates": dates,
        "date_mode": date_mode,
    }


def public_analysis(analysis: dict) -> dict:
    return {
        "stem": analysis["stem"],
        "round": analysis["round"],
        "round_style": analysis["round_style"],
        "date_mode": analysis["date_mode"],
        "dates": [item["date"].isoformat() for item in analysis["dates"]],
    }


def _render_date(item: dict, value, start_value=None) -> str:
    style = dict(item["style"])
    with_year = bool(style.get("has_year", True))
    # 文件名里结束日期常省略年份；跨年时必须补上，避免歧义
    if (not with_year and start_value is not None and value.year != start_value.year):
        with_year = True
        style["year_sep"] = style.get("year_sep") or "."
    return period_mod._fmt_date(value, style, with_year)


def suggest_filenames(stem: str, plans=None, rounds=None) -> list:
    """按轮次方案生成建议文件名（不含 .xlsx）。

    plans: [{"start": date|None, "end": date|None}, ...]
    rounds: 轮数；缺省取 len(plans)
    """
    analysis = analyze_filename(stem)
    plans = plans or []
    count = int(rounds or len(plans) or 1)
    count = max(1, count)
    base_stem = analysis["stem"]
    round_match = analysis["round_match"]
    dates = analysis["dates"]

    results = []
    for index in range(count):
        text = base_stem
        plan = plans[index] if index < len(plans) else {}
        start_value = plan.get("start")
        end_value = plan.get("end")
        replacements = []

        if round_match and analysis["round"] is not None:
            # 原文件是第 N 轮，导出的第 1 个文件应是第 N+1 轮
            new_number = analysis["round"] + index + 1
            raw = round_match.group("num")
            replacements.append((
                round_match.start(), round_match.end(),
                "第" + format_round_number(raw, new_number) + "轮",
            ))

        if dates and start_value is not None:
            for date_index, item in enumerate(dates):
                if date_index == 0:
                    new_value = start_value
                else:
                    new_value = end_value if end_value is not None else start_value
                rendered = _render_date(item, new_value, start_value)
                replacements.append((item["match"].start(), item["match"].end(), rendered))

        for start, stop, replacement in sorted(replacements, key=lambda x: x[0], reverse=True):
            text = text[:start] + replacement + text[stop:]

        if not round_match and not dates and count > 1:
            text = base_stem + f"_第{index + 1}轮"
        results.append(text)
    return results
