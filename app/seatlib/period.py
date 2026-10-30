"""座位周期（表头日期）识别与自动顺延。

支持识别形如以下写法的周期文字（通常位于表头合并格，如 F1:J1）：

- ``座位周期：2026.10.30～11.13``
- ``座位周期：2026年10月30日—11月13日``
- ``2026-10-30 ~ 2026-11-13``
- ``2026/10/30-11/13``

顺延规则：默认从原周期结束的**次日**开始，保持原周期长度（含首尾）；跨年时自动补年份，
并尽量保留原来的日期写法、前缀与分隔符。
"""

from __future__ import annotations

import datetime as _dt
import re

PERIOD_KEYWORDS = ("座位周期", "座次周期", "座位表周期", "座位安排周期")

# 允许“2026.10.30”“2026年10月30日”“10.30”“11.13”等写法
DATE_RE = re.compile(
    r"(?:(?P<y>\d{4})\s*(?P<ys>[./\-年])\s*)?"
    r"(?P<m>\d{1,2})\s*(?P<ms>[./\-月])\s*"
    r"(?P<d>\d{1,2})(?P<dsu>\s*日)?"
)

MAX_ROUNDS = 12


def _today():
    return _dt.date.today()


def _padded(sample: str) -> bool:
    """原始日期里出现前导零（如 01）才认为需要补零；10/30 这种不算。"""
    return len(sample) == 2 and sample.startswith("0")


def _parse_date_match(match):
    """返回 (date, style, year_explicit, raw_year)。"""
    year_text = match.group("y")
    year_explicit = bool(year_text)
    style = {
        "has_year": year_explicit,
        "year_sep": match.group("ys") or ".",
        "month_sep": match.group("ms") or ".",
        "day_suffix": "日" if match.group("dsu") else "",
        "pad_month": _padded(match.group("m")),
        "pad_day": _padded(match.group("d")),
    }
    month = int(match.group("m"))
    day = int(match.group("d"))
    if not year_explicit:
        year = _today().year
    else:
        year = int(year_text)
    try:
        value = _dt.date(year, month, day)
    except ValueError as exc:
        raise ValueError(f"日期不合法：{match.group(0)!r}") from exc
    return value, style, year_explicit


def _fmt_date(value: _dt.date, style: dict, with_year: bool) -> str:
    parts = []
    if with_year:
        parts.append(f"{value.year}{style['year_sep']}")
    month = f"{value.month:02d}" if style.get("pad_month") else str(value.month)
    day = f"{value.day:02d}" if style.get("pad_day") else str(value.day)
    return "".join(parts) + month + style["month_sep"] + day + style["day_suffix"]


def _parse_period_text(text: str):
    matches = list(DATE_RE.finditer(text))
    if len(matches) < 2:
        return None
    first, second = matches[0], matches[1]
    try:
        start, start_style, start_has_year = _parse_date_match(first)
        end, end_style, end_has_year = _parse_date_match(second)
    except ValueError:
        return None
    # 结束日期没有写年份且早于开始日期 ⇒ 视为跨年
    if not end_has_year and end < start:
        try:
            end = end.replace(year=end.year + 1)
        except ValueError:  # 2 月 29 日之类
            end = end.replace(year=end.year + 1, day=28)
    if end < start:
        return None
    # 补零风格按「整段文字里是否出现过前导零」统一，避免开始/结束各判各的
    pad_month = start_style.get("pad_month") or end_style.get("pad_month")
    pad_day = start_style.get("pad_day") or end_style.get("pad_day")
    for style in (start_style, end_style):
        style["pad_month"] = bool(pad_month)
        style["pad_day"] = bool(pad_day)
    template = {
        "prefix": text[: first.start()],
        "mid": text[first.end(): second.start()],
        "suffix": text[second.end():],
        "start_style": start_style,
        "end_style": end_style,
    }
    return {"text": text, "start": start, "end": end, "template": template}


def _render(period: dict, start: _dt.date, end: _dt.date) -> str:
    template = period["template"]
    start_style = template["start_style"]
    end_style = template["end_style"]
    start_text = _fmt_date(start, start_style, start_style.get("has_year", True))
    # 跨年时结束日期必须带年份；否则沿用原文是否有年份
    end_has_year = bool(end_style.get("has_year")) or end.year != start.year
    if end_has_year and not end_style.get("has_year"):
        end_style = dict(end_style)
        end_style["has_year"] = True
        if not end_style.get("year_sep"):
            end_style["year_sep"] = start_style.get("year_sep") or "."
    end_text = _fmt_date(end, end_style, end_has_year)
    return template["prefix"] + start_text + template["mid"] + end_text + template["suffix"]


def find_period(sheet: dict):
    """在工作表里寻找「座位周期」文字并解析。找不到返回 None。"""
    cells = sheet.get("cells") or []
    for cell in sorted(cells, key=lambda c: (int(c.get("r") or 0), int(c.get("c") or 0))):
        text = str(cell.get("text") or "")
        if not text or not any(keyword in text for keyword in PERIOD_KEYWORDS):
            continue
        parsed = _parse_period_text(text)
        if not parsed:
            continue
        # 用解析结果重画原文，能对得上才认为格式被完整理解
        try:
            rebuilt = _render(parsed, parsed["start"], parsed["end"])
        except (KeyError, ValueError):
            continue
        if rebuilt != text:
            # 允许空白差异（Excel 里偶有不可见空格），但结构必须一致
            if re.sub(r"\s+", "", rebuilt) != re.sub(r"\s+", "", text):
                continue
        r, c = int(cell.get("r") or 0), int(cell.get("c") or 0)
        return {
            "ref": _a1(r, c),
            "r": r,
            "c": c,
            "text": text,
            "start": parsed["start"],
            "end": parsed["end"],
            "days": (parsed["end"] - parsed["start"]).days + 1,
            "template": parsed["template"],
        }
    return None


def public_period(period):
    """去掉内部模板，返回给前端/API 的字段。"""
    if not period:
        return None
    return {
        "ref": period["ref"],
        "r": period["r"],
        "c": period["c"],
        "text": period["text"],
        "start": period["start"].isoformat(),
        "end": period["end"].isoformat(),
        "days": period["days"],
    }


def _parse_start_date(value):
    if not value:
        return None
    if isinstance(value, _dt.date):
        return value
    text = str(value).strip()
    try:
        return _dt.date.fromisoformat(text)
    except ValueError as exc:
        raise ValueError("开始日期格式不正确，请使用 YYYY-MM-DD（例如 2026-11-14）") from exc


def plan(period, rounds=1, start_date=None, cycle_weeks=2):
    """生成逐轮周期。

    period: find_period 的内部结构（含 template）
    rounds: 轮数（1~12）
    start_date: 第 1 轮开始日期；缺省 = 原周期结束的次日
    cycle_weeks: 每轮间隔周数（默认 2 周 = 14 天），周期长度 = cycle_weeks × 7 天

    返回 [{"round", "start", "end", "text"}]。
    """
    if not period:
        return []
    try:
        rounds = int(rounds)
    except (TypeError, ValueError):
        rounds = 1
    rounds = max(1, min(MAX_ROUNDS, rounds))
    try:
        cycle_weeks = int(cycle_weeks)
    except (TypeError, ValueError):
        cycle_weeks = 2
    cycle_weeks = max(1, min(8, cycle_weeks))
    duration = cycle_weeks * 7
    base_start = _parse_start_date(start_date)
    if base_start is None:
        base_start = period["end"] + _dt.timedelta(days=1)

    result = []
    for index in range(rounds):
        current_start = base_start + _dt.timedelta(days=duration * index)
        current_end = current_start + _dt.timedelta(days=duration - 1)
        result.append({
            "round": index + 1,
            "start": current_start.isoformat(),
            "end": current_end.isoformat(),
            "days": duration,
            "text": _render(period, current_start, current_end),
        })
    return result


def _a1(row: int, col: int) -> str:
    letters = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"
