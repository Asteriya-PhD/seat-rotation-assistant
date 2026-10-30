"""结构识别：从 xlsx_io.read_sheet() 的结果推断座位、大组、讲台与结构格。

设计原则：
- 只做“猜测”，所有猜测都通过 roles 暴露给前端，允许老师点选修正；
- 真实班级表的格式千变万化，识别管线按「文字语义 → 行带 → 组标签 → 兜底聚类」推进；
- 绝不把纯数字（排号）、门窗/讲台/走道、课代表区当成学生座位。
"""

from __future__ import annotations

import re

try:  # 包内导入（server/tests 均以 seatlib 包方式使用）
    from . import period as period_mod
except ImportError:  # 兼容直接以脚本方式导入
    import period as period_mod  # type: ignore

# ---------------------------------------------------------------- 关键词

PODIUM_KEYWORDS = ("讲台", "讲桌", "讲坛", "黑板", "白板", "屏幕", "投影", "多媒体", "电视")

STRUCT_KEYWORDS = (
    "门", "窗", "墙", "走道", "过道", "走廊", "通道", "楼梯", "卫生间", "厕所",
    "空调", "柜", "桌", "椅", "图书", "卫生角", "公告", "学习园地", "荣誉",
    "饮水", "绿植", "垃圾", "座位周期", "座位表", "座次表", "座位安排",
    "班级", "日期", "姓名", "学号", "备注", "黑板报", "音箱", "开关", "插座",
    "外侧", "内侧", "走廊", "讲台", "窗户", "后门", "前门", "外窗", "内窗",
)

SUBJECT_KEYWORDS = (
    "语文", "数学", "英语", "物理", "化学", "生物", "政治", "历史", "地理",
    "体育", "音乐", "美术", "信息", "心理", "通用技术", "班会",
    "课代表", "班主任", "班长", "组长", "团支书", "学习委员", "劳动委员",
    "体育委员", "文娱委员", "生活委员", "纪律委员", "宣传委员", "值日",
)

VOID_RE = re.compile(r"^(?:空|空位|空座|无人|——|—|–|-|/|／|×|x|X|·|\.|。|\?|？)+$")
GROUP_LABEL_RE = re.compile(r"^第?[0-9一二三四五六七八九十百零〇]+(?:大)?组$")
ROW_LABEL_RE = re.compile(r"^第?[0-9一二三四五六七八九十百零〇]+排$")
NUMBER_RE = re.compile(r"^[0-9]+(?:[.．]0+)?$")


def _norm(text):
    return re.sub(r"\s+", "", str(text or ""))


def _classify(norm_text, merge):
    """返回 kind：void / group / podium / struct / candidate。"""
    if not norm_text:
        return None
    if VOID_RE.match(norm_text):
        return "void"
    if GROUP_LABEL_RE.match(norm_text):
        return "group"
    if ROW_LABEL_RE.match(norm_text) or NUMBER_RE.match(norm_text):
        return "struct"
    for kw in PODIUM_KEYWORDS:
        if kw in norm_text:
            return "podium"
    for kw in SUBJECT_KEYWORDS:
        if kw in norm_text:
            return "struct"
    for kw in STRUCT_KEYWORDS:
        if kw in norm_text:
            return "struct"
    # 跨 3 列以上的合并单元格一般是标题/说明，不是学生座位
    if merge and int(merge.get("colspan") or 1) >= 3:
        return "struct"
    return "candidate"


# ---------------------------------------------------------------- 行带与主座位区

def _row_bands(coords, row_spans=None):
    """把候选格按“行带”聚类。

    纵向合并的座位格（如两行合并为一张课桌）会把 anchor 行的下一行留空，
    因此这里用合并后的行覆盖范围来计算连续性，避免把同一片座位区切成多条行带。
    """
    row_spans = row_spans or {}
    if not coords:
        return []
    intervals = []
    for coord in coords:
        span = int(row_spans.get(coord, 1) or 1)
        intervals.append((coord[0], coord[0] + span - 1))
    intervals.sort()
    merged = []
    for start, stop in intervals:
        if merged and start <= merged[-1][1] + 1:
            merged[-1] = (merged[-1][0], max(merged[-1][1], stop))
        else:
            merged.append((start, stop))
    bands = []
    for start, stop in merged:
        band = sorted(p for p in coords if start <= p[0] <= stop)
        if band:
            bands.append(band)
    return bands


def _pick_podium(podium_cells, cell_map):
    if not podium_cells:
        return None
    exact = [c for c in podium_cells
             if _norm(cell_map[c].get("text")) in ("讲台", "讲桌", "讲坛")]
    pool = exact or sorted(podium_cells, key=lambda p: (p[0], p[1]))
    return sorted(pool, key=lambda p: (p[0], p[1]))[0]


def _relation(podium_coord, seat_coords):
    if not podium_coord or not seat_coords:
        return None
    pr, pc = podium_coord
    rows = [p[0] for p in seat_coords]
    cols = [p[1] for p in seat_coords]
    r1, r2 = min(rows), max(rows)
    c1, c2 = min(cols), max(cols)
    if pr < r1:
        return "above"
    if pr > r2:
        return "below"
    if pc < c1:
        return "left"
    if pc > c2:
        return "right"
    return None


def _infer_relation_from_row_labels(cell_map, seat_coords):
    """讲台识别不到时，用座位左右两侧的纯数字排号推断讲台在上/下。"""
    if not seat_coords:
        return None
    rows = [p[0] for p in seat_coords]
    cols = [p[1] for p in seat_coords]
    r1, r2, c1, c2 = min(rows), max(rows), min(cols), max(cols)
    sides = {}
    for coord, cell in cell_map.items():
        r, c = coord
        text = _norm(cell.get("text"))
        if not NUMBER_RE.match(text):
            continue
        if r1 <= r <= r2 and (c < c1 or c > c2):
            sides.setdefault((c, "left" if c < c1 else "right"), []).append((r, int(float(text))))
    for key in sorted(sides):
        series = sorted(sides[key])
        values = [v for _, v in series]
        if len(values) < 2 or len(set(values)) < 2:
            continue
        increasing = values[-1] > values[0]
        # 排号一般从前往后 1..N；数字向下递增 ⇒ 前排在下方 ⇒ 讲台在下方
        # 数字向下递减（上大下小）⇒ 1 在下方 ⇒ 前排在下 ⇒ 讲台在下
        return "below" if not increasing else "above"
    return None


# ---------------------------------------------------------------- 大组

def _pick_label_row(label_cells, seat_coords):
    if not label_cells or not seat_coords:
        return []
    rows = [p[0] for p in seat_coords]
    r1, r2 = min(rows), max(rows)
    by_row = {}
    for label in label_cells:
        by_row.setdefault(label["r"], []).append(label)
    best = None
    best_key = None
    for row, items in by_row.items():
        # 同一列只保留一个标签
        dedup = {}
        for item in items:
            dedup.setdefault(item["c"], item)
        items = sorted(dedup.values(), key=lambda x: x["c"])
        dist = 0 if r1 - 2 <= row <= r2 + 2 else min(abs(row - r1), abs(row - r2))
        key = (len(items), -dist)
        if best_key is None or key > best_key:
            best_key = key
            best = items
    return best or []


def _cluster_columns(seat_cols):
    runs = []
    current = [seat_cols[0]]
    for col in seat_cols[1:]:
        if col == current[-1] + 1:
            current.append(col)
        else:
            runs.append(current)
            current = [col]
    runs.append(current)
    return runs


def _cn_number(number):
    digits = "零一二三四五六七八九"
    if number <= 10:
        return "十" if number == 10 else digits[number]
    if number < 20:
        return "十" + digits[number % 10]
    if number < 100:
        tail = digits[number % 10] if number % 10 else ""
        return digits[number // 10] + "十" + tail
    return str(number)


def _build_groups(seat_coords, label_cells):
    """返回 (groups, warnings)。groups 内每项含 name / cols / coords。

    规则：
    - 有组标签：合并标签按合并宽度覆盖列；单格标签延伸到下一个标签之前
      （因此「8组」「7组」两张相邻单格标签各占 1 列，不会把相邻列并成一组）；
    - 标签没有覆盖到的座位列：按连续列聚成独立大组；
    - 完全无标签：按连续座位列分组。
    """
    warnings = []
    if not seat_coords:
        return [], ["没有识别到任何座位"]
    seat_cols = sorted({c for _, c in seat_coords})
    min_col, max_col = seat_cols[0], seat_cols[-1]
    by_col = {}
    for coord in seat_coords:
        by_col.setdefault(coord[1], []).append(coord)

    nearby = []
    for label in label_cells or []:
        if min_col - 2 <= int(label["c"]) <= max_col + 2:
            nearby.append({"r": int(label["r"]), "c": int(label["c"]),
                           "text": str(label.get("text") or "").strip(),
                           "colspan": int(label.get("colspan") or 1)})
    chosen = _pick_label_row(nearby, seat_coords)

    entries = []
    covered = set()
    if chosen:
        chosen = sorted(chosen, key=lambda item: item["c"])
        for index, label in enumerate(chosen):
            col = label["c"]
            span = label["colspan"]
            if span > 1:
                cols = [c for c in seat_cols if col <= c <= col + span - 1]
            else:
                next_col = chosen[index + 1]["c"] if index + 1 < len(chosen) else None
                if next_col is None:
                    cols = [c for c in seat_cols if c >= col]
                else:
                    cols = [c for c in seat_cols if col <= c < next_col]
                if index == 0:
                    # 第一张标签左侧可能还有座位列，一并归入第一组
                    left = [c for c in seat_cols if c < col]
                    cols = left + cols
            cols = [c for c in cols if c not in covered]
            if not cols:
                warnings.append(f"大组标签「{label['text'] or '未命名'}」覆盖的列与其它组重复，已忽略")
                continue
            entries.append({"cols": cols, "name": label["text"] or None})
            covered.update(cols)

    rest = [c for c in seat_cols if c not in covered]
    if rest:
        runs = _cluster_columns(rest)
        for run in runs:
            entries.append({"cols": run, "name": None})
        if chosen:
            warnings.append(
                f"有 {len(rest)} 列座位没有组标签，已自动按连续列分为 {len(runs)} 个大组，可在界面调整"
            )
        else:
            warnings.append(
                f"未找到「N组/第N大组」标签，已按连续座位列自动分为 {len(runs)} 个大组，可在界面合并或拆分"
            )

    if not entries:
        runs = _cluster_columns(seat_cols)
        entries = [{"cols": run, "name": None} for run in runs]
        warnings.append("未能用组标签分组，已按连续座位列自动分组")

    entries.sort(key=lambda item: min(item["cols"]))
    groups = []
    for index, entry in enumerate(entries):
        coords = [coord for col in entry["cols"] for coord in by_col[col]]
        name = entry["name"] or f"第{_cn_number(index + 1)}组"
        groups.append({"name": name, "cols": entry["cols"], "coords": coords})

    groups.sort(key=lambda g: (min(g["cols"]), min(p[0] for p in g["coords"])))
    for index, group in enumerate(groups):
        group["id"] = index
    return groups, warnings


# ---------------------------------------------------------------- 主入口

def detect_sheet(sheet):
    """输入 xlsx_io.read_sheet() 的 sheet dict，输出 ARCHITECTURE.md §3.2 的结构。"""
    cells = sheet.get("cells") or []
    rows = int(sheet.get("rows") or 0)
    cols = int(sheet.get("cols") or 0)
    merges = sheet.get("merges") or []
    merge_anchor = {(int(m["r"]), int(m["c"])): m for m in merges}
    cell_map = {}
    for cell in cells:
        try:
            coord = (int(cell["r"]), int(cell["c"]))
        except (KeyError, TypeError, ValueError):
            continue
        cell_map[coord] = cell

    roles = {}
    candidates = []
    podium_cells = []
    label_cells = []

    for coord, cell in cell_map.items():
        text = cell.get("text") or ""
        norm = _norm(text)
        merge = merge_anchor.get(coord)
        if not norm:
            if merge and (int(merge.get("rowspan") or 1) > 1 or
                          int(merge.get("colspan") or 1) > 1):
                roles[coord] = "struct"
            continue
        kind = _classify(norm, merge)
        if kind == "void":
            roles[coord] = "void"
        elif kind == "group":
            roles[coord] = "struct"
            label_cells.append({"r": coord[0], "c": coord[1], "text": text,
                                "colspan": int(merge.get("colspan") or 1) if merge else 1})
        elif kind == "podium":
            roles[coord] = "struct"
            podium_cells.append(coord)
        elif kind == "struct":
            roles[coord] = "struct"
        else:
            candidates.append(coord)

    if not candidates:
        raise ValueError("没有识别到任何座位文字：请确认导入的是座位表，或检查表格是否被图片化")

    row_spans = {}
    for coord in candidates:
        merge = merge_anchor.get(coord)
        row_spans[coord] = int(merge.get("rowspan") or 1) if merge else 1
    bands = _row_bands(candidates, row_spans)
    main_band = max(bands, key=len)
    seat_coords = sorted(main_band)
    ignored_coords = sorted(c for band in bands for c in band if c not in set(main_band))

    for coord in seat_coords:
        roles[coord] = "seat"
    for coord in ignored_coords:
        roles[coord] = "ignored"

    # 讲台
    podium_coord = _pick_podium(podium_cells, cell_map)
    relation = _relation(podium_coord, seat_coords)
    if podium_coord and relation is None:
        relation = _infer_relation_from_row_labels(cell_map, seat_coords)
    if not podium_coord:
        relation = _infer_relation_from_row_labels(cell_map, seat_coords)
    podium = None
    if podium_coord:
        podium = {"r": podium_coord[0], "c": podium_coord[1],
                  "relation": relation or "above",
                  "text": str(cell_map[podium_coord].get("text") or "")}
    elif relation:
        # 没找到「讲台」文字，但能从左右两侧排号推断朝向
        podium = {"r": None, "c": None, "relation": relation, "text": "", "inferred": True}

    # 大组
    groups, group_warnings = _build_groups(seat_coords, label_cells)

    # 座位 → 组、组内相对坐标
    group_index = {}
    for group in groups:
        for coord in group["coords"]:
            group_index[coord] = group["id"]

    seat_rows_by_group = {}
    seat_cols_by_group = {}
    for group in groups:
        coords = group["coords"]
        seat_rows_by_group[group["id"]] = sorted({p[0] for p in coords})
        seat_cols_by_group[group["id"]] = sorted({p[1] for p in coords})

    seats = []
    for coord in seat_coords:
        gid = group_index.get(coord, 0)
        seats.append({
            "r": coord[0],
            "c": coord[1],
            "text": str(cell_map[coord].get("text") or ""),
            "group_id": gid,
            "row_in_group": seat_rows_by_group.get(gid, []).index(coord[0])
            if coord[0] in seat_rows_by_group.get(gid, []) else 0,
            "col_in_group": seat_cols_by_group.get(gid, []).index(coord[1])
            if coord[1] in seat_cols_by_group.get(gid, []) else 0,
        })

    group_out = []
    for group in groups:
        coords = group["coords"]
        rr = [p[0] for p in coords]
        cc = [p[1] for p in coords]
        group_out.append({
            "id": group["id"],
            "name": group["name"],
            "r1": min(rr), "r2": max(rr),
            "c1": min(cc), "c2": max(cc),
            "rows": len(set(rr)), "cols": len(set(cc)),
            "seat_count": len(coords),
            "seat_cols": sorted(set(cc)),
            "seat_count_by_col": {str(col): len([p for p in coords if p[1] == col])
                                  for col in sorted(set(cc))},
        })

    # ---- 警告
    warnings = list(group_warnings)
    if ignored_coords:
        ignored_rows = sorted({p[0] for p in ignored_coords})
        warnings.append(
            f"第 {ignored_rows[0]}~{ignored_rows[-1]} 行另有 {len(ignored_coords)} 个文字格"
            "（疑似课代表/名单区），已忽略；如其实是座位可在界面上改回"
        )
    if not podium_coord:
        warnings.append("未识别到「讲台/黑板」等文字，前后方向默认按「讲台在座位上方」计算，可在界面切换")
    elif relation is None:
        warnings.append("讲台位置在座位区域内，未能判断前后方向，请在界面确认")

    counts = sorted({g["seat_count"] for g in group_out})
    if len(counts) > 1:
        detail = "、".join(f"{g['name']}{g['seat_count']}人" for g in group_out[:8])
        warnings.append(f"各大组座位数不一致（{detail}），请确认是否正常（如单座/空位）")

    texts = {}
    for seat in seats:
        norm = _norm(seat["text"])
        texts.setdefault(norm, []).append(seat)
    duplicated = {k: v for k, v in texts.items() if len(v) > 1}
    if duplicated:
        sample = "、".join(list(duplicated)[:5])
        warnings.append(f"检测到 {len(duplicated)} 个重复姓名/文字（{sample}），请确认是否正常")

    period_info = period_mod.public_period(period_mod.find_period(sheet))

    return {
        "podium": podium,
        "seats": seats,
        "groups": group_out,
        "roles": {f"{r},{c}": role for (r, c), role in sorted(roles.items())},
        "period": period_info,
        "warnings": warnings,
    }


# ---------------------------------------------------------------- 用户改角色后重算大组

def regroup(seats, group_labels=None):
    """根据用户修正后的座位坐标与组标签重新计算大组。"""
    coords = []
    for seat in seats or []:
        try:
            coords.append((int(seat["r"]), int(seat["c"])))
        except (KeyError, TypeError, ValueError):
            continue
    coords = sorted(set(coords))
    if not coords:
        raise ValueError("没有座位可分组")

    labels = []
    for label in group_labels or []:
        text = label.get("text")
        if text is None:
            continue
        norm = _norm(text)
        if not GROUP_LABEL_RE.match(norm):
            continue
        try:
            labels.append({"r": int(label["r"]), "c": int(label["c"]),
                           "text": str(text),
                           "colspan": int(label.get("colspan") or 1)})
        except (KeyError, TypeError, ValueError):
            continue

    groups, warnings = _build_groups(coords, labels)
    # 只返回与契约一致的字段（附加 seat_cols 方便前端按列映射）
    out = []
    for group in groups:
        rr = [p[0] for p in group["coords"]]
        cc = [p[1] for p in group["coords"]]
        out.append({
            "id": group["id"], "name": group["name"],
            "r1": min(rr), "r2": max(rr), "c1": min(cc), "c2": max(cc),
            "rows": len(set(rr)), "cols": len(set(cc)),
            "seat_count": len(group["coords"]),
            "seat_cols": sorted(set(cc)),
        })
    return {"groups": out, "warnings": warnings}
