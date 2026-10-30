"""轮转引擎（纯函数，无第三方依赖）。

输入：座位列表、大组列表、讲台方位、语义化规则；
输出：第 1..N 轮的落座映射，映射的 `from` 始终指向**原始（第 0 轮）座位**，
便于导出时把原始单元格的富文本/样式原样搬运到目标格。

方向语义（中文界面直接使用）：
- ``row_move.toward == "back"``  向教室后方（远离讲台）
- ``podium.relation == "below"`` 表示讲台画在座位下方，因此「后方」= 表格行号减小
"""

from __future__ import annotations

import math

MODES = ("inner", "class", "swap", "combo")
MAX_ROUNDS = 12


# ---------------------------------------------------------------- 基础工具

def _as_coord(value):
    """接受 [r, c] / (r, c) / "r,c" / "C3" 之外的坐标表达（导出用 A1 由调用方转换）。"""
    if isinstance(value, str):
        text = value.strip().replace(" ", "")
        if "," in text:
            r, c = text.split(",")
            return int(r), int(c)
        raise ValueError(f"无法识别的坐标：{value!r}（应为 [r, c] 或 'r,c'）")
    if isinstance(value, (list, tuple)) and len(value) == 2:
        return int(value[0]), int(value[1])
    raise ValueError(f"无法识别的坐标：{value!r}")


def _bbox(coords):
    rows = [p[0] for p in coords]
    cols = [p[1] for p in coords]
    return min(rows), max(rows), min(cols), max(cols)


def _row_sign(podium):
    """「向教室后方」在表格里的行增量符号。"""
    relation = (podium or {}).get("relation") or "above"
    return -1 if relation == "below" else 1


def _grid_delta(rules, podium):
    rm = rules.get("row_move") or {}
    cm = rules.get("col_move") or {}
    row_steps = max(0, min(3, int(rm.get("steps") or 0)))
    col_steps = max(0, min(3, int(cm.get("steps") or 0)))
    row_toward = rm.get("toward") or "back"
    col_toward = cm.get("toward") or "right"

    sign = _row_sign(podium)
    dr = 0
    if row_steps:
        dr = row_steps * sign if row_toward == "back" else -row_steps * sign
    dc = 0
    if col_steps:
        dc = col_steps if col_toward == "right" else -col_steps
    return dr, dc


def _group_shift(rules):
    gm = rules.get("group_move") or {}
    steps = max(0, min(3, int(gm.get("steps") or 0)))
    if not steps:
        return 0
    return steps if (gm.get("toward") or "right") == "right" else -steps


def _norm_fixed_seats(rules):
    fixed = set()
    for item in rules.get("fixed_seats") or []:
        fixed.add(_as_coord(item))
    return fixed


def _norm_fixed_groups(rules):
    result = set()
    for item in rules.get("fixed_groups") or []:
        try:
            result.add(int(item))
        except (TypeError, ValueError):
            continue
    return result


# ---------------------------------------------------------------- 组内/全班循环位移

def _inner_map(seat_set, holes, bbox, dr, dc, warnings=None, tag=""):
    """目标格 -> 原始格 的循环位移映射（跳过空洞/固定座位）。

    目标格 t 的学生来自：从 t 沿 (-dr, -dc) 方向、在 bbox 环面上回退，
    遇到的第一个真实座位。固定座位视为空洞，因此固定座位自身不参与。
    """
    moves = set(seat_set) - set(holes)
    mapping = {}
    if not moves:
        return mapping
    if dr == 0 and dc == 0:
        return {p: p for p in moves}

    r1, r2, c1, c2 = bbox
    height = r2 - r1 + 1
    width = c2 - c1 + 1
    if height <= 0 or width <= 0:
        return {p: p for p in moves}

    period_h = height // math.gcd(height, abs(dr)) if dr else 1
    period_w = width // math.gcd(width, abs(dc)) if dc else 1

    for target in sorted(moves):
        found = None
        for k in range(1, period_h * period_w + 1):
            rr = r1 + ((target[0] - r1 - dr * k) % height)
            cc = c1 + ((target[1] - c1 - dc * k) % width)
            if (rr, cc) in moves:
                found = (rr, cc)
                break
        if found is None:  # 理论上不可达
            mapping[target] = target
            if warnings is not None:
                warnings.append(f"{tag}座位 {target} 未找到轮转来源，已保持原地".strip())
        else:
            mapping[target] = found
    return mapping


# ---------------------------------------------------------------- 整组互换

def _group_geometry(coords):
    """返回 (排序后的座位列表, rel->seat 映射, (行数, 列数))，rel 用组内序号表示。"""
    coords = sorted(coords)
    rows = sorted({p[0] for p in coords})
    cols = sorted({p[1] for p in coords})
    row_rank = {r: i for i, r in enumerate(rows)}
    col_rank = {c: i for i, c in enumerate(cols)}
    rel_to_seat = {(row_rank[r], col_rank[c]): (r, c) for r, c in coords}
    return coords, rel_to_seat, (len(rows), len(cols))


def _swap_map(all_seats, seats_by_group, groups, fixed_groups, shift, fixed_seats, warnings):
    """整组互换：目标格 -> 源格。返回 (mapping, unplaced_sources, empty_targets)。"""
    if shift == 0:
        return {p: p for p in all_seats}, [], []

    rotatable = [g for g in groups if g["id"] not in fixed_groups]
    if len(rotatable) < 2:
        warnings.append("可轮转的大组少于 2 个，整组互换不会产生位移")
        return {p: p for p in all_seats}, [], []

    # 按物理位置从左到右排序
    rotatable.sort(key=lambda g: (g.get("c1", 0), g.get("r1", 0)))
    n = len(rotatable)
    effective = shift % n
    if effective == 0:
        warnings.append("整组移动的步长等于大组数量，整组互换后位置不变")
        return {p: p for p in all_seats}, [], []

    mapping = {}
    unplaced = []
    empty_targets = []
    shape_warnings = []

    for index, target_group in enumerate(rotatable):
        source_group = rotatable[(index - effective) % n]
        target_coords = [p for p in seats_by_group.get(target_group["id"], [])
                         if p not in fixed_seats]
        source_coords = [p for p in seats_by_group.get(source_group["id"], [])
                         if p not in fixed_seats]

        if not target_coords:
            continue

        t_list, t_rel, t_shape = _group_geometry(target_coords)
        s_list, s_rel, s_shape = _group_geometry(source_coords)

        if t_rel.keys() == s_rel.keys():
            # 形状一致：按组内相对位置搬
            for rel, target in t_rel.items():
                source = s_rel[rel]
                mapping[target] = source
        else:
            shape_warnings.append(
                f"大组「{target_group.get('name') or target_group['id']}」与来源组形状不同，"
                f"按行列顺序对应"
            )
            for k, target in enumerate(t_list):
                if k < len(s_list):
                    mapping[target] = s_list[k]
                else:
                    empty_targets.append(target)

    # 固定座位保持原地
    for p in fixed_seats:
        mapping[p] = p
    # 固定大组保持原地
    for g in groups:
        if g["id"] in fixed_groups:
            for p in seats_by_group.get(g["id"], []):
                mapping[p] = p

    used = {s for s in mapping.values() if s is not None}
    for p in sorted(all_seats):
        if p not in used:
            unplaced.append(p)

    if shape_warnings:
        warnings.extend(shape_warnings)
    return mapping, unplaced, empty_targets


# ---------------------------------------------------------------- 主入口

def _auto_fill(base, seat_set, holes):
    """整组形状不一致时，把未能随组落座的学生就近补进空位（保持一一对应）。"""
    used = {s for s in base.values() if s is not None}
    empty = sorted(t for t in seat_set - set(base) - holes)
    unplaced = sorted(p for p in seat_set - used - holes)
    if not empty or not unplaced:
        return 0
    filled = 0
    for source in unplaced:
        target = min(empty, key=lambda t: (abs(t[0] - source[0]) + abs(t[1] - source[1]), t))
        base[target] = source
        empty.remove(target)
        filled += 1
    return filled


def rotate(seats, groups, podium, rules):
    """执行 1..N 轮轮转。

    seats: [{"r":int,"c":int,"text":str,"group_id":int}, ...]
    groups: [{"id":int,"name":str,"r1","r2","c1","c2"}, ...]
    podium: {"relation": "above"|"below"|"left"|"right"|None}
    rules: 见 ARCHITECTURE.md §3.3
    """
    if not seats:
        raise ValueError("没有可轮转的座位：请先识别座位")

    rules = rules or {}
    mode = rules.get("mode") or "inner"
    if mode not in MODES:
        raise ValueError(f"未知的轮转模式：{mode!r}")

    # 座位坐标去重校验
    coord_text = {}
    for s in seats:
        try:
            coord = (int(s["r"]), int(s["c"]))
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"座位坐标不合法：{s!r}") from exc
        if coord in coord_text:
            raise ValueError(f"座位 {coord} 重复出现，无法轮转")
        coord_text[coord] = "" if s.get("text") is None else str(s.get("text"))

    seat_set = set(coord_text)
    fixed_seats = _norm_fixed_seats(rules)
    fixed_groups = _norm_fixed_groups(rules)

    # 组信息兜底：没有组时视为 1 个全班大组
    if not groups:
        r1, r2, c1, c2 = _bbox(seat_set)
        groups = [{"id": 0, "name": "全班", "r1": r1, "r2": r2, "c1": c1, "c2": c2}]
        warnings_no_group = True
    else:
        warnings_no_group = False

    # 座位 → 组（以 seats[].group_id 为主；缺失/无效时按组的列范围兜底）
    provided = []
    for g in groups:
        try:
            gid = int(g["id"])
        except (KeyError, TypeError, ValueError):
            continue
        provided.append({"id": gid, "name": g.get("name") or f"第{gid + 1}组",
                         "c1": g.get("c1"), "c2": g.get("c2"), "r1": g.get("r1")})
    provided_ids = {g["id"] for g in provided}

    raw_group_ids = {}
    for s in seats:
        coord = (int(s["r"]), int(s["c"]))
        try:
            raw_group_ids[coord] = int(s.get("group_id", 0))
        except (TypeError, ValueError):
            raw_group_ids[coord] = 0
    for coord, gid in list(raw_group_ids.items()):
        if gid in provided_ids:
            continue
        for g in provided:
            if g.get("c1") is None or g.get("c2") is None:
                continue
            if int(g["c1"]) <= coord[1] <= int(g["c2"]):
                raw_group_ids[coord] = g["id"]
                break

    warnings = []
    group_by_id = {}
    for g in provided:
        coords = sorted(p for p, gid in raw_group_ids.items() if gid == g["id"])
        if not coords:
            continue
        bbox = _bbox(coords)
        group_by_id[g["id"]] = {
            "id": g["id"],
            "name": g["name"],
            "coords": coords,
            "bbox": bbox,
            "c1": g["c1"] if g.get("c1") is not None else bbox[2],
            "r1": g["r1"] if g.get("r1") is not None else bbox[0],
        }
    if not group_by_id:
        raise ValueError("大组信息无效：没有任何座位归属到大组")

    missing = seat_set - {p for info in group_by_id.values() for p in info["coords"]}
    if missing:
        for coord in sorted(missing):
            nearest = min(group_by_id.values(),
                          key=lambda info: min(abs(coord[1] - p[1]) + abs(coord[0] - p[0])
                                               for p in info["coords"]))
            nearest["coords"] = sorted(nearest["coords"] + [coord])
            nearest["bbox"] = _bbox(nearest["coords"])
        warnings.append(f"有 {len(missing)} 个座位未归属大组，已按就近原则分配，建议在界面核对")

    seat_to_group = {}
    for gid, info in group_by_id.items():
        for p in info["coords"]:
            seat_to_group[p] = gid
    seats_by_group = {}
    for coord, gid in seat_to_group.items():
        seats_by_group.setdefault(gid, []).append(coord)

    relation = (podium or {}).get("relation")
    if relation in ("left", "right"):
        warnings.append("讲台在座位侧面，前后方向暂按「讲台在上方」计算，如有需要请改讲台方位")
    if warnings_no_group:
        warnings.append("未提供大组信息，已把全部座位视为一个整体")

    dr, dc = _grid_delta(rules, podium)
    group_shift = _group_shift(rules)

    # 固定大组：整组所有座位都不动
    holes = set(fixed_seats)
    for gid in fixed_groups:
        info = group_by_id.get(gid)
        if info is None:
            warnings.append(f"固定大组 {gid} 不存在，已忽略")
            continue
        holes.update(info["coords"])

    # ---- 计算基础置换（目标 -> 原始）
    if mode in ("inner", "class"):
        if group_shift:
            warnings.append("「组内/全班」模式下忽略整组移动参数，如需整组换位请用「整组互换」或「整组+组内」")
        if mode == "inner":
            base = {}
            for gid, info in group_by_id.items():
                group_seats = set(info["coords"])
                base.update(_inner_map(group_seats, holes, info["bbox"], dr, dc,
                                       warnings, tag=f"{info['name']}："))
        else:
            if fixed_groups:
                warnings.append("「全班整体」模式下固定大组的座位同样保持不动")
            base = _inner_map(seat_set, holes, _bbox(seat_set), dr, dc, warnings)
        unplaced_sources, empty_targets = [], []
        if dr == 0 and dc == 0:
            warnings.append("没有设置任何移动，结果与原座位表相同")
        else:
            # 组内一列时左右方向无位移提示
            if dc and mode == "inner":
                single_col = [info["name"] for info in group_by_id.values()
                              if len({p[1] for p in info["coords"]}) == 1]
                if single_col and len(single_col) == len(group_by_id):
                    warnings.append("所有大组都只有 1 列，左右方向的组内轮转不会产生位移")
                elif single_col:
                    warnings.append("部分大组只有 1 列（" + "、".join(single_col[:6]) +
                                    "），这些组内左右方向不会产生位移")

    elif mode == "swap":
        if dr or dc:
            warnings.append("「整组互换」模式下忽略前后/左右移动参数；如需同时组内移动请用「整组+组内」")
        if not group_by_id:
            raise ValueError("整组互换需要至少一个大组")
        base, unplaced_sources, empty_targets = _swap_map(
            seat_set, seats_by_group,
            [{"id": gid, "name": info["name"], "c1": info["c1"], "r1": info["r1"]}
             for gid, info in group_by_id.items()],
            fixed_groups, group_shift, fixed_seats, warnings)

    else:  # combo
        swap_base, unplaced_sources, empty_targets = _swap_map(
            seat_set, seats_by_group,
            [{"id": gid, "name": info["name"], "c1": info["c1"], "r1": info["r1"]}
             for gid, info in group_by_id.items()],
            fixed_groups, group_shift, fixed_seats, warnings)
        inner_by_group = {}
        for gid, info in group_by_id.items():
            inner_by_group[gid] = _inner_map(set(info["coords"]), holes, info["bbox"],
                                             dr, dc, warnings, tag=f"{info['name']}：")
        base = {}
        for target in seat_set:
            source = swap_base.get(target)
            if source is None:
                continue
            gid = seat_to_group.get(source)
            if gid is None:
                base[target] = source
            else:
                base[target] = inner_by_group[gid].get(source, source)
        if not dr and not dc and not group_shift:
            warnings.append("没有设置任何移动，结果与原座位表相同")

    # 固定座位/固定组补齐恒等映射
    for p in holes:
        base[p] = p

    # 大组形状不一致时自动补位（真实班级里常见：某组多一个单座）
    filled = _auto_fill(base, seat_set, holes)
    if filled:
        warnings.append(
            f"大组形状不一致，{filled} 名同学无法随组就位，已按最近距离自动安排到空位；"
            "请在预览里核对这几个人的位置"
        )

    # ---- 多轮：置换复合
    rounds = rules.get("rounds", 1)
    try:
        rounds = int(rounds)
    except (TypeError, ValueError):
        rounds = 1
    if rounds < 1 or rounds > MAX_ROUNDS:
        warnings.append(f"轮转次数应在 1~{MAX_ROUNDS} 之间，已自动调整为 {max(1, min(MAX_ROUNDS, rounds))}")
        rounds = max(1, min(MAX_ROUNDS, rounds))

    round_maps = []
    base_perm = dict(base)
    current = dict(base_perm)
    for _ in range(rounds):
        round_maps.append(dict(current))
        current = {t: base_perm.get(s) if s is not None else None
                   for t, s in current.items()}

    text_by_coord = {p: coord_text[p] for p in seat_set}
    ordered_targets = sorted(seat_set)

    out_rounds = []
    for index, mapping in enumerate(round_maps, start=1):
        assignments = []
        seats_after = []
        empty_seats = []
        used = set()
        for target in ordered_targets:
            source = mapping.get(target)
            if source is None or source not in seat_set:
                empty_seats.append(list(target))
                assignments.append({"to": list(target), "from": None, "text": ""})
                seats_after.append({"r": target[0], "c": target[1], "text": ""})
                continue
            used.add(source)
            assignments.append({"to": list(target), "from": list(source),
                                "text": text_by_coord[source]})
            seats_after.append({"r": target[0], "c": target[1],
                                "text": text_by_coord[source]})

        unplaced = [{"from": list(p), "text": text_by_coord[p]}
                    for p in sorted(seat_set - used)]

        # 一一对应校验
        sources = [s for s in mapping.values() if s is not None]
        if len(sources) != len(set(sources)):
            warnings.append(f"第 {index} 轮存在重复落座，请检查大组形状/固定设置")
        if len(unplaced) != len(empty_seats):
            warnings.append(
                f"第 {index} 轮有 {len(unplaced)} 人未落座、{len(empty_seats)} 个空位，"
                "人数校验不一致，请检查大组形状或固定设置"
            )

        out_rounds.append({
            "round": index,
            "assignments": assignments,
            "seats_after": seats_after,
            "unplaced": unplaced,
            "empty_seats": empty_seats,
            "warnings": [],
        })

    # 全局去重警告
    seen = set()
    deduped = []
    for w in warnings:
        if w not in seen:
            seen.add(w)
            deduped.append(w)

    return {"rounds": out_rounds, "warnings": deduped}


def _seat_group_id(seats, coord):
    for s in seats:
        if (int(s["r"]), int(s["c"])) == coord:
            try:
                return int(s.get("group_id", 0))
            except (TypeError, ValueError):
                return 0
    return 0
