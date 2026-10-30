"""xlsx 高保真导出（纯 Python 标准库，零第三方依赖）。

核心思路（ARCHITECTURE §4.2）：**不用 openpyxl 重建工作簿**，而是在原始 xlsx 的
目标 sheet XML 上做字符串级原位替换：

1. 一次性扫描并快照目标 sheet 中**所有** <c> 元素（r/t/s 属性、内部 XML、位置），
   再统一写回目标格 —— 避免「循环置换」时后写的 move 读到已污染的内容；
2. 目标格保留自己的 r 与 s（样式）等属性，只替换内容与 t 属性；源格的 cm/vm 等
   位置相关属性不搬运；
3. 源格是公式（<f>）时退化为 <is><t>…</t></is>（t="inlineStr"），文本取缓存值；
4. 源格是共享字符串（t="s"）时直接复用同一个 sharedStrings 索引，富文本
   （楷体姓名 + Times New Roman 学号两段 run）100% 保留；
5. 其他 zip 条目逐字节原样复制，[Content_Types].xml 不动。

可选 move_marks=True（组长/标记底色跟随学生）：
- 不改整格的 s（那会把边框/字体/对齐一起搬走，真实表里同排座位样式并不相同），
  而是解析 xl/styles.xml 的 cellXfs，**克隆目标格当前的 xf**、只把 fillId 换成源格
  的 fillId（补 applyFill="1"），追加到 cellXfs 并更新 count，目标格写入新 s；
- 目标格除填充色外的 xf 字段与它原来的 s 完全一致；缓存 (dst_xf, fillId) 复用；
- 只额外改写 xl/styles.xml 与目标 sheet XML，其余 zip 条目逐字节保留。

可选 text_updates=[{"to": "F1", "text": "座位周期：…"}]（字面量文本更新）：
- 在 moves 之后按顺序应用，同一 ref 后者覆盖；用于自动更新表头周期日期；
- 目标格保留原 s，内容写成 t="inlineStr"（XML 转义、xml:space="preserve"）；
- 合并非左上角目标重定向 anchor；目标格不存在时新建（复用插入逻辑）；
- 只额外改写目标 sheet XML（move_marks 需要时才会改写 styles.xml）。

对外接口：

    export_workbook(data: bytes, index: int, moves: list[dict], out_path: str,
                    move_marks: bool = False, text_updates: list = None) -> dict
        moves: [{"to": "C3", "from": "C9"}]，A1 引用，允许带 $（大小写不敏感）。
        返回 {"out_path": str, "applied": int, "marks_moved": int,
              "text_updates_applied": int, "skipped": [{"to","from"/"text","reason"}]}
        - applied: 实际生效的 move 数；
        - marks_moved: move_marks=True 时填充色确实发生变化的目标格数（=标记搬入数
          + 旧位置清除数；默认 False 时为 0）；
        - text_updates_applied: 实际写入的字面量文本目标格数；
        - skipped: 无效引用 / 目标落在合并区域非左上角（已自动改用 anchor）等提示。
"""

from __future__ import annotations

import copy
import os
import re
import tempfile
import xml.sax.saxutils as saxutils
import zipfile

from .xlsx_io import (
    _attr,
    _children,
    _covered_cells,
    _first_child,
    _load_shared_strings,
    _num_text,
    _open_package,
    _parse_merges,
    _parse_xml,
    _part_path,
    _ref_key,
    _workbook_sheets,
    make_a1,
    parse_a1,
)

__all__ = ["export_workbook"]

# <c ...> / </c>
_CELL_OPEN_RE = re.compile(r"<(?P<prefix>[A-Za-z_][\w.\-]*:)?c(?=[\s/>])")
_CELL_CLOSE_RE = re.compile(r"</(?:[A-Za-z_][\w.\-]*:)?c\s*>")
# <row ...> / </row>
_ROW_OPEN_RE = re.compile(r"<(?P<prefix>[A-Za-z_][\w.\-]*:)?row(?=[\s/>])")
_ROW_CLOSE_RE = re.compile(r"</(?:[A-Za-z_][\w.\-]*:)?row\s*>")
# <sheetData ...> / </sheetData>
_SHEETDATA_CLOSE_RE = re.compile(r"</(?:[A-Za-z_][\w.\-]*:)?sheetData\s*>")
_SHEETDATA_SELF_RE = re.compile(r"<((?:[A-Za-z_][\w.\-]*:)?)sheetData(?P<attrs>[^>]*)/>")
# 属性
_ATTR_RE = re.compile(r"([^\s=/>]+)\s*=\s*(?:\"([^\"]*)\"|'([^']*)')")
# 单元格内部片段
_V_RE = re.compile(r"<(?:[A-Za-z_][\w.\-]*:)?v(?:\s[^>]*)?>(.*?)</(?:[A-Za-z_][\w.\-]*:)?v>", re.S)
_T_RE = re.compile(r"<(?:[A-Za-z_][\w.\-]*:)?t(?:\s[^>]*)?>(.*?)</(?:[A-Za-z_][\w.\-]*:)?t>", re.S)
_F_RE = re.compile(r"<(?:[A-Za-z_][\w.\-]*:)?f(?=[\s/>])")
_RPH_RE = re.compile(r"<(?:[A-Za-z_][\w.\-]*:)?rPh\b.*?</(?:[A-Za-z_][\w.\-]*:)?rPh\s*>", re.S)
_DIMENSION_RE = re.compile(r"(<(?:[A-Za-z_][\w.\-]*:)?dimension\b[^>]*?\bref\s*=\s*\")([^\"]*)(\")")

# styles.xml（move_marks 用）：<cellXfs> 与其中的 <xf>
_CELLXFS_OPEN_RE = re.compile(
    r"<(?P<prefix>[A-Za-z_][\w.\-]*:)?cellXfs(?P<attrs>[^>]*?)(?P<self_closing>/?)>")
_CELLXFS_CLOSE_RE = re.compile(r"</(?:[A-Za-z_][\w.\-]*:)?cellXfs\s*>")
_XF_OPEN_RE = re.compile(r"<(?P<prefix>[A-Za-z_][\w.\-]*:)?xf(?=[\s/>])")
_XF_CLOSE_RE = re.compile(r"</(?:[A-Za-z_][\w.\-]*:)?xf\s*>")
_INT_ATTR_RE = re.compile(r"\bfillId\s*=\s*[\"'](\d+)[\"']")


# ---------------------------------------------------------------------------
# XML 片段工具
# ---------------------------------------------------------------------------

def _parse_attrs(head: str) -> dict:
    """解析一个开始标签里的属性（name -> value，name 去掉命名空间前缀）。"""
    result = {}
    for match in _ATTR_RE.finditer(head):
        name = match.group(1).rsplit(":", 1)[-1]
        value = match.group(2) if match.group(2) is not None else match.group(3)
        result.setdefault(name, value)
    return result


def _strip_attr(head: str, name: str) -> str:
    """从开始标签（含结尾 > 或 />）里删掉一个属性。"""
    pattern = re.compile(r"\s+" + re.escape(name) + r"\s*=\s*(?:\"[^\"]*\"|'[^']*')")
    return pattern.sub("", head, count=1)


def _head_body(head: str) -> str:
    """开始标签去掉结尾的 '>' 或 '/>'。"""
    body = head[:-1]
    if body.endswith("/"):
        body = body[:-1]
    return body


def _set_attr(head: str, name: str, value: str) -> str:
    """给开始标签（含结尾 '>' 或 '/>'）设置/替换一个属性，其余字节不动。"""
    body = head[:-1]
    self_closing = body.endswith("/")
    if self_closing:
        body = body[:-1]
    pattern = re.compile(r"\s+" + re.escape(name) + r"\s*=\s*(?:\"[^\"]*\"|'[^']*')")
    match = pattern.search(body)
    if match:
        body = body[:match.start()] + ' %s="%s"' % (name, value) + body[match.end():]
    else:
        body = body + ' %s="%s"' % (name, value)
    return body + ("/" if self_closing else "") + ">"


def _scan_cells(xml: str) -> dict:
    """扫描 sheet XML 里所有 <c>，返回 {key: cell_info}（key 为规范化的 A1）。"""
    cells = {}
    position = 0
    length = len(xml)
    while position < length:
        match = _CELL_OPEN_RE.search(xml, position)
        if not match:
            break
        start = match.start()
        gt = xml.find(">", match.end())
        if gt < 0:
            break
        head = xml[start:gt + 1]
        self_closing = head.rstrip().endswith("/>")
        if self_closing:
            end = gt + 1
            inner = ""
        else:
            close = _CELL_CLOSE_RE.search(xml, gt + 1)
            if not close:
                break
            end = close.end()
            inner = xml[gt + 1:close.start()]
        attrs = _parse_attrs(head)
        ref = attrs.get("r")
        if ref:
            key = _ref_key(ref)
            prefix = match.group("prefix") or ""
            cells[key] = {
                "key": key,
                "ref": ref,
                "t": attrs.get("t"),
                "inner": inner,
                "head": head,
                "prefix": prefix,
                "self_closing": self_closing,
                "start": start,
                "end": end,
            }
        position = end
    return cells


def _scan_rows(xml: str) -> dict:
    """扫描 sheet XML 里所有 <row>，返回 {row_number: row_info}。"""
    rows = {}
    position = 0
    length = len(xml)
    while position < length:
        match = _ROW_OPEN_RE.search(xml, position)
        if not match:
            break
        start = match.start()
        gt = xml.find(">", match.end())
        if gt < 0:
            break
        head = xml[start:gt + 1]
        self_closing = head.rstrip().endswith("/>")
        prefix = match.group("prefix") or ""
        if self_closing:
            end = gt + 1
            inner_start = inner_end = gt + 1
        else:
            close = _ROW_CLOSE_RE.search(xml, gt + 1)
            if not close:
                break
            end = close.end()
            inner_start = gt + 1
            inner_end = close.start()
        attrs = _parse_attrs(head)
        row_attr = attrs.get("r")
        if row_attr and row_attr.strip().isdigit():
            rows[int(row_attr)] = {
                "r": int(row_attr),
                "start": start,
                "end": end,
                "head": head,
                "prefix": prefix,
                "self_closing": self_closing,
                "inner_start": inner_start,
                "inner_end": inner_end,
            }
        position = end
    return rows


def _build_cell(prefix: str, head: str, t_attr, inner: str) -> str:
    """用目标格的原始开始标签（保留 r/s 等属性）+ 新内容重建 <c> 元素。"""
    body = _strip_attr(_head_body(head), "t")
    if t_attr:
        body += ' t="%s"' % t_attr
    tag = "%sc" % (prefix or "")
    if inner:
        return "%s>%s</%s>" % (body, inner, tag)
    return body + "/>"


def _build_new_cell(ref: str, t_attr, inner: str) -> str:
    """目标格原本不存在时新建一个 <c>（无 s，继承列默认样式）。"""
    head = '<c r="%s"' % ref
    if t_attr:
        head += ' t="%s"' % t_attr
    if inner:
        return head + ">" + inner + "</c>"
    return head + "/>"


def _inline_str(text: str) -> str:
    """生成 <is><t xml:space="preserve">text</t></is>。"""
    if not text:
        return ""
    return '<is><t xml:space="preserve">%s</t></is>' % saxutils.escape(text)


def _inner_value(inner: str) -> str:
    """取单元格内部 <v> 的文本（已反转义）；没有则空串。"""
    match = _V_RE.search(inner)
    if not match:
        return ""
    return saxutils.unescape(match.group(1))


def _inner_inline_text(inner: str) -> str:
    """取 <is> 里所有 <t> 文本（富文本拼接，忽略 <rPh> 注音）。"""
    cleaned = _RPH_RE.sub("", inner)
    return "".join(saxutils.unescape(part) for part in _T_RE.findall(cleaned))


def _source_plain_text(cell: dict, shared_strings: list) -> str:
    """源单元格的显示文本（用于公式退化 / t="str" 退化）。"""
    cell_type = cell.get("t")
    inner = cell.get("inner") or ""
    if cell_type == "inlineStr":
        return _inner_inline_text(inner)
    raw = _inner_value(inner)
    if cell_type == "s":
        try:
            index = int(raw.strip())
        except ValueError:
            return ""
        if 0 <= index < len(shared_strings):
            return shared_strings[index]
        return ""
    if cell_type == "b":
        return "TRUE" if raw.strip() in ("1", "true", "TRUE") else "FALSE"
    if cell_type in ("str", "e"):
        return raw
    return _num_text(raw)


def _content_for_source(cell, shared_strings: list):
    """计算目标格应写入的 (t_attr, inner)；cell 为 None 表示源格为空（清空目标）。"""
    if cell is None:
        return None, ""
    inner = cell.get("inner") or ""
    if not inner.strip():
        return None, ""
    cell_type = cell.get("t")
    if _F_RE.search(inner):  # 公式：退化为 inlineStr 文本（取缓存值）
        return "inlineStr", _inline_str(_source_plain_text(cell, shared_strings))
    if cell_type == "s":
        raw = _inner_value(inner).strip()
        try:
            index = int(raw)
        except ValueError:
            return None, ""
        if not (0 <= index < len(shared_strings)):
            return None, ""
        return "s", inner  # 复用 sharedStrings 索引 -> 富文本原样保留
    if cell_type == "inlineStr":
        return "inlineStr", inner
    if cell_type in ("str", "e"):
        return "inlineStr", _inline_str(_source_plain_text(cell, shared_strings))
    if cell_type == "b":
        return "b", inner
    if cell_type == "n":
        return "n", inner
    return None, inner


# ---------------------------------------------------------------------------
# 标记跟随（move_marks）：只搬运填充色，克隆「目标位置的 xf」后仅替换 fillId
# ---------------------------------------------------------------------------

def _parse_styles_cellxfs(styles_xml: str):
    """解析 styles.xml 的 <cellXfs>，返回定位信息 + 每个 <xf> 的原文/位置/fillId。"""
    match = _CELLXFS_OPEN_RE.search(styles_xml)
    if not match:
        return None
    prefix = match.group("prefix") or ""
    attrs_text = match.group("attrs") or ""
    self_closing = match.group("self_closing") == "/"
    info = {
        "prefix": prefix,
        "open_start": match.start(),
        "open_end": match.end(),
        "self_closing": self_closing,
        "close_start": match.end(),
        "xfs": [],
    }
    if self_closing:
        return info
    close = _CELLXFS_CLOSE_RE.search(styles_xml, match.end())
    info["close_start"] = close.start() if close else len(styles_xml)
    region_end = info["close_start"]
    position = match.end()
    while position < region_end:
        xf_match = _XF_OPEN_RE.search(styles_xml, position, region_end)
        if not xf_match:
            break
        gt = styles_xml.find(">", xf_match.end(), region_end)
        if gt < 0:
            break
        head = styles_xml[xf_match.start():gt + 1]
        xf_self_closing = head.rstrip().endswith("/>")
        if xf_self_closing:
            end = gt + 1
            inner = ""
        else:
            close_xf = _XF_CLOSE_RE.search(styles_xml, gt + 1, region_end)
            if not close_xf:
                break
            end = close_xf.end()
            inner = styles_xml[gt + 1:close_xf.start()]
        fill_match = _INT_ATTR_RE.search(head)
        info["xfs"].append({
            "start": xf_match.start(),
            "end": end,
            "prefix": xf_match.group("prefix") or "",
            "head": head,
            "self_closing": xf_self_closing,
            "inner": inner,
            "fillId": int(fill_match.group(1)) if fill_match else 0,
            "raw": styles_xml[xf_match.start():end],
        })
        position = end
    return info


def _clone_xf_with_fill(xf: dict, fill_id: int) -> str:
    """克隆一个 <xf> 原文，只替换 fillId 并补 applyFill="1"（其他属性/子元素原样）。"""
    head = _set_attr(xf["head"], "fillId", str(fill_id))
    head = _set_attr(head, "applyFill", "1")
    if xf["self_closing"]:
        return head
    return "%s%s</%sxf>" % (head, xf["inner"], xf["prefix"])


def _append_cellxfs(styles_xml: str, info: dict, new_xf_texts: list, new_count: int) -> str:
    """把新 <xf> 追加到 cellXfs 末尾并更新 count 属性；其余字节不动。"""
    open_start, open_end = info["open_start"], info["open_end"]
    open_tag = styles_xml[open_start:open_end]
    new_open = _set_attr(open_tag, "count", str(new_count))
    joined = "".join(new_xf_texts)
    if info["self_closing"]:
        # <cellXfs count="0"/> -> <cellXfs count="N">…</cellXfs>
        replacement = "%s>%s</%scellXfs>" % (new_open[:-2], joined, info["prefix"])
        return styles_xml[:open_start] + replacement + styles_xml[open_end:]
    close_start = info["close_start"]
    return (styles_xml[:open_start] + new_open
            + styles_xml[open_end:close_start] + joined + styles_xml[close_start:])


def _cell_xf_index(cell: dict, xf_count: int):
    """单元格的 s 属性 -> cellXfs 索引；无 s 视为 0；异常值返回 None（安全跳过）。"""
    if xf_count <= 0:
        return None
    attrs = _parse_attrs(cell["head"])
    raw = attrs.get("s")
    if raw is None:
        return 0
    try:
        index = int(raw)
    except (TypeError, ValueError):
        return None
    if index < 0 or index >= xf_count:
        return None
    return index


def _plan_mark_styles(archive, cells: dict, final: dict):
    """计算 move_marks 的样式改写方案。

    返回 (styles_path, new_styles_xml, target_new_s, marks_moved)：
    - styles_path / new_styles_xml 为 None 表示 styles.xml 不需要改写；
    - target_new_s: {目标格 key: 新 cellXfs 索引}；
    - marks_moved: 填充色确实被搬运的目标格数量。
    """
    styles_path = _part_path(archive, "styles", "xl/styles.xml")
    if styles_path is None:
        return None, None, {}, 0
    try:
        styles_xml = archive.read(styles_path).decode("utf-8")
    except UnicodeDecodeError:
        styles_xml = archive.read(styles_path).decode("utf-8", "surrogateescape")
    info = _parse_styles_cellxfs(styles_xml)
    if info is None:
        return None, None, {}, 0
    xfs = info["xfs"]
    if not xfs:
        return None, None, {}, 0

    base_index = len(xfs)          # 实际 <xf> 元素个数（= 新索引起点）
    cache = {}                     # (dst_xf, src_fill) -> new_xf_index
    new_xf_texts = []
    target_new_s = {}
    marks_moved = 0
    for to_key in final:
        source = cells.get(final[to_key])
        target = cells.get(to_key)
        if source is None or target is None:
            continue  # 无对应 <c> 的 move 安全跳过
        src_xf = _cell_xf_index(source, base_index)
        dst_xf = _cell_xf_index(target, base_index)
        if src_xf is None or dst_xf is None:
            continue
        src_fill = xfs[src_xf]["fillId"]
        dst_fill = xfs[dst_xf]["fillId"]
        if src_fill == dst_fill:
            continue  # 填充色相同，保持目标格原来的 s 不变
        cache_key = (dst_xf, src_fill)
        new_index = cache.get(cache_key)
        if new_index is None:
            new_index = base_index + len(new_xf_texts)
            new_xf_texts.append(_clone_xf_with_fill(xfs[dst_xf], src_fill))
            cache[cache_key] = new_index
        target_new_s[to_key] = new_index
        marks_moved += 1

    if not new_xf_texts:
        return None, None, {}, 0
    new_count = base_index + len(new_xf_texts)
    new_styles_xml = _append_cellxfs(styles_xml, info, new_xf_texts, new_count)
    return styles_path, new_styles_xml, target_new_s, marks_moved


# ---------------------------------------------------------------------------
# 插入目标格（目标格原本不存在时）
# ---------------------------------------------------------------------------

def _insert_cells(xml: str, insertions: list) -> str:
    """把 (row, col, cell_xml) 插入到对应 <row>；按列序插入，尽量保持 XML 合法。"""
    if not insertions:
        return xml
    rows = _scan_rows(xml)
    cells = _scan_cells(xml)
    by_row = {}
    for row_index, col_index, cell_xml in insertions:
        by_row.setdefault(row_index, []).append((col_index, cell_xml))

    row_positions = []  # (pos, sort_key, text)
    missing_rows = []
    for row_index, items in by_row.items():
        row = rows.get(row_index)
        if row is None:
            missing_rows.append(row_index)
            continue
        items.sort(key=lambda item: item[0], reverse=True)  # 同一行先插后面的列
        if row["self_closing"]:
            # <row .../> -> <row ...>cells</row>
            body = _head_body(row["head"])
            inner_xml = "".join(cell for _col, cell in items)
            tag = "%srow" % (row["prefix"] or "")
            text = "%s>%s</%s>" % (body, inner_xml, tag)
            row_positions.append((row["start"], 0, text))
            continue
        row_cells = [
            (cell["start"], cell["key"])
            for cell in cells.values()
            if row["inner_start"] <= cell["start"] < row["inner_end"]
        ]
        row_cells.sort(key=lambda item: item[0])
        for col_index, cell_xml in items:
            position = row["inner_end"]
            for cell_start, _key in row_cells:
                try:
                    _r, cell_col = parse_a1(_key)
                except ValueError:
                    continue
                if cell_col > col_index:
                    position = cell_start
                    break
            row_positions.append((position, col_index, cell_xml))

    if row_positions:
        row_positions.sort(key=lambda item: (item[0], item[1]), reverse=True)
        for position, _order, text in row_positions:
            xml = xml[:position] + text + xml[position:]

    if missing_rows:
        xml = _insert_new_rows(xml, missing_rows, by_row)
    return xml


def _insert_new_rows(xml: str, row_numbers: list, by_row: dict) -> str:
    """为原本不存在的行插入 <row r="N"><c/></row>（按行号顺序）。"""
    # 每次插入后位置会变化，行数通常很少，这里逐行重扫，避免复杂索引维护。
    for row_index in sorted(row_numbers, reverse=True):
        cells = "".join(cell for _col, cell in sorted(by_row[row_index], key=lambda item: item[0]))
        row_xml = '<row r="%d">%s</row>' % (row_index, cells)
        rows = _scan_rows(xml)
        sheet_data = _first_child(_parse_xml(xml.encode("utf-8", "surrogateescape")), "sheetData")
        insert_at = None
        if rows:
            for existing_row in sorted(rows.values(), key=lambda item: item["r"]):
                if existing_row["r"] > row_index:
                    insert_at = existing_row["start"]
                    break
        if insert_at is None:
            close = _SHEETDATA_CLOSE_RE.search(xml)
            if close:
                insert_at = close.start()
            else:
                self_closing = _SHEETDATA_SELF_RE.search(xml)
                if self_closing:
                    prefix = self_closing.group(1) or ""
                    replacement = "<%ssheetData%s>%s</%ssheetData>" % (
                        prefix, self_closing.group("attrs"), row_xml, prefix)
                    return xml[:self_closing.start()] + replacement + xml[self_closing.end():]
                return xml  # 找不到插入点：放弃（sheetData 结构异常）
        xml = xml[:insert_at] + row_xml + xml[insert_at:]
    return xml


# ---------------------------------------------------------------------------
# zip 写出
# ---------------------------------------------------------------------------

def _write_archive(archive: zipfile.ZipFile, out_path: str, replacements: dict) -> None:
    """按原顺序写出 xlsx；replacements: {zip 内路径: 新 bytes}，其余条目逐字节复制。"""
    directory = os.path.dirname(os.path.abspath(out_path))
    if directory:
        os.makedirs(directory, exist_ok=True)
    handle, temp_path = tempfile.mkstemp(prefix=".seat_export_", suffix=".xlsx", dir=directory or None)
    os.close(handle)
    try:
        with zipfile.ZipFile(temp_path, "w") as out:
            for info in archive.infolist():
                payload = replacements.get(info.filename)
                if payload is None:
                    payload = archive.read(info)
                out.writestr(copy.copy(info), payload)
        os.replace(temp_path, out_path)
    except BaseException:
        try:
            os.unlink(temp_path)
        except OSError:
            pass
        raise


# ---------------------------------------------------------------------------
# 对外接口
# ---------------------------------------------------------------------------

def export_workbook(data: bytes, index: int, moves: list, out_path: str,
                    move_marks: bool = False, text_updates: list = None) -> dict:
    """把 moves（A1 引用）应用到第 index 个 sheet 并写出新 xlsx。

    move_marks=True 时，除文字内容外还让「填充色」跟随学生走（如组长橙色底）：
    克隆**目标位置**的 xf、只把 fillId 换成源格的 fillId，追加到 cellXfs 并改写
    目标格 s —— 边框/字体/对齐/数字格式等位置样式保持目标格原样，不会被搬走。

    text_updates=[{"to": "F1", "text": "座位周期：…"}] 用于字面量文本更新
    （如自动更新表头周期日期），在 moves 之后按顺序应用：同一 ref 后者覆盖；
    目标格保留原 s，内容写成 t="inlineStr"（XML 转义 + xml:space="preserve"）；
    合并区域非左上角的目标会重定向到 anchor；目标格不存在时新建。

    返回 {"out_path": str, "applied": int, "marks_moved": int,
          "text_updates_applied": int, "skipped": [{"to","from"/"text","reason"}]}。
    """
    archive = _open_package(data)
    sheets = _workbook_sheets(archive)
    try:
        sheet_index = int(index)
    except (TypeError, ValueError):
        raise ValueError("工作表索引无效：%r" % (index,)) from None
    if sheet_index < 0 or sheet_index >= len(sheets):
        raise ValueError("工作表索引超出范围：%d" % sheet_index)

    sheet_path = sheets[sheet_index]["path"]
    sheet_bytes = archive.read(sheet_path)
    try:
        xml = sheet_bytes.decode("utf-8")
    except UnicodeDecodeError:
        xml = sheet_bytes.decode("utf-8", "surrogateescape")
    root = _parse_xml(sheet_bytes)
    if root is None:
        raise ValueError("不是有效的 .xlsx 文件")

    merges = _parse_merges(root)
    covered = _covered_cells(merges)
    merge_refs = {(merge["r"], merge["c"]): merge["ref"] for merge in merges}
    shared_strings = _load_shared_strings(archive)

    # ---- 第一阶段：规范化 move，解析合并区域重定向 -------------------------
    parsed = []   # [(to_key, from_key)]
    skipped = []
    for move in moves or []:
        if not isinstance(move, dict):
            skipped.append({"to": None, "from": None, "reason": "move 必须是 {to, from} 对象"})
            continue
        to_raw = move.get("to")
        from_raw = move.get("from")
        try:
            to_row, to_col = parse_a1(to_raw)
            from_row, from_col = parse_a1(from_raw)
        except (ValueError, TypeError):
            skipped.append({"to": to_raw, "from": from_raw, "reason": "无效的单元格引用"})
            continue
        anchor = covered.get((to_row, to_col))
        if anchor is not None:
            merge_ref = merge_refs.get(anchor, make_a1(*anchor))
            adjusted = make_a1(*anchor)
            skipped.append({
                "to": to_raw,
                "from": from_raw,
                "reason": "目标位于合并区域 %s 内，已改用左上角 %s" % (merge_ref, adjusted),
            })
            to_row, to_col = anchor
        parsed.append((make_a1(to_row, to_col), make_a1(from_row, from_col)))

    # ---- 第一阶段 b：解析字面量文本更新（在 moves 之后写，覆盖同格文字）----
    text_map = {}  # 目标 key -> 要写入的纯文本
    if text_updates:
        for item in text_updates:
            if not isinstance(item, dict):
                skipped.append({"to": None, "reason": "text_update 必须是 {to, text} 对象"})
                continue
            to_raw = item.get("to")
            text_value = item.get("text")
            if not isinstance(text_value, str):
                skipped.append({"to": to_raw, "text": text_value,
                                "reason": "text_update 缺少纯文本 text"})
                continue
            try:
                to_row, to_col = parse_a1(to_raw)
            except (ValueError, TypeError):
                skipped.append({"to": to_raw, "text": text_value,
                                "reason": "无效的单元格引用"})
                continue
            anchor = covered.get((to_row, to_col))
            if anchor is not None:
                merge_ref = merge_refs.get(anchor, make_a1(*anchor))
                adjusted = make_a1(*anchor)
                skipped.append({
                    "to": to_raw,
                    "text": text_value,
                    "reason": "目标位于合并区域 %s 内，已改用左上角 %s" % (merge_ref, adjusted),
                })
                to_row, to_col = anchor
            text_map[make_a1(to_row, to_col)] = text_value  # 同一 ref 后者覆盖
    text_updates_applied = len(text_map)

    # ---- 恒等映射：内容与填充色都不可能变化，直接逐字节复制 ---------------
    if not text_map and all(to_key == from_key for to_key, from_key in parsed):
        _write_archive(archive, os.fspath(out_path), {})
        return {"out_path": os.fspath(out_path), "applied": len(parsed),
                "marks_moved": 0, "text_updates_applied": 0, "skipped": skipped}

    # ---- 第二阶段：一次性快照所有源格，再统一生成目标内容 -------------------
    cells = _scan_cells(xml)
    final = {}  # 目标 -> 源（后出现的 move 覆盖先出现的，符合“顺序执行”直觉）
    for to_key, from_key in parsed:
        final[to_key] = from_key

    # 标记跟随：只改写 styles.xml 的 cellXfs（追加克隆 xf）和目标格的 s 属性
    styles_path = None
    styles_payload = None
    target_new_s = {}
    marks_moved = 0
    if move_marks:
        styles_path, new_styles_xml, target_new_s, marks_moved = _plan_mark_styles(
            archive, cells, final)
        if new_styles_xml is not None:
            styles_payload = new_styles_xml.encode("utf-8", "surrogateescape")

    replacements = []
    insertions = []
    targets = list(final.keys())
    for key in text_map:
        if key not in final:
            targets.append(key)  # 纯文本更新的目标（不是任何 move 的目标）
    for to_key in targets:
        target = cells.get(to_key)
        if to_key in text_map:
            # 字面量文本更新：以 text_updates 为准，覆盖 move 搬来的文字
            inner = _inline_str(text_map[to_key])
            if not inner:
                inner = '<is><t xml:space="preserve"></t></is>'  # 空文本也保留 inlineStr
            t_attr = "inlineStr"
        else:
            source = cells.get(final[to_key])
            t_attr, inner = _content_for_source(source, shared_strings)
        if target is None:
            if t_attr is None and not inner:
                continue  # 目标本来就不存在、源又为空：无需写入
            to_row, to_col = parse_a1(to_key)
            insertions.append((to_row, to_col, _build_new_cell(to_key, t_attr, inner)))
            continue
        target_head = target["head"]
        new_style = target_new_s.get(to_key)
        if new_style is not None:
            target_head = _set_attr(target_head, "s", str(new_style))
        replacements.append((target["start"], target["end"],
                             _build_cell(target["prefix"], target_head, t_attr, inner)))

    # ---- 第三阶段：从后往前替换，索引不互相污染 -----------------------------
    for start, end, new_text in sorted(replacements, key=lambda item: item[0], reverse=True):
        xml = xml[:start] + new_text + xml[end:]

    if insertions:
        xml = _insert_cells(xml, insertions)
        xml = _fix_dimension(xml)

    replacements_map = {sheet_path: xml.encode("utf-8", "surrogateescape")}
    if styles_payload is not None and styles_path is not None:
        replacements_map[styles_path] = styles_payload
    _write_archive(archive, os.fspath(out_path), replacements_map)
    return {"out_path": os.fspath(out_path), "applied": len(parsed),
            "marks_moved": marks_moved, "text_updates_applied": text_updates_applied,
            "skipped": skipped}


def _fix_dimension(xml: str) -> str:
    """插入新格后，必要时扩大 <dimension ref="A1:O14"> 的 used range。"""
    match = _DIMENSION_RE.search(xml)
    if not match:
        return xml
    old_ref = match.group(2)
    pieces = old_ref.split(":")
    try:
        start_row, start_col = parse_a1(pieces[0])
        end_row, end_col = parse_a1(pieces[1]) if len(pieces) > 1 else (start_row, start_col)
    except ValueError:
        return xml
    max_row = end_row
    max_col = end_col
    root = _parse_xml(xml.encode("utf-8", "surrogateescape"))
    if root is not None:
        sheet_data = _first_child(root, "sheetData")
        if sheet_data is not None:
            for row_el in _children(sheet_data, "row"):
                for cell_el in _children(row_el, "c"):
                    ref = _attr(cell_el, "r")
                    if not ref:
                        continue
                    try:
                        row, col = parse_a1(ref)
                    except ValueError:
                        continue
                    max_row = max(max_row, row)
                    max_col = max(max_col, col)
    new_end = make_a1(max_row, max_col)
    if new_end == make_a1(end_row, end_col):
        return xml
    new_ref = "%s:%s" % (make_a1(start_row, start_col), new_end)
    return xml[:match.start(2)] + new_ref + xml[match.end(2):]
