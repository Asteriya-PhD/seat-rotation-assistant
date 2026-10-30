"""xlsx 读取引擎（纯 Python 标准库，零第三方依赖）。

对外接口（供 app/server.py 直接调用）：

    list_sheets(data: bytes) -> list[dict]
        每项 {"index": int, "name": str, "hidden": bool, "rows": int, "cols": int}
        （rows/cols 是附带便利字段；room 见 ARCHITECTURE §3.1 的 sheets 列表）

    read_sheet(data: bytes, index: int) -> dict
        返回 ARCHITECTURE §3.1 的 sheet 对象（不含 workbook_id，由 server 补）：
        index/name/rows/cols/merges/cells/row_heights/col_widths/
        default_col_width/default_row_height/hidden_rows/hidden_cols/orientation

    sheet_xml_path(data: bytes, index: int) -> str
        返回该 sheet 在 zip 内的路径，如 "xl/worksheets/sheet1.xml"。

实现要点：
- 不 import openpyxl / lxml；只用 zipfile + xml.etree.ElementTree + re。
- 命名空间一律按 localname 匹配，兼容 x: 前缀与未知命名空间。
- 颜色统一输出**无 # 的 "RRGGBB"**（与 ARCHITECTURE §3.1、mock.json 一致）；
  rgb 值若为 8 位 ARGB 则剥掉 alpha；theme 颜色从 xl/theme/theme1.xml 近似解析。
- 单元格 r/c 从 1 开始；合并区域只在左上角（anchor）输出 cell。
- 数字按 "7" 输出（不是 "7.0"）；富文本 run 直接拼接（<rPh> 注音不拼）。
- 本模块没有 import 期副作用，不读任何文件。
"""

from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET
import zipfile

__all__ = ["list_sheets", "read_sheet", "sheet_xml_path"]

_FALLBACK_COL_WIDTH = 8.43
_FALLBACK_ROW_HEIGHT = 15.0

# 标准 indexed 调色板（ECMA-376 默认色，索引 0..63；64 表示系统默认，不解析）
_INDEXED_COLORS = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
]

_THEME_ORDER = [
    "lt1", "dk1", "lt2", "dk2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
]

_A1_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")
_HEX6_RE = re.compile(r"^[0-9A-Fa-f]{6}$")


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def _localname(tag: str) -> str:
    """去掉命名空间，返回 localname（tag 可能是 "{ns}name"）。"""
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def _attr(el, name: str):
    """按 localname 取属性；找不到返回 None。"""
    if el is None:
        return None
    for key, value in el.attrib.items():
        if _localname(key) == name:
            return value
    return None


def _first_child(el, name: str):
    if el is None:
        return None
    for child in el:
        if _localname(child.tag) == name:
            return child
    return None


def _children(el, name: str):
    if el is None:
        return []
    return [child for child in el if _localname(child.tag) == name]


def _parse_xml(data: bytes):
    """宽容地解析 XML；失败返回 None。"""
    if data is None:
        return None
    if isinstance(data, str):
        data = data.encode("utf-8", "surrogateescape")
    try:
        return ET.fromstring(data)
    except ET.ParseError:
        pass
    try:
        return ET.fromstring(data.lstrip(b"\xef\xbb\xbf").lstrip())
    except ET.ParseError:
        return None


def col_to_index(letters: str) -> int:
    """'A' -> 1, 'C' -> 3, 'AB' -> 28。"""
    n = 0
    for ch in letters.upper():
        if not ("A" <= ch <= "Z"):
            raise ValueError("无效的列名：%r" % letters)
        n = n * 26 + (ord(ch) - 64)
    if n <= 0:
        raise ValueError("无效的列名：%r" % letters)
    return n


def index_to_col(index: int) -> str:
    """3 -> 'C'。"""
    if index < 1:
        raise ValueError("列号必须从 1 开始")
    out = []
    n = index
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out.append(chr(65 + rem))
    return "".join(reversed(out))


def parse_a1(ref: str):
    """'$C$3' / 'c3' -> (3, 3)。无效时抛 ValueError。"""
    if not isinstance(ref, str):
        raise ValueError("无效的单元格引用：%r" % (ref,))
    m = _A1_RE.match(ref.strip())
    if not m:
        raise ValueError("无效的单元格引用：%r" % (ref,))
    col = col_to_index(m.group(1))
    row = int(m.group(2))
    if row < 1:
        raise ValueError("无效的单元格引用：%r" % (ref,))
    return row, col


def make_a1(row: int, col: int) -> str:
    return "%s%d" % (index_to_col(col), row)


def _ref_key(ref: str) -> str:
    """把 A1 引用规范化为大写无 $ 形式（无效则原样返回）。"""
    try:
        r, c = parse_a1(ref)
    except (ValueError, TypeError):
        return str(ref)
    return make_a1(r, c)


def _num_text(raw: str) -> str:
    """数字缓存值转显示文本：'7' -> '7'（不是 '7.0'）。"""
    if raw is None:
        return ""
    s = raw.strip()
    if not s:
        return ""
    try:
        value = float(s)
    except ValueError:
        return s
    if value != value or value in (float("inf"), float("-inf")):
        return s
    if value.is_integer() and abs(value) < 10 ** 15:
        return str(int(value))
    return s


# ---------------------------------------------------------------------------
# zip 包 / 关系
# ---------------------------------------------------------------------------

def _open_package(data):
    """打开 xlsx（zip）；不是有效 xlsx 时抛 ValueError("不是有效的 .xlsx 文件")。"""
    if isinstance(data, (bytearray, memoryview)):
        data = bytes(data)
    if not isinstance(data, bytes) or not data:
        raise ValueError("不是有效的 .xlsx 文件")
    bio = io.BytesIO(data)
    if not zipfile.is_zipfile(bio):
        raise ValueError("不是有效的 .xlsx 文件")
    try:
        archive = zipfile.ZipFile(bio)
    except (zipfile.BadZipFile, OSError):
        raise ValueError("不是有效的 .xlsx 文件") from None
    names = set(archive.namelist())
    if "xl/workbook.xml" not in names:
        raise ValueError("不是有效的 .xlsx 文件")
    return archive


def _resolve_target(base_dir: str, target: str) -> str:
    """把关系文件里的 Target 解析成 zip 内路径（不使用 os.path，保持 POSIX 语义）。"""
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base_dir + "/" + target).split("/")
    stack = []
    for part in parts:
        if part in ("", "."):
            continue
        if part == "..":
            if stack:
                stack.pop()
            continue
        stack.append(part)
    return "/".join(stack)


def _rels_for_part(z, part: str):
    """读取 part 的 _rels 文件（part 如 'xl/workbook.xml'），返回 {Id: (Type, Target)}。"""
    directory, filename = part.rsplit("/", 1) if "/" in part else ("", part)
    rels_path = "%s/_rels/%s.rels" % (directory, filename) if directory else "_rels/%s.rels" % filename
    names = set(z.namelist())
    if rels_path not in names:
        return {}
    root = _parse_xml(z.read(rels_path))
    if root is None:
        return {}
    rels = {}
    for rel in _children(root, "Relationship"):
        rid = _attr(rel, "Id")
        target = _attr(rel, "Target")
        rel_type = _attr(rel, "Type") or ""
        if not rid or not target:
            continue
        if (_attr(rel, "TargetMode") or "").lower() == "external":
            continue
        rels[rid] = (rel_type, target)
    return rels


def _part_path(z, type_suffix: str, default: str):
    """按关系类型后缀（如 'styles'）找 zip 内路径，找不到用默认路径（存在才返回）。"""
    rels = _rels_for_part(z, "xl/workbook.xml")
    names = set(z.namelist())
    for _rid, (rel_type, target) in rels.items():
        if rel_type.endswith("/" + type_suffix):
            path = _resolve_target("xl", target)
            if path in names:
                return path
    if default in names:
        return default
    return None


def _rel_id(el) -> str:
    """取 <sheet r:id="rId1"/> 里的 r:id（也可能是 id）。"""
    for key, value in el.attrib.items():
        local = _localname(key)
        if local == "id" and "}" in key:  # 带命名空间的 r:id
            return value
    return _attr(el, "id")


def _workbook_sheets(z) -> list:
    """解析 workbook.xml + rels，返回 [{index,name,hidden,path}]。"""
    root = _parse_xml(z.read("xl/workbook.xml"))
    if root is None:
        raise ValueError("不是有效的 .xlsx 文件")
    rels = _rels_for_part(z, "xl/workbook.xml")
    names = set(z.namelist())
    sheets = []
    candidates = sorted(n for n in names if n.startswith("xl/worksheets/") and n.endswith(".xml"))
    sheets_el = _first_child(root, "sheets")
    sheet_elements = _children(sheets_el, "sheet") if sheets_el is not None else _children(root, "sheet")
    for index, sheet_el in enumerate(sheet_elements):
        name = _attr(sheet_el, "name") or ("Sheet%d" % (index + 1))
        state = (_attr(sheet_el, "state") or "").lower()
        hidden = state in ("hidden", "veryhidden")
        rid = _rel_id(sheet_el)
        path = None
        if rid and rid in rels:
            candidate = _resolve_target("xl", rels[rid][1])
            if candidate in names:
                path = candidate
        if path is None:
            guess = "xl/worksheets/sheet%d.xml" % (index + 1)
            if guess in names:
                path = guess
            elif index < len(candidates):
                path = candidates[index]
        if path is None:
            raise ValueError("不是有效的 .xlsx 文件")
        sheets.append({"index": index, "name": name, "hidden": hidden, "path": path})
    if not sheets:
        raise ValueError("不是有效的 .xlsx 文件")
    return sheets


# ---------------------------------------------------------------------------
# sharedStrings / styles / theme
# ---------------------------------------------------------------------------

def _si_text(si) -> str:
    """拼一个 <si> / <is> 的可见文本（富文本各 run 拼接，<rPh> 注音不拼）。"""
    parts = []
    for child in si:
        local = _localname(child.tag)
        if local == "t":
            parts.append(child.text or "")
        elif local == "r":
            for sub in child:
                if _localname(sub.tag) == "t":
                    parts.append(sub.text or "")
    return "".join(parts)


def _load_shared_strings(z) -> list:
    path = _part_path(z, "sharedStrings", "xl/sharedStrings.xml")
    if path is None:
        return []
    root = _parse_xml(z.read(path))
    if root is None:
        return []
    return [_si_text(si) for si in _children(root, "si")]


def _color_rgb(el, theme):
    """颜色元素 -> 无 # 的 'RRGGBB'；theme/indexed 做近似解析；无法解析返回 None。"""
    if el is None:
        return None
    rgb = _attr(el, "rgb")
    if rgb:
        value = rgb.strip().upper()
        if len(value) == 8:
            value = value[2:]  # 剥掉 alpha
        if _HEX6_RE.match(value):
            return value
        return None
    tint = _attr(el, "tint")
    theme_index = _attr(el, "theme")
    if theme_index is not None:
        try:
            base = theme[int(theme_index)]
        except (IndexError, ValueError, TypeError):
            return None
        if not base:
            return None
        return _apply_tint(base, tint)
    indexed = _attr(el, "indexed")
    if indexed is not None:
        try:
            idx = int(indexed)
        except ValueError:
            return None
        if 0 <= idx < len(_INDEXED_COLORS):
            return _INDEXED_COLORS[idx]
        return None
    # auto="1" 等：交给前端默认色
    return None


def _apply_tint(hex6: str, tint) -> str:
    if not tint:
        return hex6
    try:
        ratio = float(tint)
    except (TypeError, ValueError):
        return hex6
    if ratio == 0:
        return hex6
    channels = []
    for i in (0, 2, 4):
        channel = int(hex6[i:i + 2], 16)
        value = channel * (1 + ratio) if ratio < 0 else channel + (255 - channel) * ratio
        channels.append(max(0, min(255, int(round(value)))))
    return "%02X%02X%02X" % tuple(channels)


def _load_theme(z) -> list:
    """返回 theme 索引 -> RGB 的列表（0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..=accent1..）。"""
    path = _part_path(z, "theme", "xl/theme/theme1.xml")
    if path is None:
        return []
    root = _parse_xml(z.read(path))
    if root is None:
        return []
    clr_scheme = None
    for el in root.iter():
        if _localname(el.tag) == "clrScheme":
            clr_scheme = el
            break
    if clr_scheme is None:
        return []
    found = {}
    for child in clr_scheme:
        local = _localname(child.tag)
        rgb = None
        for sub in child:
            sub_local = _localname(sub.tag)
            if sub_local == "srgbClr":
                value = (_attr(sub, "val") or "").strip()
                if _HEX6_RE.match(value):
                    rgb = value.upper()
            elif sub_local == "sysClr":
                value = _attr(sub, "lastClr") or _attr(sub, "val") or ""
                value = value.strip()
                if _HEX6_RE.match(value):
                    rgb = value.upper()
        found[local] = rgb
    return [found.get(name) for name in _THEME_ORDER]


def _truthy_attr(el, default=True) -> bool:
    value = _attr(el, "val")
    if value is None:
        return default
    return value.strip().lower() not in ("0", "false", "off", "none")


def _parse_font(el, theme) -> dict:
    font = {"bold": False, "italic": False, "size": 11.0, "name": "宋体", "color": None}
    if el is None:
        return font
    for child in el:
        local = _localname(child.tag)
        if local == "b":
            font["bold"] = _truthy_attr(child)
        elif local == "i":
            font["italic"] = _truthy_attr(child)
        elif local == "sz":
            try:
                font["size"] = float(_attr(child, "val"))
            except (TypeError, ValueError):
                pass
        elif local in ("name", "rFont"):
            value = _attr(child, "val")
            if value:
                font["name"] = value
        elif local == "color":
            font["color"] = _color_rgb(child, theme)
    return font


def _parse_fill(el, theme):
    if el is None:
        return None
    pattern = _first_child(el, "patternFill")
    if pattern is None:
        return None
    pattern_type = (_attr(pattern, "patternType") or "none").lower()
    if pattern_type in ("none", "gray125"):
        return None
    fg = _color_rgb(_first_child(pattern, "fgColor"), theme)
    bg = _color_rgb(_first_child(pattern, "bgColor"), theme)
    return fg or bg


def _parse_border(el) -> bool:
    if el is None:
        return False
    draw_diagonal = (_attr(el, "diagonal") or "").strip().lower() in ("1", "true")
    for side in ("left", "right", "top", "bottom", "diagonal", "vertical", "horizontal"):
        side_el = _first_child(el, side)
        style = _attr(side_el, "style")
        if not style or style.lower() in ("none", "nil"):
            continue
        if side == "diagonal" and not draw_diagonal:
            continue  # diagonal="0"/缺省时对角边框不画
        return True
    return False


def _parse_alignment(xf) -> dict:
    alignment = _first_child(xf, "alignment")
    result = {"align": None, "valign": None, "wrap": False}
    if alignment is None:
        return result
    horizontal = _attr(alignment, "horizontal")
    vertical = _attr(alignment, "vertical")
    if horizontal and horizontal.lower() not in ("general",):
        result["align"] = horizontal
    if vertical and vertical.lower() not in ("general",):
        result["valign"] = vertical
    wrap = _attr(alignment, "wrapText")
    if wrap is not None:
        result["wrap"] = wrap.strip().lower() in ("1", "true", "on")
    return result


class _StyleTable:
    """styles.xml 的精简解析结果。"""

    __slots__ = ("fonts", "fills", "borders", "xfs", "default_font")

    def __init__(self):
        self.fonts = []
        self.fills = []
        self.borders = []
        self.xfs = []
        self.default_font = {"bold": False, "italic": False, "size": 11.0,
                             "name": "宋体", "color": None}


def _load_styles(z) -> _StyleTable:
    table = _StyleTable()
    path = _part_path(z, "styles", "xl/styles.xml")
    if path is None:
        return table
    root = _parse_xml(z.read(path))
    if root is None:
        return table
    theme = _load_theme(z)

    fonts_el = _first_child(root, "fonts")
    if fonts_el is not None:
        table.fonts = [_parse_font(font, theme) for font in _children(fonts_el, "font")]
    fills_el = _first_child(root, "fills")
    if fills_el is not None:
        table.fills = [_parse_fill(fill, theme) for fill in _children(fills_el, "fill")]
    borders_el = _first_child(root, "borders")
    if borders_el is not None:
        table.borders = [_parse_border(border) for border in _children(borders_el, "border")]
    xfs_el = _first_child(root, "cellXfs")
    if xfs_el is not None:
        for xf in _children(xfs_el, "xf"):
            try:
                font_id = int(_attr(xf, "fontId") or 0)
            except ValueError:
                font_id = 0
            try:
                fill_id = int(_attr(xf, "fillId") or 0)
            except ValueError:
                fill_id = 0
            try:
                border_id = int(_attr(xf, "borderId") or 0)
            except ValueError:
                border_id = 0
            alignment = _parse_alignment(xf)
            table.xfs.append({
                "fontId": font_id,
                "fillId": fill_id,
                "borderId": border_id,
                "align": alignment["align"],
                "valign": alignment["valign"],
                "wrap": alignment["wrap"],
            })
    if table.fonts:
        table.default_font = table.fonts[0]
    return table


def _style_for(table: _StyleTable, style_index) -> dict:
    """cellXfs 索引 -> §3.1 的精简 style dict。"""
    xf = None
    try:
        idx = int(style_index)
    except (TypeError, ValueError):
        idx = 0
    if 0 <= idx < len(table.xfs):
        xf = table.xfs[idx]
    font = table.default_font
    fill = None
    border = False
    align = valign = None
    wrap = False
    if xf is not None:
        if 0 <= xf["fontId"] < len(table.fonts):
            font = table.fonts[xf["fontId"]]
        if 0 <= xf["fillId"] < len(table.fills):
            fill = table.fills[xf["fillId"]]
        if 0 <= xf["borderId"] < len(table.borders):
            border = table.borders[xf["borderId"]]
        align, valign, wrap = xf["align"], xf["valign"], xf["wrap"]

    style = {}
    if font.get("bold"):
        style["bold"] = True
    if font.get("italic"):
        style["italic"] = True
    style["size"] = float(font.get("size", 11.0))
    style["font"] = font.get("name") or "宋体"
    if font.get("color"):
        style["color"] = font["color"]
    if fill:
        style["fill"] = fill
    if align:
        style["align"] = align
    if valign:
        style["valign"] = valign
    if wrap:
        style["wrap"] = True
    if border:
        style["border"] = True
    return style


# ---------------------------------------------------------------------------
# 单元格
# ---------------------------------------------------------------------------

def _cell_element_text(cell_el, shared_strings) -> str:
    """<c> 元素 -> 显示文本（富文本拼接 / 数字归一 / 公式取缓存值）。"""
    cell_type = _attr(cell_el, "t")
    if cell_type == "inlineStr":
        inline = _first_child(cell_el, "is")
        return _si_text(inline) if inline is not None else ""
    value_el = _first_child(cell_el, "v")
    raw = value_el.text if (value_el is not None and value_el.text is not None) else ""
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


def _sheet_extent(root, cells=None):
    """估算 used range 的 (max_row, max_col)。"""
    max_row = 0
    max_col = 0
    dimension = _first_child(root, "dimension")
    if dimension is not None:
        ref = _attr(dimension, "ref") or ""
        parts = ref.split(":")
        for part in parts:
            try:
                r, c = parse_a1(part)
            except ValueError:
                continue
            max_row = max(max_row, r)
            max_col = max(max_col, c)
    sheet_data = _first_child(root, "sheetData")
    if sheet_data is not None:
        for row_el in _children(sheet_data, "row"):
            row_number = _attr(row_el, "r")
            if row_number:
                try:
                    max_row = max(max_row, int(row_number))
                except ValueError:
                    pass
            for cell_el in _children(row_el, "c"):
                try:
                    r, c = parse_a1(_attr(cell_el, "r") or "")
                except ValueError:
                    continue
                max_row = max(max_row, r)
                max_col = max(max_col, c)
    merge_cells = _first_child(root, "mergeCells")
    if merge_cells is not None:
        for merge_el in _children(merge_cells, "mergeCell"):
            ref = _attr(merge_el, "ref") or ""
            for part in ref.split(":"):
                try:
                    r, c = parse_a1(part)
                except ValueError:
                    continue
                max_row = max(max_row, r)
                max_col = max(max_col, c)
    if cells:
        for cell in cells:
            max_row = max(max_row, cell["r"])
            max_col = max(max_col, cell["c"])
    return max(1, max_row), max(1, max_col)


def _parse_merges(root) -> list:
    """解析 <mergeCells>，按 (r, c) 排序返回 §3.1 的 merges 列表。"""
    merge_cells = _first_child(root, "mergeCells")
    if merge_cells is None:
        return []
    result = []
    for merge_el in _children(merge_cells, "mergeCell"):
        ref = (_attr(merge_el, "ref") or "").strip()
        if not ref:
            continue
        pieces = ref.split(":")
        try:
            r1, c1 = parse_a1(pieces[0])
            r2, c2 = parse_a1(pieces[1]) if len(pieces) > 1 else (r1, c1)
        except ValueError:
            continue
        if r2 < r1:
            r1, r2 = r2, r1
        if c2 < c1:
            c1, c2 = c2, c1
        result.append({
            "ref": "%s:%s" % (make_a1(r1, c1), make_a1(r2, c2)),
            "r": r1,
            "c": c1,
            "rowspan": r2 - r1 + 1,
            "colspan": c2 - c1 + 1,
        })
    result.sort(key=lambda item: (item["r"], item["c"]))
    return result


def _covered_cells(merges) -> dict:
    """合并区域（除 anchor 外）-> anchor 坐标 dict[(r,c)] = (r, c)。"""
    covered = {}
    for merge in merges:
        for r in range(merge["r"], merge["r"] + merge["rowspan"]):
            for c in range(merge["c"], merge["c"] + merge["colspan"]):
                if (r, c) == (merge["r"], merge["c"]):
                    continue
                covered[(r, c)] = (merge["r"], merge["c"])
    return covered


def _sheet_sizes(root):
    """解析 row_heights / col_widths / hidden / defaults / orientation。"""
    row_heights = {}
    hidden_rows = []
    col_widths = {}
    hidden_cols = []

    sheet_data = _first_child(root, "sheetData")
    if sheet_data is not None:
        for row_el in _children(sheet_data, "row"):
            row_attr = _attr(row_el, "r")
            if not row_attr:
                continue
            try:
                row_number = int(row_attr)
            except ValueError:
                continue
            height = _attr(row_el, "ht")
            if height is not None:
                try:
                    row_heights[str(row_number)] = float(height)
                except ValueError:
                    pass
            if (_attr(row_el, "hidden") or "").strip().lower() in ("1", "true"):
                hidden_rows.append(row_number)

    cols_el = _first_child(root, "cols")
    if cols_el is not None:
        for col_el in _children(cols_el, "col"):
            try:
                min_col = int(_attr(col_el, "min") or 0)
                max_col = int(_attr(col_el, "max") or min_col)
            except ValueError:
                continue
            if min_col < 1 or max_col < min_col:
                continue
            if max_col - min_col > 16384:  # 防御异常文件
                max_col = min_col
            width_raw = _attr(col_el, "width")
            width = None
            if width_raw is not None:
                try:
                    width = float(width_raw)
                except ValueError:
                    width = None
            hidden = (_attr(col_el, "hidden") or "").strip().lower() in ("1", "true")
            for col_index in range(min_col, max_col + 1):
                if width is not None:
                    col_widths[str(col_index)] = width
                if hidden:
                    hidden_cols.append(col_index)

    sheet_format = _first_child(root, "sheetFormatPr")
    default_col_width = _FALLBACK_COL_WIDTH
    default_row_height = _FALLBACK_ROW_HEIGHT
    if sheet_format is not None:
        raw = _attr(sheet_format, "defaultColWidth")
        if raw is not None:
            try:
                default_col_width = float(raw)
            except ValueError:
                pass
        raw = _attr(sheet_format, "defaultRowHeight")
        if raw is not None:
            try:
                default_row_height = float(raw)
            except ValueError:
                pass

    orientation = "portrait"
    page_setup = _first_child(root, "pageSetup")
    if page_setup is not None:
        value = _attr(page_setup, "orientation")
        if value:
            orientation = value

    hidden_rows.sort()
    hidden_cols.sort()
    return {
        "row_heights": row_heights,
        "col_widths": col_widths,
        "default_col_width": default_col_width,
        "default_row_height": default_row_height,
        "hidden_rows": hidden_rows,
        "hidden_cols": hidden_cols,
        "orientation": orientation,
    }


def _sheet_has_text_or_visible_style(cell: dict, default_font: dict, is_anchor: bool) -> bool:
    """§3.1：有文本、或有可见样式（fill/border/字体非默认等）才输出。"""
    if is_anchor:
        return True  # 合并区域 anchor 一律输出，UI 靠它画结构块
    if cell["text"]:
        return True
    style = cell.get("style") or {}
    if style.get("fill") or style.get("border") or style.get("bold") or style.get("italic"):
        return True
    if style.get("color"):
        return True
    if style.get("align") or style.get("valign") or style.get("wrap"):
        return True
    if style.get("font") != default_font.get("name"):
        return True
    if abs(float(style.get("size", 11.0)) - float(default_font.get("size", 11.0))) > 1e-9:
        return True
    return False


def read_sheet(data: bytes, index: int) -> dict:
    """读取第 index 个工作表，返回 ARCHITECTURE §3.1 的 sheet dict。"""
    archive = _open_package(data)
    sheets = _workbook_sheets(archive)
    try:
        sheet_index = int(index)
    except (TypeError, ValueError):
        raise ValueError("工作表索引无效：%r" % (index,)) from None
    if sheet_index < 0 or sheet_index >= len(sheets):
        raise ValueError("工作表索引超出范围：%d" % sheet_index)

    info = sheets[sheet_index]
    root = _parse_xml(archive.read(info["path"]))
    if root is None:
        raise ValueError("不是有效的 .xlsx 文件")

    shared_strings = _load_shared_strings(archive)
    styles = _load_styles(archive)
    merges = _parse_merges(root)
    covered = _covered_cells(merges)

    cells = []
    sheet_data = _first_child(root, "sheetData")
    if sheet_data is not None:
        for row_el in _children(sheet_data, "row"):
            row_attr = _attr(row_el, "r")
            try:
                row_number = int(row_attr) if row_attr else 0
            except ValueError:
                row_number = 0
            next_col = 1
            for cell_el in _children(row_el, "c"):
                ref = _attr(cell_el, "r")
                if ref:
                    try:
                        row_index, col_index = parse_a1(ref)
                    except ValueError:
                        continue
                else:
                    if row_number <= 0:
                        continue
                    row_index, col_index = row_number, next_col
                next_col = col_index + 1
                key = (row_index, col_index)
                if key in covered:
                    continue  # 合并区域非左上角不输出
                text = _cell_element_text(cell_el, shared_strings)
                style = _style_for(styles, _attr(cell_el, "s"))
                is_anchor = any(m["r"] == row_index and m["c"] == col_index for m in merges)
                cell = {"r": row_index, "c": col_index, "text": text, "style": style}
                if _sheet_has_text_or_visible_style(cell, styles.default_font, is_anchor):
                    cells.append(cell)

    # 合并区域 anchor 必须出现（即使 XML 里没有对应的 <c>），UI 靠它画结构块
    present = {(cell["r"], cell["c"]) for cell in cells}
    for merge in merges:
        key = (merge["r"], merge["c"])
        if key not in present:
            cells.append({
                "r": merge["r"],
                "c": merge["c"],
                "text": "",
                "style": _style_for(styles, None),
            })
            present.add(key)

    cells.sort(key=lambda item: (item["r"], item["c"]))
    rows, cols = _sheet_extent(root, cells)
    sizes = _sheet_sizes(root)

    sheet = {
        "index": sheet_index,
        "name": info["name"],
        "rows": rows,
        "cols": cols,
        "merges": merges,
        "cells": cells,
    }
    sheet.update(sizes)
    return sheet


def list_sheets(data: bytes) -> list:
    """列出工作表：[{index, name, hidden, rows, cols}]。"""
    archive = _open_package(data)
    sheets = _workbook_sheets(archive)
    result = []
    for info in sheets:
        rows = cols = 1
        try:
            root = _parse_xml(archive.read(info["path"]))
            if root is not None:
                rows, cols = _sheet_extent(root)
        except KeyError:
            pass
        result.append({
            "index": info["index"],
            "name": info["name"],
            "hidden": info["hidden"],
            "rows": rows,
            "cols": cols,
        })
    return result


def sheet_xml_path(data: bytes, index: int) -> str:
    """返回 sheet 在 xlsx（zip）内的路径，如 'xl/worksheets/sheet1.xml'。"""
    archive = _open_package(data)
    sheets = _workbook_sheets(archive)
    try:
        sheet_index = int(index)
    except (TypeError, ValueError):
        raise ValueError("工作表索引无效：%r" % (index,)) from None
    if sheet_index < 0 or sheet_index >= len(sheets):
        raise ValueError("工作表索引超出范围：%d" % sheet_index)
    return sheets[sheet_index]["path"]
