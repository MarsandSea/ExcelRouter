"""读表与写值的公共底层：拆分（splitter）和汇总（collect）共用，只用 openpyxl / xlrd。

Copyright (c) 2026 AbeLin · MIT License

放在这里的：表头识别、只读逐 sheet 读值、.xls 读成内存工作簿、隐藏的 sheet 与列、按列名对齐的列键、
合计行识别、写值时防止「=」开头的文字变成公式、用 zipfile 打包。

规矩（内部 MobileWork 技能会原样带上这个文件，那边的审核按函数名和字眼扫代码）：不导入 pandas，
不碰删文件、起子进程、读环境变量这类调用；注释和字符串避开内部打包的禁用字眼。tests/test_xlio.py 守着这几条。
"""
import os
import zipfile

import openpyxl
from openpyxl.cell.cell import ERROR_CODES, ILLEGAL_CHARACTERS_RE

from core.utils import hard_clean, soft_clean

HEAD_SCAN = 20           # 表头只在前这么多行里找（与引擎一致）
MATCH_RATIO = 0.5        # 能对上的列名至少占这个 sheet 列名的一半，否则当作表头对不上
TOTAL_WORDS = ("合计", "小计", "总计")


# =====================================================
# 表头识别
# =====================================================

def _looks_numeric(v):
    """单元格内容看起来像数字（含带千分位/百分号/货币符号的文本）。"""
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    s = soft_clean(v)
    if s == "":
        return False
    s = s.replace(',', '').replace('%', '').replace('¥', '').replace('￥', '').replace('$', '')
    try:
        float(s)
        return True
    except ValueError:
        return False


def detect_header_row_auto(rows, max_scan=15):
    """通用启发式：从前若干行的值矩阵里挑「最像表头的行」。

    表头行的特征：非空单元格多、内容多为文字标签（而非数字）、且下一行起开始出现数据。
    rows 是「行 -> 该行各列值」的二维列表。返回 1 基行号，找不到返回 -1。
    """
    best_idx, best_score = -1, -1.0
    limit = min(len(rows), max_scan)
    for r in range(limit):
        cells = rows[r]
        nonempty = [c for c in cells if soft_clean(c) != ""]
        if not nonempty:
            continue
        n = len(nonempty)
        # 文字占比：表头通常是文字标签，少有纯数字
        str_like = sum(1 for c in nonempty if not _looks_numeric(c))
        str_ratio = str_like / n
        # 数据感：下一个非空行里数字占比越高，越说明这行是表头。
        # 表头正下方偶尔会有一整行空行（分隔用），若只看紧邻的下一行会看到全空、data_hint
        # 归零，导致真正的表头因为"看起来不像后面跟数据"而在打分上输给深处某行数据行
        # （尤其像点名表这种大多列是姓名/工号/日期的宽表，数据行本身就有一半以上"看起来不像数字"）。
        # 跳过空行往下找最近的非空行来判断，最多向下看 3 行。
        # 分母取「本行、下一行」里非空格更多的那个：下一行比本行稀疏时（比如表尾只填了金额的合计行），
        # 它只能证明本行一小部分列下面跟着数字。旧算法按下一行自己的格数算，那行唯一的数字就让
        # 最后一条数据行拿满分；日期列又不算数字，数据行看着和表头一样像文字，整张表的表头就被认成
        # 最后一条数据行，按它找不到拆分字段，整个 sheet 被跳过。
        data_hint = 0.0
        for look in range(r + 1, min(r + 4, len(rows))):
            nxt = [c for c in rows[look] if soft_clean(c) != ""]
            if nxt:
                data_hint = sum(1 for c in nxt if _looks_numeric(c)) / max(len(nxt), n)
                break
        score = n * (0.5 + 0.5 * str_ratio) + n * 0.3 * data_hint
        if score > best_score:
            best_score, best_idx = score, r
    return best_idx + 1 if best_idx >= 0 else -1


def find_header_row(rows, grid_keys, id_keys, max_scan=15):
    """关键词法：在前若干行中找同时包含「网格关键词」和「工号关键词」的表头行。

    rows 为值矩阵，返回 1 基行号，找不到返回 -1。grid_keys/id_keys 任一为空则不启用该条件。
    """
    for r in range(min(len(rows), max_scan)):
        row_str = "".join(hard_clean(c) for c in rows[r])
        ok_grid = (not grid_keys) or any(hard_clean(k) in row_str for k in grid_keys)
        ok_id = (not id_keys) or any(hard_clean(k) in row_str for k in id_keys)
        if ok_grid and ok_id:
            return r + 1
    return -1


def resolve_header_row(rows, config):
    """按 config['header_mode'] 分发：auto / row(指定行号) / keyword。返回 1 基行号或 -1。"""
    mode = config.get('header_mode', 'auto')
    if mode == 'row':
        try:
            hr = int(config.get('header_row', 1))
        except (TypeError, ValueError):
            return -1
        return hr if hr >= 1 else -1
    if mode == 'keyword':
        grid_keys, id_keys = config.get('grid_keys', []), config.get('id_keys', [])
        if not grid_keys and not id_keys:
            # 两类关键词都没填时，find_header_row 的"任一为空则不启用该条件"会退化成
            # "永远命中第 1 行"——多 sheet 表格里真实表头不在第 1 行的 sheet 会被强行
            # 按错误的行去找列，导致拆分字段/到人字段找不到而被静默跳过。关键词法本意是
            # "专家兜底"，没配关键词就不该比 auto 启发式更武断，退回 auto。
            return detect_header_row_auto(rows)
        return find_header_row(rows, grid_keys, id_keys)
    return detect_header_row_auto(rows)


def _head_limit(cfg):
    """识别表头时读前几行：一般是 HEAD_SCAN 行；指定了表头行号且更靠后时，读到那一行。"""
    if cfg.get("header_mode") == "row":
        try:
            return max(HEAD_SCAN, int(cfg.get("header_row", 1)))
        except (TypeError, ValueError):
            pass
    return HEAD_SCAN


# =====================================================
# 行与单元格的值
# =====================================================

def _nonempty(row):
    return any(soft_clean(v) != "" for v in row)


def _cell(row, ci):
    return row[ci] if ci is not None and ci < len(row) else None


def is_total_row(row):
    """行里前 3 个有内容的格子有「合计 / 小计 / 总计」字样（如「合计」「全市合计」）：这是各表自己的合计行。"""
    for v in [v for v in row if soft_clean(v) != ""][:3]:
        t = hard_clean(v)
        if 0 < len(t) <= 8 and t.endswith(TOTAL_WORDS):
            return True
    return False


def cell_value(cell):
    """单元格的值；错误格（#N/A、#DIV/0! 等）当作空。"""
    return None if cell.data_type == "e" else cell.value


def key_value(v):
    """拿来分组、比列名的值：整数形式的小数当整数（3.0 → 3），Excel 错误码当作空。

    只读模式按值读取时拿不到单元格类型，错误格读出来就是「#N/A」这样的文字；这里统一当空，
    预览取值和真正拆分时的口径才一致。
    """
    if isinstance(v, float) and v == v and v.is_integer():
        return int(v)
    if isinstance(v, str) and v in ERROR_CODES:
        return None
    return v


def put(cell, v):
    """写值。openpyxl 会把以「=」开头的文字当成公式，这里把类型钉回文字：表格里写着「=…」的文字原样保留。

    真公式要写成公式时，调用方在这之后直接给 cell.value 赋公式文本。
    """
    cell.value = v
    if isinstance(v, str) and v.startswith("=") and cell.data_type == "f":
        cell.data_type = "s"


def _legal(v):
    return ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v


def _no_formulas(wb):
    """保存前把产出里所有「公式」单元格定为文字。

    值都来自计算好的结果，本来没有公式；但 openpyxl 写入以「=」开头的文字时会自动当成公式，
    源表里一个写着「=…」的文字格子，复制过来就会变成能执行的公式。这里改回文字，打开时原样显示。
    """
    for ws in wb.worksheets:
        for cell in ws._cells.values():
            if cell.data_type == "f":
                cell.data_type = "s"


# =====================================================
# 读表
# =====================================================

def _xls_sheets(path):
    """用 xlrd 读 .xls 的值（.xls 的格式本来就保不住），返回 [(sheet 名, [行元组])]。"""
    import xlrd
    book = xlrd.open_workbook(path)
    out = []
    for sh in book.sheets():
        rows = []
        for r in range(sh.nrows):
            vals = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                v = cell.value
                if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
                    v = None
                elif cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        v = xlrd.xldate.xldate_as_datetime(v, book.datemode)
                    except Exception:
                        pass
                elif cell.ctype == xlrd.XL_CELL_NUMBER and float(v).is_integer():
                    v = int(v)
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    v = bool(v)
                vals.append(v)
            rows.append(tuple(vals))
        out.append((sh.name, rows))
    return out


def _iter_sheets_ro(path):
    """只读方式逐 sheet 产出 (sheet 名, 行迭代器)：快、省内存。中间缺的空行也会产出（空元组），行号不会错位。"""
    if path.lower().endswith(".xls"):
        for name, rows in _xls_sheets(path):
            yield name, iter(rows)
        return
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            ws.reset_dimensions()   # 有的程序生成的表把尺寸写错，不重置的话只读模式会少读行列
            yield ws.title, ws.iter_rows(values_only=True)
    finally:
        wb.close()


_NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_NS_PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _xlsx_hidden(path):
    """直接读 xlsx 包里的 XML：{sheet 名: (状态, {隐藏列 1 基列号})}。

    只读模式拿不到列属性，整本加载又太慢；<cols> 在 <sheetData> 之前，读到那里就停。
    列宽小到看不见（< 0.5）的列也算隐藏。
    """
    import posixpath
    from xml.etree.ElementTree import fromstring, iterparse
    out = {}
    with zipfile.ZipFile(path) as z:
        rels = {r.get("Id"): r.get("Target") for r in fromstring(z.read("xl/_rels/workbook.xml.rels"))
                if r.tag == _NS_PKG + "Relationship"}
        sheets = fromstring(z.read("xl/workbook.xml")).find(_NS_MAIN + "sheets")
        for sh in (sheets if sheets is not None else []):
            target = rels.get(sh.get(_NS_REL + "id"), "")
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            cols = set()
            try:
                with z.open(part) as f:
                    for _ev, el in iterparse(f, events=("start",)):
                        if el.tag == _NS_MAIN + "sheetData":
                            break
                        if el.tag == _NS_MAIN + "col":
                            width = el.get("width")
                            hidden = el.get("hidden") in ("1", "true") or (
                                el.get("customWidth") in ("1", "true") and width is not None and float(width) < 0.5)
                            if hidden:
                                cols.update(range(int(el.get("min")), int(el.get("max")) + 1))
            except (KeyError, ValueError):
                pass
            out[sh.get("name")] = (sh.get("state") or "visible", cols)
    return out


def _xls_hidden(path):
    """.xls 的隐藏 sheet 与隐藏列（xlrd 要 formatting_info 才有列信息）。"""
    import xlrd
    out = {}
    book = xlrd.open_workbook(path, formatting_info=True, on_demand=True)
    try:
        for i, name in enumerate(book.sheet_names()):
            sh = book.sheet_by_index(i)
            cols = {c + 1 for c, info in sh.colinfo_map.items() if info.hidden or info.width < 128}   # 1/256 字符宽
            out[name] = ("visible" if sh.visibility == 0 else "hidden", cols)
            book.unload_sheet(i)
    finally:
        book.release_resources()
    return out


def hidden_info(path):
    """{sheet 名: (状态 visible|hidden|veryHidden, {隐藏列 1 基列号})}；读不出来时返回 {}（当作没有隐藏）。"""
    try:
        return _xls_hidden(path) if path.lower().endswith(".xls") else _xlsx_hidden(path)
    except Exception:
        return {}


def _load_sheets(path):
    """完整加载一个源文件，返回 (wb, [(sheet 名, 行值列表, ws)], 是否 .xls)。

    .xlsx 用 openpyxl 加载一次，值和格式都从这一份取；.xls 转成只有值的内存工作簿，复制路径与 .xlsx 相同。
    """
    if path.lower().endswith(".xls"):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        sheets = []
        for name, rows in _xls_sheets(path):
            ws = wb.create_sheet(title=name[:31])
            for r in rows:
                ws.append([_legal(v) for v in r])
            sheets.append((name, rows, ws))
        return wb, sheets, True
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    return wb, [(ws.title, list(ws.iter_rows(values_only=True)), ws) for ws in wb.worksheets], False


def open_book(path):
    """完整加载一个源文件，返回 (wb, 是否 .xls)：.xlsx / .xlsm 读计算好的值和格式；.xls 读成只有值的内存工作簿。"""
    if path.lower().endswith(".xls"):
        wb, _sheets, _ = _load_sheets(path)
        return wb, True
    return openpyxl.load_workbook(path, read_only=False, data_only=True), False


# =====================================================
# 表头与列：按列名对齐
# =====================================================

def col_names(head, h):
    """表头第 h 行（1 基）的列名。多行表头里某列在这一行是空的（上面纵向合并的格子），往上找两行。"""
    row = head[h - 1]
    names = []
    for c in range(len(row)):
        v, k = soft_clean(row[c]), h - 2
        while not v and k >= max(0, h - 3):
            v = soft_clean(_cell(head[k], c))
            k -= 1
        names.append(v)
    while names and not names[-1]:
        names.pop()
    return names


def col_keys(names):
    """列名 → 对齐用的键：只留字母数字汉字后比对；重名的按出现顺序加「#2」；没有列名的按位置对齐。"""
    seen, keys = {}, []
    for c, n in enumerate(names):
        k = hard_clean(n).lower()
        if not k:
            keys.append(f"#{c + 1}")
            continue
        seen[k] = seen.get(k, 0) + 1
        keys.append(k if seen[k] == 1 else f"{k}#{seen[k]}")
    return keys


def match(tpl_keys, keys):
    """返回 (对上的列数, 是否表头对不上)。只算有列名的列。"""
    named = [k for k in keys if not k.startswith("#")]
    hit = sum(1 for k in named if k in set(tpl_keys))
    return hit, (not named or hit < max(1, MATCH_RATIO * len(named)))


# =====================================================
# 打包
# =====================================================

def zip_folder(folder, zip_path):
    """把 folder 里的文件打成 zip_path（包里的路径相对 folder，用「/」分隔）。返回打进去的文件数。

    zip_path 放在 folder 里面也不会把自己打进去。
    """
    n = 0
    me = os.path.normcase(os.path.abspath(zip_path))
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(folder):
            dirs.sort()
            for fn in sorted(files):
                path = os.path.join(root, fn)
                if os.path.normcase(os.path.abspath(path)) == me:
                    continue
                z.write(path, os.path.relpath(path, folder).replace(os.sep, "/"))
                n += 1
    return n
