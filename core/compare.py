"""两表核对：两张表按关键列逐行对上，找出只在一边的行、改过的格子，写成一个核对结果工作簿。

Copyright (c) 2026 AbeLin · MIT License

    ta = read_table(path_a)                 # 表头自动识别，按列名对齐
    tb = read_table(path_b)
    keys = suggest_keys(ta, tb)             # 推荐关键列（单列或两列组合）
    res = compare(ta, tb, keys[0]["cols"])
    write_result(res, ta, tb, out_path)

比较口径：数字按数值比（千分位、货币符号、文字形式的数字、百分号都先换成数值），相差不到 TOL 算一样；
日期格和「2026-10-01」这类文字按日期比；文字统一全角半角、去掉首尾空白再比；空格子和空文字算一样。
关键列的值统一全角半角、去掉空白、字母大写；123 和「123」是同一个，前导零保留（00123 和 123 不同）。
关键列在任何一张表里重复或空着的行不参与配对，单独列出，免得配错。空行和各表自己的合计行跳过。

规矩同 xlio：不导入 pandas，不碰删文件、起子进程、读环境变量这类调用，注释和字符串避开内部打包的禁用字眼。
"""
import datetime
import itertools
import re
import unicodedata

import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill

from core.utils import hard_clean, soft_clean
from core.xlio import (HEAD_SCAN, _iter_sheets_ro, _legal, col_keys, col_names, hidden_info, is_total_row,
                       key_value, put, resolve_header_row)

MAX_CELLS = 1_000_000    # 两张表合起来最多这么多个格子（本机读 100 万格约 26 秒，内部技能一条命令要在 70 秒内做完）
TOL = 0.005               # 数字相差不到这么多算一样
KEY_RATIO = 0.95          # 推荐关键列：两表都至少这么多行填了、且不重复（个别重复的照样推荐，重复行单独列出）
KEY_OVERLAP = 0.3         # 推荐关键列：两表的值至少这么多能对上（按少的那张表算）
MAX_KEY_COLS = 3          # 关键列最多几列合起来
PAIR_POOL = 8             # 单列当不了关键列时，从这么多列里试两列组合
ID_WORDS = ("工号", "编号", "编码", "代码", "身份证", "证件", "账号", "帐号", "手机", "电话", "学号", "id", "no", "code")
NOT_KEY_WORDS = ("序号", "行号", "排名", "名次")   # 换个顺序就全变，当不了关键列
SHEETS = ("概要", "有变化", "变化行", "只在A", "只在B", "关键列重复")
RED = PatternFill("solid", fgColor="FFC7CE")
BOLD = Font(bold=True)
_DATE_TEXT = re.compile(r"^(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?"
                        r"(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$")
_UNNAMED = re.compile(r"第\d+列")    # 没有列名的列（按位置对齐），当不了关键列
_NUM_STRIP = str.maketrans("", "", ",¥$ ")


# =====================================================
# 读表
# =====================================================

class TableError(Exception):
    """表读不了或认不出表头；消息给用户看。"""


class TooBig(TableError):
    """格子数超过 max_cells，读到一半停下。"""


def _header_cfg(header_row):
    return {"header_mode": "row", "header_row": int(header_row)} if header_row else {"header_mode": "auto"}


def data_sheets(path):
    """[sheet 名]：没隐藏的 sheet，按顺序。"""
    hidden = hidden_info(path)
    return [n for n, _rows in _iter_sheets_ro(path) if hidden.get(n, ("visible",))[0] == "visible"]


def read_table(path, sheet=None, header_row=0, max_cells=None):
    """读一张表：sheet 不给时取第一个有表头和数据、没隐藏的 sheet。超过 max_cells 个格子时抛 TooBig。

    返回 {"path", "sheet", "h": 表头行号, "names": 列名, "keys": 列键, "rows": [(行号, 值元组)],
          "totals": 跳过的合计行数, "first": 第一条数据所在行}。值都是计算好的值（不读公式）。
    """
    cfg = _header_cfg(header_row)
    hidden = hidden_info(path)
    seen = []
    try:
        for name, it in _iter_sheets_ro(path):
            seen.append(name)
            if sheet is not None and name != sheet:
                continue
            if sheet is None and hidden.get(name, ("visible",))[0] != "visible":
                continue
            t = _read_sheet(path, name, it, cfg, max_cells)
            if t is not None:
                return t
            if sheet is not None:
                raise TableError(f"「{sheet}」里没认出表头和数据")
    except TableError:
        raise
    except Exception as e:
        raise TableError(f"读不了：{type(e).__name__}") from e
    if sheet is not None and sheet not in seen:
        raise TableError(f"没有叫「{sheet}」的 sheet")
    raise TableError("没认出带表头和数据的 sheet")


def _read_sheet(path, name, it, cfg, max_cells=None):
    head = []
    for row in it:
        head.append(tuple(key_value(v) for v in row))
        if len(head) >= max(HEAD_SCAN, int(cfg.get("header_row", 1) or 1)):
            break
    h = resolve_header_row(head, cfg)
    if h < 1 or h > len(head):
        return None
    names = col_names(head, h)
    if not any(names):
        return None
    rows, totals = [], 0

    def take(r, row):
        nonlocal totals
        if not any(soft_clean(v) != "" for v in row):
            return
        if is_total_row(row):
            totals += 1
            return
        rows.append((r, row))

    for r, row in enumerate(head[h:], start=h + 1):
        take(r, row)
    cap = max_cells // max(1, len(names)) if max_cells else None
    for r, row in enumerate(it, start=len(head) + 1):
        take(r, tuple(key_value(v) for v in row))
        if cap is not None and len(rows) > cap:
            raise TooBig(f"「{name}」超过 {max_cells} 个格子")
    if not rows:
        return None
    return {"path": path, "sheet": name, "h": h, "names": names, "keys": col_keys(names), "rows": rows,
            "totals": totals, "first": rows[0][0]}


# =====================================================
# 值的比较口径
# =====================================================

def _nfkc(v):
    return unicodedata.normalize("NFKC", soft_clean(v))


def _as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date() if v.time() == datetime.time(0) else v
    return v


def norm(v):
    """(类别, 值)：类别是 ""（空）/ "n"（数字）/ "d"（日期）/ "b"（是否）/ "s"（文字）。"""
    if v is None:
        return ("", None)
    if isinstance(v, bool):
        return ("b", v)
    if isinstance(v, (int, float)):
        return ("n", float(v)) if v == v else ("", None)
    if isinstance(v, (datetime.datetime, datetime.date)):
        return ("d", _as_date(v))
    if isinstance(v, datetime.time):
        return ("s", v.isoformat())
    s = _nfkc(v)
    if s == "":
        return ("", None)
    n = _number_text(s)
    if n is not None:
        return ("n", n)
    m = _DATE_TEXT.match(s)
    if m:
        try:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if m.group(4):
                return ("d", _as_date(datetime.datetime(y, mo, d, int(m.group(4)), int(m.group(5)),
                                                        int(m.group(6) or 0))))
            return ("d", datetime.date(y, mo, d))
        except ValueError:
            pass
    return ("s", re.sub(r"\s+", " ", s))


def _number_text(s):
    """文字形式的数字 → 数值；前导零（00123）和超过 15 位的数字串（证件号、卡号）当文字，免得丢精度、配错。"""
    t = s.translate(_NUM_STRIP)
    pct = t.endswith("%")
    if pct:
        t = t[:-1]
    body = t.lstrip("+-")
    if not body or not re.fullmatch(r"\d+(\.\d+)?", body):
        return None
    digits = body.split(".")[0]
    if len(digits) > 1 and digits.startswith("0"):
        return None
    if len(body.replace(".", "")) > 15:
        return None
    x = float(t)
    return x / 100 if pct else x


def same(a, b):
    """两个格子算不算一样（口径见文件开头）。"""
    ka, va = norm(a)
    kb, vb = norm(b)
    if ka != kb:
        return False
    if ka == "n":
        return round(abs(va - vb), 9) < TOL
    return va == vb


def diff(a, b):
    """两个都是数字时返回 b − a，否则 None。"""
    ka, va = norm(a)
    kb, vb = norm(b)
    if ka == kb == "n":
        d = round(vb - va, 9)
        return int(d) if d.is_integer() else d
    return None


def key_text(v):
    """关键列的值 → 对比用的文字；空返回 ""。"""
    if v is None or isinstance(v, bool):
        return "" if v is None else str(v)
    if isinstance(v, float):
        if v != v:
            return ""
        return str(int(v)) if v.is_integer() else format(v, ".12g")
    if isinstance(v, int):
        return str(v)
    if isinstance(v, (datetime.datetime, datetime.date)):
        return _as_date(v).isoformat()
    return re.sub(r"\s+", "", _nfkc(v)).upper()


# =====================================================
# 列与关键列
# =====================================================

def common_columns(ta, tb):
    """(两表都有的列 [(列名, A 列号, B 列号)], 只在 A 的列名, 只在 B 的列名)；列号 0 基，按 A 的顺序。"""
    bpos = {k: i for i, k in enumerate(tb["keys"])}
    apos = {k: i for i, k in enumerate(ta["keys"])}
    both = [(ta["names"][i] or f"第{i + 1}列", i, bpos[k]) for i, k in enumerate(ta["keys"]) if k in bpos]
    only_a = [ta["names"][i] for i, k in enumerate(ta["keys"]) if k not in bpos and ta["names"][i]]
    only_b = [tb["names"][i] for i, k in enumerate(tb["keys"]) if k not in apos and tb["names"][i]]
    return both, only_a, only_b


def _cell(row, i):
    return row[i] if i < len(row) else None


def _key_of(row, idxs):
    return tuple(key_text(_cell(row, i)) for i in idxs)


def key_stats(t, idxs):
    """{"filled", "unique", "empty", "dups": 重复的行数, "values": 值集合}：按 idxs 这几列合起来当关键列。"""
    counts, empty = {}, 0
    for _r, row in t["rows"]:
        k = _key_of(row, idxs)
        if any(x == "" for x in k):
            empty += 1
            continue
        counts[k] = counts.get(k, 0) + 1
    n = len(t["rows"]) or 1
    dups = sum(c for c in counts.values() if c > 1)
    filled = n - empty
    return {"filled": filled / n, "unique": (filled - dups) / max(1, filled), "empty": empty, "dups": dups,
            "values": set(counts)}


def _id_like(name):
    low = hard_clean(name).lower()
    return any(w in low for w in ID_WORDS)


def _not_key(name):
    return any(w in name for w in NOT_KEY_WORDS)


def _candidate(ta, tb, cols, names):
    sa = key_stats(ta, [c[1] for c in cols])
    sb = key_stats(tb, [c[2] for c in cols])
    if min(sa["filled"], sb["filled"], sa["unique"], sb["unique"]) < KEY_RATIO:
        return None
    both = len(sa["values"] & sb["values"])
    overlap = both / max(1, min(len(sa["values"]), len(sb["values"])))
    if overlap < KEY_OVERLAP:
        return None
    clean = not (sa["dups"] or sb["dups"] or sa["empty"] or sb["empty"])
    score = overlap * 2 + (1 if clean else 0) + (1 if any(_id_like(n) for n in names) else 0) - 0.3 * (len(cols) - 1)
    return {"cols": names, "score": round(score, 4), "matched": both, "dup_a": sa["dups"], "dup_b": sb["dups"],
            "empty_a": sa["empty"], "empty_b": sb["empty"]}


def _mostly_amount(t, i, name):
    """这一列多半是数（金额、数量、比例），列名又不像工号编码：碰巧不重复也当不了关键列。"""
    if _id_like(name):
        return False
    vals = [_cell(row, i) for _r, row in t["rows"][:500]]
    vals = [v for v in vals if v is not None and soft_clean(v) != ""]
    num = sum(1 for v in vals if norm(v)[0] == "n")
    return bool(vals) and num > len(vals) / 2


def suggest_keys(ta, tb, limit=3):
    """推荐关键列，好的在前：[{"cols": [列名], "score", "matched": 能对上的个数, "dup_a", "dup_b", "empty_a", "empty_b"}]。

    先试单列；单列都不行时，从内容最分散的几列里试两列组合。都不行返回 []。
    """
    both, _a, _b = common_columns(ta, tb)
    pool = [c for c in both if not _UNNAMED.fullmatch(c[0]) and not _not_key(c[0])
            and not _mostly_amount(ta, c[1], c[0])]
    out = [x for c in pool if (x := _candidate(ta, tb, [c], [c[0]]))]
    if not out:
        spread = sorted(pool, key=lambda c: -len(key_stats(ta, [c[1]])["values"]))[:PAIR_POOL]
        for pair in itertools.combinations(spread, 2):
            x = _candidate(ta, tb, list(pair), [c[0] for c in pair])
            if x:
                out.append(x)
    out.sort(key=lambda x: -x["score"])
    return out[:limit]


def resolve_key(ta, tb, names):
    """用户说的关键列名 → [(列名, A 列号, B 列号)]；对不上的列名放进第二个返回值。"""
    both, _a, _b = common_columns(ta, tb)
    by = {hard_clean(c[0]).lower(): c for c in both}
    cols, missing = [], []
    for n in names:
        c = by.get(hard_clean(n).lower())
        (cols.append(c) if c else missing.append(n))
    return cols, missing


# =====================================================
# 核对
# =====================================================

def compare(ta, tb, key_names):
    """按 key_names 这几列对上两张表。返回核对结果（只有行号、列号和计数，值在写结果时再从表里取）。"""
    kcols, missing = resolve_key(ta, tb, key_names)
    if missing or not kcols:
        raise TableError(f"两张表里找不到关键列「{'、'.join(missing or key_names)}」")
    both, only_a_cols, only_b_cols = common_columns(ta, tb)
    kia, kib = [c[1] for c in kcols], [c[2] for c in kcols]
    kset = {c[1] for c in kcols}
    vcols = [c for c in both if c[1] not in kset]

    def index(t, idxs):
        groups, empty = {}, []
        for n, (_r, row) in enumerate(t["rows"]):
            k = _key_of(row, idxs)
            if any(x == "" for x in k):
                empty.append(n)
            else:
                groups.setdefault(k, []).append(n)
        return groups, empty

    ga, ea = index(ta, kia)
    gb, eb = index(tb, kib)
    bad = {k for k, v in ga.items() if len(v) > 1} | {k for k, v in gb.items() if len(v) > 1}
    unpaired = ([("A", n, "关键列空着") for n in ea] + [("B", n, "关键列空着") for n in eb]
                + [("A", n, "关键列重复") for k in ga if k in bad for n in ga[k]]
                + [("B", n, "关键列重复") for k in gb if k in bad for n in gb[k]])
    only_a = sorted(v[0] for k, v in ga.items() if k not in bad and k not in gb)
    only_b = sorted(v[0] for k, v in gb.items() if k not in bad and k not in ga)
    pairs = sorted((gb[k][0], ga[k][0]) for k in ga if k not in bad and k in gb)
    changes, changed_rows, per_col = [], [], {}
    for nb, na in pairs:
        ra, rb = ta["rows"][na][1], tb["rows"][nb][1]
        cols = [c for c in vcols if not same(_cell(ra, c[1]), _cell(rb, c[2]))]
        if cols:
            changed_rows.append((na, nb, [c[0] for c in cols]))
            for c in cols:
                changes.append((na, nb, c))
                per_col[c[0]] = per_col.get(c[0], 0) + 1
    return {"key": [c[0] for c in kcols], "kcols": kcols, "vcols": vcols, "pairs": len(pairs),
            "only_a": only_a, "only_b": only_b, "changes": changes, "changed_rows": changed_rows,
            "per_col": dict(sorted(per_col.items(), key=lambda x: -x[1])), "unpaired": unpaired,
            "only_a_cols": only_a_cols, "only_b_cols": only_b_cols,
            "rows_a": len(ta["rows"]), "rows_b": len(tb["rows"]), "totals_a": ta["totals"], "totals_b": tb["totals"]}


def counts(res):
    return {"rows_a": res["rows_a"], "rows_b": res["rows_b"], "matched": res["pairs"],
            "only_a": len(res["only_a"]), "only_b": len(res["only_b"]), "changed_rows": len(res["changed_rows"]),
            "changed_cells": len(res["changes"]), "unpaired": len(res["unpaired"])}


# =====================================================
# 写核对结果
# =====================================================

def _wcell(ws, v, bold=False, fill=None):
    c = WriteOnlyCell(ws)
    put(c, _legal(v))
    if bold:
        c.font = BOLD
    if fill is not None:
        c.fill = fill
    return c


def _header(ws, names):
    ws.append([_wcell(ws, n, bold=True) for n in names])


def _label(t, path_name):
    return f"{path_name}（{t['sheet']}）"


def write_result(res, ta, tb, out_path, name_a="A 表", name_b="B 表"):
    """写核对结果工作簿：概要、有变化、变化行、只在A、只在B、关键列重复。值写原值（日期还是日期）。"""
    wb = openpyxl.Workbook(write_only=True)
    c = counts(res)
    kcols = res["kcols"]
    ra, rb = ta["rows"], tb["rows"]

    ws = wb.create_sheet(SHEETS[0])
    _header(ws, ["项目", "内容"])
    lines = [("A 表", _label(ta, name_a)), ("B 表", _label(tb, name_b)), ("关键列", "、".join(res["key"])),
             ("A 表数据行", c["rows_a"]), ("B 表数据行", c["rows_b"]), ("对上的行", c["matched"]),
             ("其中有变化的行", c["changed_rows"]), ("改过的格子", c["changed_cells"]),
             ("只在 A 表的行", c["only_a"]), ("只在 B 表的行", c["only_b"]),
             ("关键列重复或空着、没有配对的行", c["unpaired"])]
    lines += [(f"「{k}」改了几行", v) for k, v in res["per_col"].items()]
    if res["only_a_cols"]:
        lines.append(("只在 A 表的列（不比较）", "、".join(res["only_a_cols"])))
    if res["only_b_cols"]:
        lines.append(("只在 B 表的列（不比较）", "、".join(res["only_b_cols"])))
    if res["totals_a"] or res["totals_b"]:
        lines.append(("跳过的合计行", f"A 表 {res['totals_a']} 行，B 表 {res['totals_b']} 行"))
    lines.append(("比较口径", f"数字相差不到 {TOL} 算一样；100 和 100.00、1,000 和 1000、10% 和 0.1 都算一样；"
                  "日期按日期比；文字不分全角半角、不看首尾空格；空格子和空文字算一样。"))
    for k, v in lines:
        ws.append([_wcell(ws, k), _wcell(ws, v)])

    ws = wb.create_sheet(SHEETS[1])
    _header(ws, res["key"] + ["列名", "A 表的值", "B 表的值", "差额（B−A）", "A 表行号", "B 表行号"])
    for na, nb, col in res["changes"]:
        a, b = _cell(ra[na][1], col[1]), _cell(rb[nb][1], col[2])
        ws.append([_wcell(ws, _cell(rb[nb][1], k[2])) for k in kcols]
                  + [_wcell(ws, col[0]), _wcell(ws, a), _wcell(ws, b), _wcell(ws, diff(a, b)),
                     _wcell(ws, ra[na][0]), _wcell(ws, rb[nb][0])])

    ws = wb.create_sheet(SHEETS[2])
    width = len(tb["names"])
    _header(ws, list(tb["names"]) + ["原来的值（A 表）", "B 表行号"])
    bcol = {c[0]: c for c in res["vcols"]}
    for na, nb, names in res["changed_rows"]:
        marked = {bcol[n][2] for n in names}
        row = rb[nb][1]
        was = "；".join(f"{n}：{_text(_cell(ra[na][1], bcol[n][1]))}" for n in names)
        ws.append([_wcell(ws, _cell(row, i), fill=RED if i in marked else None) for i in range(width)]
                  + [_wcell(ws, was), _wcell(ws, rb[nb][0])])

    for title, t, idxs in ((SHEETS[3], ta, res["only_a"]), (SHEETS[4], tb, res["only_b"])):
        ws = wb.create_sheet(title)
        w = len(t["names"])
        _header(ws, list(t["names"]) + ["所在行"])
        for n in idxs:
            r, row = t["rows"][n]
            ws.append([_wcell(ws, _cell(row, i)) for i in range(w)] + [_wcell(ws, r)])

    ws = wb.create_sheet(SHEETS[5])
    _header(ws, ["出自", "所在行", "原因"] + res["key"])
    for side, n, why in res["unpaired"]:
        t = ta if side == "A" else tb
        r, row = t["rows"][n]
        pos = 1 if side == "A" else 2
        ws.append([_wcell(ws, f"{side} 表"), _wcell(ws, r), _wcell(ws, why)]
                  + [_wcell(ws, _cell(row, k[pos])) for k in kcols])
    wb.save(out_path)
    return out_path


def _text(v):
    if v is None:
        return "（空）"
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d %H:%M") if v.time() != datetime.time(0) else v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return soft_clean(v)
