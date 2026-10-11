"""Excel 汇总：把同一套模板的多个表合成一张；表是发出去让大家填的，还能先查收通报。
全程只用 openpyxl，不导入 pandas。

Copyright (c) 2026 AbeLin · MIT License

三个入口：
  · 普通汇总（merge_tables）：文件夹里的表（或一个文件里的各个 sheet）合成一张，每行标上来源，
    最后一个 sheet「汇总说明」列出每个来源合了几行、哪些没合进去、为什么。不需要名单。
  · 查收（check_collect）：收回几份就查一次，几秒出结果，不合并。按应交名单逐人列出
    未交 / 交了空表 / 有漏填（哪一列、第几行）/ 和下发时不一样（少了行、改了不该改的格子）/ 合格，
    写「回收进度.xlsx」「回收通报.txt」「回收结果.json」（逐人的问题编码，见 ISSUE_STATUS），
    并记下这一轮，下一轮能说出新交了谁、谁改好了、谁还差。
  · 查收后汇总（run_collect）：只合交了的表，「来源」写交表人，漏填的格子标浅红，
    「汇总说明」逐人列出状态和行数。

谁该交（名单）：config['collect_expect'] → 下发记录.json 里的分组 → 文件夹里的「应交名单」「发送清单」表。
文件是谁的（认领）：文件里「填报说明」sheet 写的填报对象 → 文件名（含所在文件夹名）里的名字，长名字优先，
    「城区二」的表不会被认成「城区」→ 分组列（config['collect_by']）里占多数的取值。
哪些列该填：config['collect_required'] → 下发记录里的填写列 → 大多数表都填满了的列（备注、说明这类除外）。
一张大家各填各行的总表也能查：输入单个文件并指定 config['collect_by']，按这一列的取值分人。

合并规则（与内部技能共用，分段执行见 merge_start / merge_step）：
  · 表头以第一个表为准（值、格式、合并单元格、列宽原样），其余表按列名对齐，列顺序不同也能合；
    缺的列留空，多出的列放在最后；能对上的列名不到一半，就当作表头对不上，这个 sheet 不合。
  · 每行最后加一列「来源」；各表自己的合计 / 小计行不合进来；内容完全相同的文件只合一份；
    隐藏的 sheet 与「填报说明」sheet 不合。
  · 产出里没有公式：源表按计算好的值读取，以「=」开头的文字保存前定为文字（_no_formulas）。

下发记录.json（拆分时指定了「让对方填写的列」才会写）只有分组名、文件名、列名、行数和每行的指纹（短哈希），
不含任何单元格内容；查收时用指纹比对，找出被删掉的行和被改动的格子。
"""
import datetime
import hashlib
import json
import os
import re
import time
import unicodedata
from collections import Counter
from copy import copy
from itertools import islice

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from core import splitter as sp
from core.utils import hard_clean, safe_filename, soft_clean
from core.xlio import (HEAD_SCAN, _cell, _head_limit, _iter_sheets_ro, _load_sheets, _no_formulas, _nonempty,
                       col_keys, col_names, hidden_info, is_total_row, match)

SOURCE_COL = "来源"
REPORT_SHEET = "汇总说明"
GUIDE_SHEET = "填报说明"
RECORD_NAME = "下发记录.json"
PROGRESS_NAME = "回收进度.xlsx"
NOTICE_NAME = "回收通报.txt"
ROUNDS_NAME = "回收记录.json"
RESULT_NAME = "回收结果.json"
RESULT_DIR = "汇总结果"
CLAIM_LABELS = ("填报对象", "下发对象", "填报单位", "填报人")       # 「填报说明」sheet 里写明这份表是谁的
LIST_WORDS = ("应交名单", "发送清单", "分发清单")                   # 文件名含这些字的是名单，不是交回来的表
NAME_HEADS = ("应交", "填报对象", "分组", "名称", "单位", "网格", "姓名", "接收人")
FILLED_SHARE = 0.8       # 没指定该填的列时：总体八成以上的行填了、且大多数表填满的列算该填的
OPTIONAL_WORDS = ("备注", "说明", "附注", "其他", "其它")   # 这类列本来就可填可不填，不算该填的
BLANK_FILL = PatternFill("solid", fgColor="FFC7CE")         # 该填没填的格子在结果里标浅红
BLANK_KEEP = 500         # 每列最多记这么多个漏填的行号（个数照常全数）
SHOW_ROWS = 8            # 通报里每列最多列出这么多个行号
ROUNDS_KEEP = 30         # 回收记录.json 最多留这么多轮
SAVE_PER_CELL = 20e-6    # 结果工作簿存盘每个格子约用的秒数（本机实测约 17 微秒），没存过盘时用它估算
# 分段续做时要把已合的结果整本读回再存盘，结果越大每段能做的越少。内部技能用它限制一次合并的量。
MAX_CELLS = 600_000

ST_MISSING, ST_EMPTY, ST_PARTIAL, ST_CHANGED, ST_BAD, ST_OK = "未交", "空表", "有漏填", "有改动", "表头对不上", "合格"
ST_INVALID = "填写有误"
EXTRA = "名单外"
_RANK = {ST_MISSING: 0, ST_EMPTY: 1, ST_PARTIAL: 2, ST_INVALID: 3, ST_CHANGED: 4, ST_BAD: 5, ST_OK: 6}
TABLE_EXTS = (".xlsx", ".xlsm", ".xls")
# 收回来的文件里常见、但读不了的格式：照样算「交了」（不然通报会把交了的人写成未交），原因写成人话
OTHER_EXTS = {".et": "是 WPS 表格自己的格式（.et）", ".ett": "是 WPS 表格模板（.ett）",
              ".csv": "是 CSV 文本", ".xlsb": "是二进制工作簿（.xlsb）", ".ods": "是 OpenDocument 表格（.ods）",
              ".numbers": "是苹果 Numbers 表格"}
_STATUS_COLOR = {ST_MISSING: "FFC7CE", ST_EMPTY: "FFC7CE", ST_PARTIAL: "FFEB9C", ST_INVALID: "FFEB9C",
                 ST_CHANGED: "FFEB9C", ST_BAD: "FFEB9C", ST_OK: "C6EFCE"}
INVALID_FILL = PatternFill("solid", fgColor="FFEB9C")       # 填错的格子（汇总时认不出、没改写法的）标浅橙
_MASK = (1 << 64) - 1


# ---------- 读表 ----------

def unreadable(path):
    """读不了的文件返回人话原因（带上怎么办），读得了返回 ""。只看扩展名和文件头几百字节。

    不先判断的话，用户和通报里看到的是「File is not a zip file」「Expected BOF record」这类原话。
    """
    ext = os.path.splitext(path)[1].lower()
    save_as = "请用 Excel 或 WPS 打开后另存为 .xlsx"
    if ext in OTHER_EXTS:
        return f"{OTHER_EXTS[ext]}，{save_as}"
    try:
        with open(path, "rb") as f:
            head = f.read(512)
    except OSError as e:
        return f"读取失败：{e}"
    if not head:
        return "是空文件"
    if ext in (".xlsx", ".xlsm"):
        if head.startswith(b"\xd0\xcf\x11\xe0"):
            return "设置了打开口令（或其实是 .xls 改了扩展名），请去掉口令、另存为 .xlsx"
        if not head.startswith(b"PK"):
            return "不是完整的 Excel 文件（可能没下载完或已损坏），请重新另存一份"
    elif ext == ".xls":
        low = head.lstrip(b"\xef\xbb\xbf \r\n\t").lower()
        if low.startswith((b"<", b"mime-version")) or b"<html" in low or b"<table" in low:
            return f"其实是网页格式（很多系统「导出 Excel」导出的就是这种），{save_as}"
        if head.startswith(b"PK"):
            return "其实是 .xlsx 格式却叫 .xls，把扩展名改成 .xlsx 即可"
        if not head.startswith(b"\xd0\xcf\x11\xe0"):
            return f"不是有效的 Excel 文件，{save_as}"
    return ""


def _chain(a, b):
    yield from a
    yield from b


# ---------- 表头与列（读法和列键在 core/xlio.py） ----------

def lock_keys(names, fill_columns):
    """下发时不该被改动的列（列键）：除了让对方填写的列、和没有列名的列以外的所有列。"""
    fill = {hard_clean(c).lower() for c in fill_columns if hard_clean(c)}
    return [k for k in col_keys(names) if not k.startswith("#") and k.split("#")[0] not in fill]


def label(rel, sheet, mode, multi_sheet, who=None):
    """「来源」列写的内容：交表人（认不出时用文件名）；一个表的多个 sheet 合并时写 sheet 名。"""
    stem = who or os.path.splitext(rel)[0].replace(os.sep, "/")
    if mode == "sheets":
        return sheet
    return f"{stem}·{sheet}" if multi_sheet and mode == "all" else stem    # by_name：每张结果表里一个文件只有一个 sheet


def group_key(mode, sheet):
    """合进哪张结果 sheet：all / sheets 模式全合成一张；by_name 模式同名 sheet 合在一起。"""
    return "*" if mode in ("all", "sheets") else soft_clean(sheet)


def file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _key_index(keys, name):
    """列名在列键里的位置：先精确、再包含。找不到返回 None。"""
    k = hard_clean(name).lower()
    if not k:
        return None
    if k in keys:
        return keys.index(k)
    return next((c for c, key in enumerate(keys) if not key.startswith("#") and k in key), None)


def _is_guide(sheet):
    return hard_clean(sheet) == GUIDE_SHEET


def _read_claim(rows):
    """「填报说明」sheet 里「填报对象」那一行写的名字；没有返回 ""。"""
    labels = {hard_clean(w) for w in CLAIM_LABELS}
    for row in islice(rows, 30):
        vals = [soft_clean(v) for v in row if soft_clean(v) != ""]
        for j, v in enumerate(vals):
            m = re.match(r"^(.{2,6}?)\s*[：:]\s*(.+)$", v)
            if m and hard_clean(m.group(1)) in labels:
                return m.group(2).strip()
            if hard_clean(v) in labels and j + 1 < len(vals):
                return vals[j + 1]
    return ""


# ---------- 指纹：和下发时比对用，只存短哈希，不存内容 ----------

def _norm(v):
    """单元格的值 → 比对用的文字：1 与 1.0、日期与日期时间零点视为相同。"""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float) and v == v and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime.datetime):
        return v.date().isoformat() if v.time() == datetime.time(0) else v.isoformat(sep=" ")
    if isinstance(v, datetime.date):
        return v.isoformat()
    return soft_clean(v)


def _digest(text):
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


def fingerprint_file(path, cfg, keys):
    """下发时给一个产出文件记指纹：{sheet 名: {rows, row_fp: [每行指纹], col_fp: {列键: 整列指纹}}}。

    keys 是要比对的列键（lock_keys 的结果）。整列指纹与行的顺序无关，对方排了序不算改动。
    """
    res = scan([(path, os.path.basename(path))], cfg, fp_keys=keys)
    return {s["sheet"]: {"rows": s["rows"], "row_fp": s["fp"]["rows"],
                         "col_fp": {k: v for k, v in s["fp"]["cols"].items() if k in s["keys"]}}
            for s in res["sheets"] if "fp" in s}


# ---------- 扫描：只读，不写任何文件 ----------

def scan(files, cfg, deadline=None, log=None, by=None, fp_keys=None, stop=None, progress=None, rules=None,
         base=None):
    """读各表各 sheet 的表头、行数、每列空了哪几行。超过 deadline 或 stop() 为真就停，如实返回 complete=False。

    files 是 [(绝对路径, 相对路径)]。返回：
      sheets   [{file: 文件序号, rel, sheet, names, keys, rows, totals, filled: {列键: 有内容的行数},
                 blank_rows: {列键: [空着的 Excel 行号]}, groups?: 按 by 列分组的同样统计, fp?: 指纹,
                 bad?: 按 rules（{列名: 规则}）查出的填错 [{col, code, count, rows}]}]
      claims   {文件序号: 「填报说明」里写的填报对象}
      seen     读成功的文件序号；skipped / dupes / errors / empty（只有表头、没填数据的表）
    只留列名、计数和行号，不留任何单元格的值。
    """
    res = {"sheets": [], "skipped": [], "dupes": [], "errors": [], "empty": [], "empty_idx": [], "claims": {},
           "seen": [], "files_scanned": 0, "complete": True}
    hashes = {}
    for i, (path, rel) in enumerate(files):
        if (deadline and time.time() > deadline) or (stop and stop()):
            res["complete"] = False
            break
        if progress:
            progress(i / max(1, len(files)))
        why = unreadable(path)
        if why:
            res["errors"].append({"file": rel, "idx": i, "error": why})
            res["files_scanned"] = i + 1
            continue
        try:
            fh = file_hash(path)
        except OSError as e:
            res["errors"].append({"file": rel, "idx": i, "error": str(e)})
            continue
        if fh in hashes:
            res["dupes"].append({"file": rel, "idx": i, "same_as": hashes[fh]})
            res["files_scanned"] = i + 1
            continue
        hashes[fh] = rel
        hid = hidden_info(path)
        has_data = header_only = False
        try:
            for sn, it in _iter_sheets_ro(path):
                if hid.get(sn, ("visible",))[0] != "visible":
                    res["skipped"].append({"file": rel, "sheet": sn, "reason": "隐藏的 sheet", "hidden": True})
                    continue
                if _is_guide(sn):
                    who = _read_claim(it)
                    if who:
                        res["claims"][i] = who
                    continue
                head = list(islice(it, _head_limit(cfg)))
                h = sp.resolve_header_row(head, cfg)
                if h == -1 or h > len(head):
                    if any(_nonempty(r) for r in head):
                        res["skipped"].append({"file": rel, "sheet": sn, "reason": "认不出表头"})
                    continue
                entry = _scan_sheet(head, h, it, by, fp_keys, rules, base)
                if not entry["rows"]:
                    header_only = True            # 只有表头的 sheet（空模板、填表说明）不合
                    continue
                has_data = True
                entry.update(file=i, rel=rel, sheet=sn)
                res["sheets"].append(entry)
            if header_only and not has_data:
                res["empty"].append(rel)
                res["empty_idx"].append(i)
            res["seen"].append(i)
        except Exception as e:   # 单个表读不了不影响其余表；原因如实带回
            res["errors"].append({"file": rel, "idx": i, "error": str(e)})
        res["files_scanned"] = i + 1
        if log and (i + 1) % 10 == 0:
            log(f"已读 {i + 1}/{len(files)} 个表")
    return res


def _bad_add(store, col, code, r, options=None):
    b = store.get((col, code))
    if b is None:
        b = store[(col, code)] = {"col": col, "code": code, "count": 0, "rows": []}
        if options and len(options) <= SHOW_OPTIONS:
            b["options"] = list(options)
    b["count"] += 1
    if len(b["rows"]) < BLANK_KEEP:
        b["rows"].append(r)


def _scan_sheet(head, h, it, by, fp_keys, rules=None, base=None):
    names = col_names(head, h)
    keys = col_keys(names)
    n = len(names)
    checks = []                       # [(列号, 列名, 规则)]：按下发时的规则查格式
    for name, rule in (rules or {}).items():
        c = _key_index(keys, name)
        if c is not None and rule.get("type", "text") in BAD_CODE:
            checks.append((c, name, rule))
    bad = {}
    rows = totals = 0
    nblank, blank_rows = [0] * n, [[] for _ in range(n)]
    bi = _key_index(keys, by) if by else None
    groups = {} if bi is not None else None
    fpi = [keys.index(k) if k in keys else None for k in fp_keys] if fp_keys is not None else None
    row_fp, row_no, col_acc = [], [], [0] * len(fp_keys or ())
    for r, row in enumerate(_chain(iter(head[h:]), it), start=h + 1):
        if not _nonempty(row):
            continue
        if is_total_row(row):
            totals += 1
            continue
        rows += 1
        empty_cols = [c for c in range(n) if soft_clean(_cell(row, c)) == ""]
        for c in empty_cols:
            nblank[c] += 1
            if len(blank_rows[c]) < BLANK_KEEP:
                blank_rows[c].append(r)
        wrong = []
        for c, name, rule in checks:
            v = _cell(row, c)
            if soft_clean(v) != "" and not _coerce(v, rule, base)[0]:
                wrong.append((name, BAD_CODE[rule["type"]], rule.get("options")))
                _bad_add(bad, *wrong[-1][:2], r, wrong[-1][2])
        if groups is not None:
            g = groups.get(soft_clean(_cell(row, bi)))
            if g is None:
                g = groups[soft_clean(_cell(row, bi))] = {"rows": 0, "nblank": [0] * n, "blank_rows": [[] for _ in range(n)],
                                                          "bad": {}}
            g["rows"] += 1
            for c in empty_cols:
                g["nblank"][c] += 1
                if len(g["blank_rows"][c]) < BLANK_KEEP:
                    g["blank_rows"][c].append(r)
            for name, code, opts in wrong:
                _bad_add(g["bad"], name, code, r, opts)
        if fpi is not None:
            vals = [_norm(_cell(row, c)) for c in fpi]
            row_fp.append(_digest("\x1f".join(vals))[:12])
            row_no.append(r)
            for j, v in enumerate(vals):
                col_acc[j] = (col_acc[j] + int(_digest(v)[:16], 16)) & _MASK
    entry = {"names": names, "keys": keys, "rows": rows, "totals": totals,
             "filled": {k: rows - nb for k, nb in zip(keys, nblank)},
             "blank_rows": {k: br for k, br in zip(keys, blank_rows) if br}}
    if groups is not None:
        entry["by"] = names[bi]
        entry["groups"] = {gv: {"rows": g["rows"], "filled": {k: g["rows"] - nb for k, nb in zip(keys, g["nblank"])},
                                "blank_rows": {k: br for k, br in zip(keys, g["blank_rows"]) if br},
                                "bad": list(g["bad"].values())}
                           for gv, g in groups.items()}
    if checks:
        entry["bad"] = list(bad.values())
    if fpi is not None:
        entry["fp"] = {"rows": row_fp, "row_no": row_no,
                       "cols": {k: f"{a:016x}" for k, a in zip(fp_keys, col_acc)}}
    return entry


def pick_mode(n_files, sheets):
    """sheets：只有一个表时把它的各 sheet 合成一张；all：每个表只有一个有数据的 sheet，全合成一张；
    by_name：表里有多个 sheet，同名 sheet 合在一起。"""
    if n_files == 1:
        return "sheets"
    per = {}
    for s in sheets:
        per[s["file"]] = per.get(s["file"], 0) + 1
    return "all" if all(n == 1 for n in per.values()) else "by_name"


def compare(sheets, mode):
    """按结果 sheet 分组，以每组第一个 sheet 的表头为准，算出每个来源缺的列、多的列、是否对不上。

    返回 {组键: {"title", "template": 来源, "rows", "items": [{file, rel, sheet, rows, totals, missing, extra,
    mismatch, filled, blank_rows, blanks, fp, group}]}}。
    """
    groups = {}
    for s in sheets:
        g = groups.setdefault(group_key(mode, s["sheet"]), {"title": s["sheet"], "template": s, "rows": 0,
                                                             "keys": list(s["keys"]), "items": []})
        tk = g["template"]["keys"]
        hit, bad = match(tk, s["keys"])
        named = dict(zip(s["keys"], s["names"]))
        tnamed = dict(zip(tk, g["template"]["names"]))
        missing = [tnamed[k] for k in tk if not k.startswith("#") and k not in named] if not bad else []
        extra = [named[k] for k in s["keys"] if not k.startswith("#") and k not in tnamed] if not bad else []
        g["items"].append({"file": s["file"], "rel": s["rel"], "sheet": s["sheet"], "rows": s["rows"],
                           "totals": s["totals"], "missing": missing, "extra": extra, "mismatch": bad,
                           "filled": s["filled"], "blank_rows": s.get("blank_rows", {}), "blanks": {},
                           "fp": s.get("fp"), "group": s.get("group"), "bad": s.get("bad", [])})
        if not bad:
            g["rows"] += s["rows"]
    return groups


def completeness(groups, wanted=()):
    """定出每组该填的列（g["required"]：列键），算出每个来源空着几格（item["blanks"]：{列名: 空格数}）。

    wanted 是用户说的必填列名；返回其中在表头里找不到的。
    """
    unknown = [w for w in wanted if not any(_find_key(g, w) for g in groups.values())]
    for g in groups.values():
        tk, tn = g["template"]["keys"], g["template"]["names"]
        items = [i for i in g["items"] if not i["mismatch"]]
        if wanted:
            req = [k for k in (_find_key(g, w) for w in wanted) if k]
        else:
            req = []
            for k, n in zip(tk, tn):
                if k.startswith("#") or any(w in n for w in OPTIONAL_WORDS):
                    continue
                have = [i for i in items if k in i["filled"]]
                rows = sum(i["rows"] for i in have)
                full = sum(1 for i in have if i["filled"][k] >= i["rows"])
                if rows and full * 2 >= len(have) and sum(i["filled"][k] for i in have) >= FILLED_SHARE * rows:
                    req.append(k)
        g["required"] = list(dict.fromkeys(req))
        names = dict(zip(tk, tn))
        for i in items:
            i["blanks"] = {names[k]: i["rows"] - i["filled"][k] for k in g["required"]
                           if k in i["filled"] and i["filled"][k] < i["rows"]}
    return unknown


def _find_key(group, name):
    """用户说的列名对应模板的列键：先精确，再包含。找不到返回 None。"""
    k = hard_clean(name).lower()
    tk = group["template"]["keys"]
    if not k:
        return None
    if k in tk:
        return k
    return next((key for key in tk if not key.startswith("#") and (k in key or key.split("#")[0] in k)), None)


def expect_check(expect, rels):
    """应报名单与文件名对照：返回 (未报的名字, 报了多份的 {名字: [文件]}, 名单外的文件)。

    一个文件只归一个名字，长名字优先：「城区二.xlsx」归「城区二」，不会同时算成「城区」交了。
    """
    keyed = _keyed(expect)
    owner = {}
    for r in rels:
        stem = hard_clean(os.path.splitext(os.path.basename(r))[0]).lower()
        name = next((n for k, n in keyed if k in stem), None)
        if name is not None:
            owner[r] = name
    hit = {n: [r for r in rels if owner.get(r) == n] for _k, n in keyed}
    order = [n for n in expect if n in hit]
    missing = [n for n in order if not hit[n]]
    multi = {n: hit[n] for n in order if len(hit[n]) > 1}
    return missing, multi, [r for r in rels if r not in owner]


# ---------- 名单、认领 ----------

def split_names(text):
    """「张三、李四，王五」/ 每行一个 → [名字]，去重保序。"""
    if isinstance(text, (list, tuple)):
        parts = [soft_clean(t) for t in text]
    else:
        parts = [soft_clean(t) for t in re.split(r"[、,，;；\s]+", text or "")]
    return list(dict.fromkeys(p for p in parts if p))


# ---------- 填报规则：「让对方填写的列」括号里写的要求 ----------

RULE_TYPES = {"日期": "date", "数字": "number", "整数": "whole", "手机号": "mobile", "手机": "mobile",
              "手机号码": "mobile", "电话": "phone"}
RULE_LABELS = {"list": "从下拉里选", "date": "日期", "number": "数字", "whole": "整数", "mobile": "手机号",
               "phone": "电话号码"}
RULE_EXAMPLES = {"date": "2026-10-15", "number": "3.5", "whole": "3", "mobile": "13800138000",
                 "phone": "13800138000 或 010-12345678"}
BAD_CODE = {"list": "not_in_options", "date": "not_date", "number": "not_number", "whole": "not_whole",
            "mobile": "not_mobile", "phone": "not_phone"}
LIST_MAX = 255               # Excel 下拉框选项文字的上限
SHOW_OPTIONS = 6             # 通报里最多列出这么多个选项，再多只说「请从下拉里选」
_NUM_LIMIT = "1E+15"


def parse_fill_spec(text, log=None):
    """「走访日期（日期）、走访结果（已走访/未走访）、备注（选填）」→ [(列名, 规则)]。

    规则：{"type": text|list|date|number|whole|mobile|phone, "options": [...]（仅 list）, "optional": bool}。
    括号外按 split_names 的分隔符切；不写括号时和 split_names 完全一样（每列都是 text、必填）。
    括号里看不懂的词不拦人：这一列按普通文字处理（只查漏填），日志里说一句。
    """
    log = log or (lambda m: None)
    if isinstance(text, (list, tuple)):
        text = "、".join(str(t) for t in text if t is not None)
    text = text or ""
    if "(" not in text and "（" not in text:
        return [(n, {"type": "text", "optional": False}) for n in split_names(text)]
    tokens, buf, depth = [], "", 0
    for ch in text:
        if ch in "(（":
            depth += 1
        elif ch in ")）":
            depth = max(0, depth - 1)
        if depth == 0 and re.match(r"[、,，;；\s]", ch):
            tokens.append(buf)
            buf = ""
        else:
            buf += ch
    tokens.append(buf)
    out, seen = [], set()
    for tok in tokens:
        m = re.match(r"^(.*?)[(（](.*)[)）]\s*$", tok.strip())
        name = soft_clean(m.group(1) if m else tok)
        if not name or name in seen:
            continue
        seen.add(name)
        rule = {"type": "text", "optional": False}
        typ, options, unknown = None, None, []
        for w in re.split(r"[\s,，、;；]+", m.group(2) if m else ""):
            w = w.strip()
            if not w:
                continue
            if re.search(r"[/｜|]", w):
                options = list(dict.fromkeys(o.strip() for o in re.split(r"[/｜|]", w) if o.strip()))
            elif w in RULE_TYPES:
                typ = RULE_TYPES[w]
            elif w == "选填":
                rule["optional"] = True
            elif w != "必填":
                unknown.append(w)
        if options:
            if typ:
                log(f"  ℹ️ 「{name}」同时写了选项和类型，按选项处理")
            rule.update(type="list", options=options)
        elif typ:
            rule["type"] = typ
        if unknown:
            hint = "；日期请写（日期），选项用 / 隔开" if rule["type"] == "text" else ""
            log(f"  ℹ️ 看不懂「{'、'.join(unknown)}」，「{name}」" + ("按普通文字处理" if rule["type"] == "text"
                                                                 else "其余要求照常") + hint)
        out.append((name, rule))
    return out


def _nfkc(v):
    return unicodedata.normalize("NFKC", str(v)).strip()


_DATE_FULL = re.compile(r"^(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$")
_DATE_COMPACT = re.compile(r"^(\d{4})(\d{2})(\d{2})$")
_DATE_SHORT = re.compile(r"^(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?$")
_NUMBER = re.compile(r"^[+-]?(\d+(\.\d*)?|\.\d+)([\u4e00-\u9fff]{1,2})?$")
_SERIAL_MIN, _SERIAL_MAX = 32874, 73415          # 1990-01-01 ～ 2100-12-31 的 Excel 日期序号


def _mk_date(y, m, d):
    try:
        return datetime.datetime(int(y), int(m), int(d))
    except ValueError:
        return None


def _near_year(m, d, base):
    """只写了月、日：年份取离下发日期最近的那一年（12 月下发、填「1月3日」算次年）。"""
    best = None
    for y in (base.year - 1, base.year, base.year + 1):
        dt = _mk_date(y, m, d)
        if dt and (best is None or abs((dt.date() - base).days) < abs((best.date() - base).days)):
            best = dt
    return best


def _as_number(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    t = _nfkc(v).replace(",", "").replace(" ", "")
    m = _NUMBER.match(t)
    if not m:
        return None
    n = float(m.group(1)) if "." in m.group(1) else int(m.group(1))
    return -n if t.startswith("-") else n


def _mobile(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    t = re.sub(r"[\s-]", "", _nfkc(v))
    for p in ("+86", "0086"):
        if t.startswith(p):
            t = t[len(p):]
    if t.startswith("86") and len(t) == 13:
        t = t[2:]
    return t if re.fullmatch(r"1[3-9]\d{9}", t) else None


def _landline(v):
    if not isinstance(v, str):
        return None
    t = _nfkc(v)
    m = re.match(r"^(.*?)((?:转|ext\.?|#)\s*\d{1,6})?$", t, re.I)
    body, ext = m.group(1).strip(), (m.group(2) or "")
    m2 = re.fullmatch(r"[(（]?(0\d{2,3})[)）]?[\s-]+(\d{7,8})|[(（](0\d{2,3})[)）](\d{7,8})", body)
    if m2:
        area, local = (m2.group(1), m2.group(2)) if m2.group(1) else (m2.group(3), m2.group(4))
    else:
        digits = body.replace(" ", "")
        if not re.fullmatch(r"0\d{9,11}", digits):
            return None
        n = 3 if digits[1] in "12" else 4           # 010、02x 是三位区号，其余四位
        area, local = digits[:n], digits[n:]
        if not 7 <= len(local) <= 8:
            return None
    return f"{area}-{local}{ext}"


def _coerce(value, rule, base=None):
    """按规则判一个非空格子：返回 (对不对, 规范写法)。查收判对错、汇总时改写法共用这一个函数，口径不会对不上。

    宽进：能无歧义认出来的算对（文字日期、「1,200」「3户」、带空格 / +86 的手机号、全角选项）；
    数字形式的「10.1」分不清是 10月1日 还是 10月10日，不算日期。base：下发日期（补年份用），默认今天。
    """
    typ = rule.get("type", "text")
    if typ == "text":
        return True, value
    if typ == "list":
        t = _nfkc(value).casefold()
        for o in rule.get("options", ()):
            if _nfkc(o).casefold() == t:
                return True, o
        return False, value
    if typ in ("number", "whole"):
        n = _as_number(value)
        if n is None:
            return False, value
        if isinstance(n, float) and n.is_integer():
            n = int(n)
        if typ == "whole" and not isinstance(n, int):
            return False, value
        return True, n
    if typ == "date":
        if isinstance(value, datetime.date):
            # 设成日期格式的格子里填了「10.1」这类数，读出来是 1900 年的日期：和序号一样只认 1990～2100 年
            if not 1990 <= value.year <= 2100:
                return False, value
            if isinstance(value, datetime.datetime):
                return True, value
            return True, datetime.datetime(value.year, value.month, value.day)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            if _SERIAL_MIN <= value <= _SERIAL_MAX:
                return True, datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(value))
            return False, value
        t = _nfkc(value)
        base = base or datetime.date.today()
        m = _DATE_FULL.match(t) or _DATE_COMPACT.match(t)
        dt = _mk_date(*m.groups()[:3]) if m else None
        if dt is None and not m:
            m = _DATE_SHORT.match(t)
            dt = _near_year(m.group(1), m.group(2), base) if m else None
        return (True, dt) if dt else (False, value)
    if typ == "mobile":
        mob = _mobile(value)
        return (True, mob) if mob else (False, value)
    if typ == "phone":
        tel = _mobile(value) or _landline(value)
        return (True, tel) if tel else (False, value)
    return True, value


def _mask(text):
    """号码打码：只留前 3 位和后 4 位数字（「汇总说明」被截图转发时不带出号码）。"""
    t = str(text)
    pos = [i for i, ch in enumerate(t) if ch.isdigit()]
    hide = set(pos[3:-4]) if len(pos) > 7 else set()
    return "".join("*" if i in hide else ch for i, ch in enumerate(t))


def _rule_text(rule):
    """填报说明里「要求」一栏。"""
    head = "选填" if rule.get("optional") else "必填"
    return f"{head} · {RULE_LABELS[rule['type']]}" if rule["type"] in RULE_LABELS else head


def _validation(rule, name, top, letter):
    """一列的数据验证；不能加（选项太长、含英文逗号、电话）时返回 None。"""
    typ = rule["type"]
    kw = {"allow_blank": True, "showErrorMessage": True, "errorStyle": "stop", "showInputMessage": True,
          "errorTitle": name[:32], "promptTitle": name[:32]}
    if typ == "list":
        opts = rule["options"]
        lit = ",".join(opts)
        if len(lit) > LIST_MAX or any("," in o or '"' in o for o in opts):
            return None
        tip = " / ".join(opts)
        return DataValidation(type="list", formula1=f'"{lit}"', error=f"请从下拉里选：{tip}"[:255],
                              prompt=f"从下拉里选：{tip}"[:255], **kw)
    if typ == "date":
        return DataValidation(type="date", operator="greaterThan", formula1="1",
                              error="请填日期，比如 2026-10-15", prompt="填日期，比如 2026-10-15", **kw)
    if typ in ("number", "whole"):
        what = "数字" if typ == "number" else "整数"
        return DataValidation(type="decimal" if typ == "number" else "whole", operator="between",
                              formula1=f"-{_NUM_LIMIT}", formula2=_NUM_LIMIT,
                              error=f"请填{what}，不要带单位或文字", prompt=f"填{what}", **kw)
    if typ == "mobile":
        c = f"{letter}{top}"
        return DataValidation(type="custom", formula1=f'AND(LEN({c})=11,LEFT({c},1)="1",ISNUMBER(--{c}))',
                              error="请填 11 位手机号", prompt="填 11 位手机号", **kw)
    return None


def _keyed(names):
    """[(比对用的键, 名字)]，长名字排在前面。"""
    pairs = [(hard_clean(n).lower(), n) for n in names if hard_clean(n)]
    return sorted(pairs, key=lambda p: -len(p[0]))


def read_name_list(path):
    """从「应交名单」「发送清单」这类表里读名字。

    表头里有「应交 / 填报对象 / 分组 / 名称 / 单位 / 网格 / 姓名 / 接收人」的列优先（按这个顺序），
    认不出表头时取第一列。合计行、空格子不算。
    """
    for _sn, it in _iter_sheets_ro(path):
        rows = [r for r in it]
        if not any(_nonempty(r) for r in rows):
            continue
        col, start = None, 0
        for word in NAME_HEADS:
            for r, row in enumerate(rows[:5]):
                c = next((c for c, v in enumerate(row) if word in hard_clean(v) and len(hard_clean(v)) <= 10), None)
                if c is not None:
                    col, start = c, r + 1
                    break
            if col is not None:
                break
        if col is None:
            col = next(c for c in range(max(len(r) for r in rows)) if any(soft_clean(_cell(r, c)) for r in rows))
        names = [soft_clean(_cell(r, col)) for r in rows[start:] if not is_total_row(r)]
        names = list(dict.fromkeys(n for n in names if n))
        if names:
            return names
    return []


def find_record(input_path):
    """在回收文件夹和它的上一层找「下发记录.json」，找不到返回 ""。"""
    base = input_path if os.path.isdir(input_path) else os.path.dirname(input_path)
    for d in (base, os.path.dirname(base)):
        p = os.path.join(d, RECORD_NAME)
        if os.path.isfile(p):
            return p
    return ""


def load_record(path):
    """读下发记录；格式不对返回 None。"""
    try:
        with open(path, encoding="utf-8") as f:
            rec = json.load(f)
    except (OSError, ValueError):
        return None
    if not isinstance(rec, dict) or not isinstance(rec.get("groups"), dict):
        return None
    return rec


_COPY_RE = re.compile(r"(\s*[-_]?\s*副本)$|(\s*[(（]\d+[)）])$|([_-]v?\d{1,2})$", re.I)


def _base_stem(rel):
    """去掉「副本」「(2)」「_v2」这类后缀：同一个人交了几次同一份表时，只查最新的那份。"""
    s = os.path.splitext(os.path.basename(rel))[0]
    prev = None
    while prev != s:
        prev, s = s, _COPY_RE.sub("", s).strip()
    return s or os.path.splitext(os.path.basename(rel))[0]


def claim_files(files, res, expect=(), record=None):
    """每个文件是谁交的：返回 (owners {文件序号: 名字}, extra [名单外的文件序号])。

    顺序：「填报说明」里写的填报对象 → 文件名 → 所在文件夹名 → 和下发时哪一份的行对得上
    → 分组列里占多数的取值。没有名单时用文件名。
    """
    keyed = _keyed(expect)
    exact = {k: n for k, n in keyed}
    dominant = _dominant_groups(res)
    by_rows = _fp_owners(res, record) if keyed and record else {}
    candidates = sorted(set(res["seen"]) | {e["idx"] for e in res["errors"] if "idx" in e}
                        | {d["idx"] for d in res["dupes"] if "idx" in d})
    owners, extra = {}, []
    for i in candidates:
        rel = files[i][1]
        stem = os.path.splitext(os.path.basename(rel))[0]
        claim = res["claims"].get(i, "")
        name = (exact.get(hard_clean(claim).lower()) if keyed else claim) or None
        if name is None and keyed:
            for target in (stem, os.path.dirname(rel)):
                t = hard_clean(target).lower()
                name = next((n for k, n in keyed if t and k in t), None)
                if name:
                    break
        if name is None and by_rows.get(i):
            name = exact.get(hard_clean(by_rows[i]).lower())
        if name is None and dominant.get(i):
            v = dominant[i]
            name = exact.get(hard_clean(v).lower()) if keyed else v
        if name is None and not keyed:
            name = _base_stem(rel)
        if name is None:
            extra.append(i)
        else:
            owners[i] = name
    return owners, extra


def _fp_owners(res, record):
    """{文件序号: 下发记录里的分组名}：按行指纹找这份表是从哪一份下发的。

    给删了「填报说明」、又改了文件名（或把数据贴进新表）的文件兜底。要求一半以上的行对得上，
    且明显多于第二名，宁可认不出也不认错。
    """
    issued = {}
    for g, info in record.get("groups", {}).items():
        c = Counter()
        for s in info.get("sheets", {}).values():
            c.update(s.get("row_fp", []))
        if c:
            issued[g] = c
    if not issued:
        return {}
    per = {}
    for s in res["sheets"]:
        if s.get("fp"):
            per.setdefault(s["file"], Counter()).update(s["fp"]["rows"])
    out = {}
    for i, mine in per.items():
        total = sum(mine.values())
        hits = sorted(((sum((mine & c).values()), g) for g, c in issued.items()), reverse=True)
        best, g = hits[0]
        second = hits[1][0] if len(hits) > 1 else 0
        if best * 2 >= total and best > second:
            out[i] = g
    return out


def _dominant_groups(res):
    """{文件序号: 分组列里占一半以上的取值}（扫描时指定了 by 才有）。"""
    per = {}
    for s in res["sheets"]:
        for gv, g in s.get("groups", {}).items():
            c = per.setdefault(s["file"], Counter())
            c[gv] += g["rows"]
    out = {}
    for i, c in per.items():
        gv, n = c.most_common(1)[0]
        if gv and n * 2 > sum(c.values()):
            out[i] = gv
    return out


# ---------- 查收：逐人判定 ----------

def _rows_text(rows, total=None):
    total = total if total is not None else len(rows)
    shown = "、".join(str(r) for r in rows[:SHOW_ROWS])
    return f"第 {shown} 行" if total <= SHOW_ROWS and len(rows) >= total else f"第 {shown} 等 {total} 行"


def _blank_text(b):
    if b.get("no_col"):
        return f"没有「{b['col']}」这一列"
    return f"{b['col']} 漏 {b['count']} 行（{_rows_text(b['rows'], b['count'])}）"


# 问题编码：每个人的问题先记成 {"code": 编码, ...}（只有列名、文件名、行号、个数，不含单元格的值），
# 状态、通报文字、旧字段 blanks / changes / notes 都由它推出来。编码定下来以后只加不改——
# 回收结果.json、回收记录.json 和技能都按编码认问题。
ISSUE_STATUS = {
    "missing": ST_MISSING, "not_in_total": ST_MISSING,
    "empty": ST_EMPTY,
    "blank": ST_PARTIAL, "no_col": ST_PARTIAL,
    "not_in_options": ST_INVALID, "not_date": ST_INVALID, "not_number": ST_INVALID, "not_whole": ST_INVALID,
    "not_mobile": ST_INVALID, "not_phone": ST_INVALID,
    "col_deleted": ST_CHANGED, "rows_lost": ST_CHANGED, "rows_added": ST_CHANGED, "rows_changed": ST_CHANGED,
    "header_mismatch": ST_BAD, "unreadable": ST_BAD, "no_header": ST_BAD,
    "old_version": None, "same_as": None,          # 只注明，不影响状态
}
ISSUE_NAG = {c: st is not None for c, st in ISSUE_STATUS.items()}    # 要不要催（只注明的不催）
_BLANK_CODES = ("blank", "no_col")
_INVALID_CODES = ("not_in_options", "not_date", "not_number", "not_whole", "not_mobile", "not_phone")
INVALID_TEXT = {"not_in_options": "不在选项里", "not_date": "不是日期", "not_number": "不是数字",
                "not_whole": "不是整数", "not_mobile": "不是手机号", "not_phone": "不是电话号码"}
_CHANGE_CODES = ("col_deleted", "rows_lost", "rows_added", "rows_changed")
_NOTE_CODES = ("not_in_total", "old_version", "same_as", "unreadable", "no_header", "header_mismatch")


def _quoted(cols):
    return "、".join(f"「{c}」" for c in cols)


ISSUE_TEXT = {
    "missing": lambda i: "还没收到你的表",
    "not_in_total": lambda i: "表里没有这个名字的行",
    "empty": lambda i: "只有表头，没填数据" if i.get("why") == "header_only" else "该填的列一格都没填",
    "blank": lambda i: _blank_text(i),
    "no_col": lambda i: _blank_text(dict(i, no_col=True)),
    **{c: (lambda i: _invalid_text(i)) for c in ("not_in_options", "not_date", "not_number", "not_whole",
                                                 "not_mobile", "not_phone")},
    "col_deleted": lambda i: "删了" + _quoted(i["cols"]) + "列",
    "rows_lost": lambda i: f"少了 {i['count']} 行",
    "rows_added": lambda i: f"多了 {i['count']} 行",
    "rows_changed": lambda i: f"{_rows_text(i['rows'], i['count'])}和下发时不一样"
                              + (f"（改了{_quoted(i['cols'])}）" if i.get("cols") else ""),
    "header_mismatch": lambda i: f"「{i['sheet']}」表头和别人的对不上，没法查",
    "unreadable": lambda i: f"「{i['file']}」打不开：{i['why']}",
    "no_header": lambda i: f"「{i['file']}」认不出表头",
    "old_version": lambda i: f"「{i['file']}」是同一份表的旧版本，按最新的「{i['newest']}」查",
    "same_as": lambda i: f"「{i['file']}」和「{i['other']}」内容完全一样，请核对是否交错了表",
}


def _invalid_text(i):
    text = f"{i['col']} {_rows_text(i.get('rows', []), i['count'])}{INVALID_TEXT[i['code']]}"
    if i["code"] == "not_in_options":
        text += f"（应为 {' / '.join(i['options'])}）" if i.get("options") else "（请从下拉里选）"
    return text


def issue_text(i):
    """一个问题的说明文字。多 sheet 时漏填、改动类问题前面带「sheet·」。"""
    fn = ISSUE_TEXT.get(i["code"])
    text = fn(i) if fn else i["code"]
    if i["code"] in _BLANK_CODES + _INVALID_CODES + _CHANGE_CODES and i.get("sheet"):
        if i["code"] in _BLANK_CODES + _INVALID_CODES:
            return fn(dict(i, col=f"{i['sheet']}·{i['col']}"))
        return f"{i['sheet']}·{text}"
    return text


def issue_counts(issues):
    """{编码: 个数}：漏填、行数类按格数 / 行数算，其余每条算 1。回收记录.json 每轮按人存这个。"""
    out = {}
    for i in issues:
        out[i["code"]] = out.get(i["code"], 0) + (i.get("count") or 1)
    return out


def _derive(p):
    """由 p["issues"] 推出状态和旧字段 blanks / changes / notes（内部技能和界面还在读旧字段）。"""
    issues = p["issues"]
    p["blanks"] = [{"col": (f"{i['sheet']}·" if i.get("sheet") else "") + i["col"], "count": i["count"],
                    "rows": i.get("rows", []), **({"no_col": True} if i["code"] == "no_col" else {})}
                   for i in issues if i["code"] in _BLANK_CODES]
    p["invalid"] = [{"col": (f"{i['sheet']}·" if i.get("sheet") else "") + i["col"], "code": i["code"],
                     "count": i["count"], "rows": i.get("rows", [])} for i in issues if i["code"] in _INVALID_CODES]
    p["changes"] = [issue_text(i) for i in issues if i["code"] in _CHANGE_CODES]
    notes = [issue_text(i) for i in issues if i["code"] in _NOTE_CODES]
    if not notes and any(i["code"] == "empty" and i.get("why") == "header_only" for i in issues):
        notes = [issue_text({"code": "empty", "why": "header_only"})]
    p["notes"] = notes
    # 打不开 / 认不出表头的那份：这个人别的表能查时只注明，全都查不了才算「表头对不上（没法查）」
    flags = {ISSUE_STATUS.get(i["code"]) for i in issues
             if not (i["code"] in ("unreadable", "no_header") and p["rows"])} - {None, ST_MISSING}
    p["flags"] = sorted(flags, key=_RANK.get)
    p["status"] = p["flags"][0] if p["flags"] else ST_OK


def _diff(fp, base, names, lock_names):
    """和下发时的指纹比：返回问题列表（col_deleted / rows_lost / rows_added / rows_changed）。"""
    out = []
    pool = Counter(base.get("row_fp", []))
    new_rows = []
    for h, r in zip(fp["rows"], fp["row_no"]):
        if pool[h] > 0:
            pool[h] -= 1
        else:
            new_rows.append(r)
    lost = sum(pool.values())
    gone = [lock_names.get(k, k) for k in base.get("col_fp", {}) if k not in names]
    if gone:
        out.append({"code": "col_deleted", "cols": gone})
    if lost > len(new_rows):
        out.append({"code": "rows_lost", "count": lost - len(new_rows)})
    elif len(new_rows) > lost:
        out.append({"code": "rows_added", "count": len(new_rows) - lost})
    if new_rows and lost and not gone:
        cols = []
        if len(fp["rows"]) == base.get("rows"):
            cols = [lock_names.get(k, k) for k, v in base.get("col_fp", {}).items() if fp["cols"].get(k) != v]
        out.append({"code": "rows_changed", "count": len(new_rows), "rows": new_rows[:BLANK_KEEP], "cols": cols})
    return out


def analyze(files, res, expect=(), wanted=(), record=None, by=None, rows_mode=False):
    """逐人判定交表情况，返回报告 dict：

      people  [{name, status, flags, files, rows, issues: [{code, ...}], blanks: [{col, count, rows, no_col?}],
                changes, notes, time}]——issues 见 ISSUE_STATUS；blanks / changes / notes 由 issues 推出来
      extra   名单外的文件（或总表里名单外的分组）；dupes / errors；required 该填的列（列名）；
      unknown_required 指定了却在表头里找不到的列；mode files|rows
    """
    record = record or {}
    sheets = res["sheets"]
    if rows_mode:
        sheets = [dict(s, rows=g["rows"], filled=g["filled"], blank_rows=g["blank_rows"], group=gv, fp=None,
                       bad=g.get("bad", []))
                  for s in res["sheets"] for gv, g in s.get("groups", {}).items()]
        mode = "by_name"
    else:
        mode = pick_mode(len(files), sheets)
    groups = compare(sheets, mode)
    unknown = completeness(groups, wanted)
    required = []
    for g in groups.values():
        tn = dict(zip(g["template"]["keys"], g["template"]["names"]))
        required += [tn[k] for k in g["required"]]
    report = {"people": [], "extra": [], "dupes": [d["file"] for d in res["dupes"]],
              "errors": [f"{e['file']}：{e['error']}" for e in res["errors"]],
              "required": list(dict.fromkeys(required)), "unknown_required": unknown,
              "mode": "rows" if rows_mode else "files", "notes": [], "has_list": bool(expect)}
    by_owner = {}                    # 名字 → [(item, group)]
    if rows_mode:
        exact = {k: n for k, n in _keyed(expect)}
        for g in groups.values():
            for it in g["items"]:
                gv = it["group"]
                if not gv:
                    report["notes"].append(f"「{it['sheet']}」有 {it['rows']} 行没写{by or '分组'}")
                    continue
                name = exact.get(hard_clean(gv).lower()) if expect else gv
                if name is None:
                    if gv not in report["extra"]:
                        report["extra"].append(gv)
                    continue
                by_owner.setdefault(name, []).append((it, g))
        owners_files, dupe_of = {}, {}
    else:
        owners, extra = claim_files(files, res, expect, record)
        # 内容完全相同的文件：同一个人交了两次就只算一份；别人交了一模一样的表，照样算他交了，并注明
        idx_of = {rel: i for i, (_p, rel) in enumerate(files)}
        dupe_of = {d["idx"]: idx_of.get(d["same_as"]) for d in res["dupes"] if "idx" in d}
        report["dupes"] = []
        for d, o in dupe_of.items():
            if d in owners and owners.get(o) != owners[d]:
                continue
            owners.pop(d, None)
            report["dupes"].append(files[d][1])
        report["extra"] = [files[i][1] for i in extra if i not in dupe_of]
        owners_files = {}
        for i, name in owners.items():
            owners_files.setdefault(name, []).append(i)
        items_by_file = {}
        for g in groups.values():
            for it in g["items"]:
                items_by_file.setdefault(it["file"], []).append((it, g))
        for d, o in dupe_of.items():
            if d in owners and o is not None:
                items_by_file[d] = items_by_file.get(o, [])
    names = list(expect) or list(dict.fromkeys(list(owners_files) + list(by_owner)))
    lock_names = dict(zip(record.get("lock_keys", []), record.get("lock_columns", [])))
    rec_groups = record.get("groups", {})
    errs = {e["idx"]: e["error"] for e in res["errors"] if "idx" in e}
    for name in names:
        p = {"name": name, "status": ST_MISSING, "flags": [], "files": [], "rows": 0, "issues": [], "blanks": [],
             "changes": [], "notes": [], "time": ""}
        if rows_mode:
            pairs = by_owner.get(name, [])
            if not pairs:
                p["issues"].append({"code": "not_in_total"})
        else:
            idxs = owners_files.get(name, [])
            chosen = _newest(files, idxs, p["issues"])
            p["files"] = [files[i][1] for i in chosen]
            p["time"] = _mtime(files, chosen)
            pairs = [pr for i in chosen for pr in items_by_file.get(i, [])]
            for i in chosen:
                o = dupe_of.get(i, i)
                if o != i and o is not None:
                    p["issues"].append({"code": "same_as", "file": files[i][1], "other": files[o][1]})
                if o in errs:
                    p["issues"].append({"code": "unreadable", "file": files[i][1], "why": errs[o]})
                elif not items_by_file.get(i) and o not in res["empty_idx"]:
                    p["issues"].append({"code": "no_header", "file": files[i][1]})
        _judge(p, pairs, chosen_count=len(p["files"]) or len(pairs), rec=rec_groups.get(name),
               lock_names=lock_names)
        report["people"].append(p)
    return report


def _newest(files, idxs, issues):
    """同一个人交了几次同一份表（「副本」「(2)」）只查最新的；不同的表都查。旧版本记进 issues（old_version）。"""
    by_base = {}
    for i in idxs:
        by_base.setdefault(_base_stem(files[i][1]), []).append(i)
    chosen = []
    for group in by_base.values():
        group.sort(key=lambda i: _safe_mtime(files[i][0]))
        chosen.append(group[-1])
        for i in group[:-1]:
            issues.append({"code": "old_version", "file": files[i][1], "newest": files[group[-1]][1]})
    return sorted(chosen)


def _safe_mtime(path):
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0.0


def _mtime(files, idxs):
    ts = [_safe_mtime(files[i][0]) for i in idxs]
    return datetime.datetime.fromtimestamp(max(ts)).strftime("%m-%d %H:%M") if ts and max(ts) else ""


def _judge(p, pairs, chosen_count, rec, lock_names):
    """根据这个人的各个 sheet 记下问题（p["issues"]），再推出状态：空表 / 有漏填 / 有改动 / 表头对不上 / 合格。"""
    issues = p["issues"]
    if not chosen_count:
        if not any(i["code"] == "not_in_total" for i in issues):
            issues.insert(0, {"code": "missing"})
        _derive(p)
        p["status"] = ST_MISSING
        return
    good = [(it, g) for it, g in pairs if not it["mismatch"]]
    multi = len(good) > 1
    filled_any = False
    for it, _g in pairs:
        if it["mismatch"]:
            issues.append({"code": "header_mismatch", "sheet": it["sheet"]})
    found = []                                   # 漏填、改动类问题（按 sheet 顺序）
    for it, g in good:
        p["rows"] += it["rows"]
        tn = dict(zip(g["template"]["keys"], g["template"]["names"]))
        req = g.get("required", [])
        if not req or any(it["filled"].get(k, 0) for k in req):
            filled_any = True
        where = {"sheet": it["sheet"]} if multi else {}
        for k in req:
            if k not in it["filled"]:
                found.append({"code": "no_col", "col": tn[k], "count": it["rows"], **where})
            elif it["filled"][k] < it["rows"]:
                found.append({"code": "blank", "col": tn[k], "count": it["rows"] - it["filled"][k],
                              "rows": it["blank_rows"].get(k, []), **where})
        found += [dict(b, **where) for b in it.get("bad", ())]
        if rec and it.get("fp"):
            sheets = rec.get("sheets", {})
            base = sheets.get(it["sheet"]) or (next(iter(sheets.values())) if len(sheets) == 1 else None)
            if base:
                found += [dict(d, **where) for d in _diff(it["fp"], base, set(it["filled"]), lock_names)]
    has_bad = any(i["code"] == "header_mismatch" for i in issues)
    if not good and not has_bad:
        # 只有表头 / 认不出表头 / 打不开：打不开、认不出表头的算「表头对不上（没法查）」，不再记空表
        if not any(i["code"] in ("unreadable", "no_header") for i in issues):
            issues.append({"code": "empty", "why": "header_only" if not pairs else "no_fill"})
    elif good and not filled_any:
        issues.append({"code": "empty", "why": "no_fill"})
        found = [i for i in found if i["code"] not in _BLANK_CODES]      # 一格没填，不必逐列列出
    issues += found
    _derive(p)


# ---------- 通报、进度表 ----------

def _now_text(now):
    return f"{now.month}月{now.day}日 {now:%H:%M}"


def counts(report):
    people = report["people"]
    submitted = [p for p in people if p["status"] != ST_MISSING]
    return {"expected": len(people) if report.get("has_list") else None, "submitted": len(submitted),
            "ok": sum(1 for p in people if p["status"] == ST_OK),
            "missing": sum(1 for p in people if p["status"] == ST_MISSING),
            "empty": sum(1 for p in people if ST_EMPTY in p["flags"]),
            "partial": sum(1 for p in people if ST_PARTIAL in p["flags"]),
            "invalid": sum(1 for p in people if ST_INVALID in p["flags"]),
            "changed": sum(1 for p in people if ST_CHANGED in p["flags"]),
            "bad": sum(1 for p in people if ST_BAD in p["flags"]),
            "extra": len(report["extra"])}


def _remind(p, title, due, has_fill):
    """给本人的一句催办话术。"""
    when = f"请于 {due} 前" if due else "请尽快"
    head = f"{p['name']}：【{title}】"
    if p["status"] == ST_MISSING:
        return f"{head}还没收到你的表，{when}交回。"
    if p["status"] == ST_EMPTY:
        return f"{head}你交的表还没填内容，{when}填好重新交。"
    parts = []
    if p["blanks"]:
        parts.append("有漏填：" + "，".join(_blank_text(b) for b in p["blanks"]))
    if p.get("invalid"):
        parts.append("填写有误：" + "；".join(issue_text(i) for i in p["issues"] if i["code"] in _INVALID_CODES))
    if p["changes"]:
        tip = "请用下发的原表填写，只填黄色的列" if has_fill else "请用下发的原表填写"
        parts.append("和下发时不一样：" + "，".join(p["changes"]) + f"（{tip}）")
    if ST_BAD in p["flags"]:
        parts.append("；".join(p["notes"]) or "表头和别人的对不上")
    if not parts:
        return ""
    return f"{head}你交的表{'；'.join(parts)}。{when}改好重新交。"


def notice_text(report, title, due="", now=None, has_fill=False):
    """回收通报：一段可以直接复制到群里的文字，后面附逐人催办话术。"""
    now = now or datetime.datetime.now()
    c = counts(report)
    people = report["people"]
    unit = "人" if report.get("has_list") else "份"
    if c["expected"] is not None:
        lines = [f"【{title}】回收进度（{_now_text(now)}）应交 {c['expected']} {unit}，"
                 f"已交 {c['submitted']} {unit}，合格 {c['ok']} {unit}"]
    else:
        lines = [f"【{title}】回收进度（{_now_text(now)}）已收 {c['submitted']} 份，合格 {c['ok']} 份"
                 "（没有应交名单，查不出谁没交）"]

    def names(flag):
        return [p for p in people if flag in p["flags"] or (flag == ST_MISSING and p["status"] == ST_MISSING)]

    miss = names(ST_MISSING)
    if miss:
        lines.append(f"未交（{len(miss)}）：" + "、".join(p["name"] for p in miss))
    empty = names(ST_EMPTY)
    if empty:
        lines.append(f"交了空表（{len(empty)}）：" + "、".join(p["name"] for p in empty))
    part = names(ST_PARTIAL)
    if part:
        lines.append(f"有漏填（{len(part)}）：")
        lines += [f"  {p['name']}：" + "，".join(_blank_text(b) for b in p["blanks"]) for p in part]
    inv = names(ST_INVALID)
    if inv:
        lines.append(f"填写有误（{len(inv)}）：")
        lines += [f"  {p['name']}：" + "；".join(issue_text(i) for i in p["issues"] if i["code"] in _INVALID_CODES)
                  for p in inv]
    chg = names(ST_CHANGED)
    if chg:
        lines.append(f"改了不该改的（{len(chg)}）：")
        lines += [f"  {p['name']}：" + "；".join(p["changes"]) for p in chg]
    bad = names(ST_BAD)
    if bad:
        lines.append(f"没法查（{len(bad)}）：")
        lines += [f"  {p['name']}：" + ("；".join(p["notes"]) or "表头和别人的对不上") for p in bad]
    if report["extra"]:
        what = "名单外的分组" if report["mode"] == "rows" else "名单外的文件"
        lines.append(f"{what}（{len(report['extra'])}）：" + "、".join(report["extra"]))
    if report["errors"] and not any(i["code"] == "unreadable" for p in people for i in p.get("issues", ())):
        lines.append(f"读不了的文件（{len(report['errors'])}）：" + "；".join(report["errors"]))
    if report.get("fixed") or report.get("still"):
        lines.append("比上次：")
        if report.get("new"):
            lines.append("  新交：" + "、".join(report["new"]))
        if report.get("fixed"):
            lines.append("  改好了：" + "、".join(report["fixed"]))
        if report.get("still"):
            lines.append("  还差：" + "、".join(report["still"]))
    elif report.get("new"):
        lines.append("比上次新交：" + "、".join(report["new"]) + "。")
    todo = [p for p in people if p["status"] != ST_OK]
    if todo:
        lines.append(f"请以上同事于 {due} 前补交。" if due else "请以上同事尽快补交。")
    elif people:
        lines.append("全部交齐、填写完整，可以汇总了。")
    reminds = [r for r in (_remind(p, title, due, has_fill) for p in todo) if r]
    if reminds:
        lines += ["", "—— 逐人催办（可单独复制发给本人）——"] + reminds
    return "\n".join(lines) + "\n"


def write_progress(report, path, title, now=None, due="", has_fill=False):
    """回收进度.xlsx：每人一行，未交 / 空表标红、有漏填 / 填写有误 / 有改动标橙、合格标绿；另有「问题明细」。

    文件正被打开、存不进去时换个带时间的文件名，返回实际写入的路径。
    """
    now = now or datetime.datetime.now()
    c = counts(report)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "回收进度"
    head = ["序号", "应交", "状态", "交的文件", "数据行数", "漏填", "填写有误", "和下发时不一样", "交表时间", "说明",
            "催办话术"]
    summary = (f"【{title}】应交 {c['expected']}，已交 {c['submitted']}，合格 {c['ok']}"
               if c["expected"] is not None else f"【{title}】已收 {c['submitted']} 份，合格 {c['ok']} 份")
    ws.append([f"{summary}（{_now_text(now)}）"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(head))
    ws.append(head)
    bold = Font(bold=True)
    for cell in ws[2]:
        cell.font = bold
        cell.fill = PatternFill("solid", fgColor="DDEBF7")
    order = sorted(range(len(report["people"])), key=lambda j: (_RANK[report["people"][j]["status"]], j))
    for n, j in enumerate(order, start=1):
        p = report["people"][j]
        status = p["status"] if len(p["flags"]) <= 1 else "、".join(p["flags"])
        ws.append([n, p["name"], status, "、".join(p["files"]), p["rows"] or "",
                   "；".join(_blank_text(b) for b in p["blanks"]),
                   "；".join(issue_text(i) for i in p.get("issues", ()) if i["code"] in _INVALID_CODES),
                   "；".join(p["changes"]), p["time"],
                   "；".join(p["notes"]), _remind(p, title, due, has_fill)])
        ws.cell(row=ws.max_row, column=3).fill = PatternFill("solid", fgColor=_STATUS_COLOR[p["status"]])
    for x in report["extra"]:
        ws.append(["", "", EXTRA, x, "", "", "", "", "", "不在应交名单里，请核对是不是这次的表", ""])
    for col, w in zip("ABCDEFGHIJK", (6, 14, 12, 30, 9, 40, 36, 30, 12, 30, 50)):
        ws.column_dimensions[col].width = w
    wrap = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=3):
        for cell in row:
            cell.alignment = wrap
    ws.freeze_panes = "C3"
    ws2 = wb.create_sheet("问题明细")
    ws2.append(["应交", "交的文件", "列", "问题", "格数", "行号"])
    for cell in ws2[1]:
        cell.font = bold

    def rows_cell(b):
        return "、".join(str(r) for r in b["rows"]) + ("…" if len(b["rows"]) < b["count"] else "")

    for p in report["people"]:
        for b in p["blanks"]:
            ws2.append([p["name"], "、".join(p["files"]), b["col"], "漏填", b["count"],
                        "整列没有" if b.get("no_col") else rows_cell(b)])
        for b in p.get("invalid", ()):
            ws2.append([p["name"], "、".join(p["files"]), b["col"], INVALID_TEXT[b["code"]], b["count"], rows_cell(b)])
    for col, w in zip("ABCDEF", (14, 30, 20, 12, 8, 60)):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = "A2"
    return _save_book(wb, path, now)


def _save_book(wb, path, now):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    try:
        wb.save(path)
        return path
    except PermissionError:          # 上一轮的进度表还开着
        root, ext = os.path.splitext(path)
        alt = f"{root}_{now:%H%M%S}{ext}"
        wb.save(alt)
        return alt
    finally:
        wb.close()


def result_data(report, title, due="", now=None):
    """回收结果.json 的内容：逐人的状态和问题清单（只有名字、列名、文件名、行号、个数和编码，不含单元格的值）。

    给技能和内部版回答追问用（「东城还差什么」「只给有漏填的人发催办」），不用回去解析通报文字。
    """
    now = now or datetime.datetime.now()
    return {"version": 1, "title": title, "time": now.isoformat(timespec="seconds"), "due": due,
            "counts": counts(report), "required": report["required"],
            "people": [{"name": p["name"], "status": p["status"], "flags": p["flags"], "files": p["files"],
                        "rows": p["rows"], "issues": p.get("issues", [])} for p in report["people"]],
            "extra": report["extra"]}


def _round_issues(report):
    """这一轮按人记问题的编码和个数（不记行号）。未交的不记——下一轮交了算「新交」，不算「改好了」。"""
    out = {}
    for p in report["people"]:
        got = issue_counts([i for i in p.get("issues", ())
                            if ISSUE_NAG.get(i["code"]) and ISSUE_STATUS[i["code"]] != ST_MISSING])
        if got:
            out[p["name"]] = got
    return out


def _round_diff(prev, report):
    """和上一轮比：(改好了的人, 还差的人及漏填变化)。上一轮没记问题（旧版写的记录）时返回 ([], [])。"""
    before = prev.get("issues")
    if not isinstance(before, dict):
        return [], []
    fixed, still = [], []
    for p in report["people"]:
        old = before.get(p["name"])
        if not old:
            continue
        if p["status"] == ST_OK:
            fixed.append(p["name"])
        elif p["status"] != ST_MISSING:
            now = issue_counts([i for i in p.get("issues", ()) if i["code"] in _BLANK_CODES])
            a = sum(v for k, v in old.items() if k in _BLANK_CODES)
            b = sum(now.values())
            still.append(f"{p['name']} 漏填 {a} 格 → {b} 格" if a and b and a != b else p["name"])
    return fixed, still


def _load_rounds(out_dir):
    try:
        with open(os.path.join(out_dir, ROUNDS_NAME), encoding="utf-8") as f:
            data = json.load(f)
        return data.get("rounds", []) if isinstance(data, dict) else []
    except (OSError, ValueError):
        return []


def _save_rounds(out_dir, rounds):
    try:
        with open(os.path.join(out_dir, ROUNDS_NAME), "w", encoding="utf-8") as f:
            json.dump({"rounds": rounds[-ROUNDS_KEEP:]}, f, ensure_ascii=False, indent=1)
    except OSError:
        pass


# ---------- 下发：拆分时指定了「让对方填写的列」 ----------

FILL_COLOR = "FFF2CC"            # 让对方填写的格子：浅黄
SEND_LIST_NAME = "发送清单.xlsx"
ISSUE_NOTE_NAME = "下发通知.txt"


def prepare_issue(output_path, saved, config, log=None):
    """把拆出来的表变成「待填写的表」，并在输出目录写下发记录、发送清单、下发通知。

    saved：[(分组名, 文件绝对路径, 是否汇总文件)]。对每个文件：
      · 让对方填写的列（config['fill_columns']）标浅黄、解除锁定；表里没有的列加在最后；
      · 其余格子锁定，工作表保护不设口令（审阅 → 撤销工作表保护即可解除）；
      · 末尾加一个「填报说明」sheet，写明填报对象、截止时间、填哪几列——查收时靠它认出是谁交的。
    下发记录.json 只有分组名、文件名、列名、行数和每行的短哈希，不含单元格内容。返回记录路径。
    """
    log = log or (lambda m: None)
    specs = parse_fill_spec(config.get("fill_columns") or [], log)
    fill = [n for n, _r in specs]
    if not fill:
        return ""
    rules = {n: r for n, r in specs if r["type"] != "text"}
    optional = [n for n, r in specs if r.get("optional")]
    due = soft_clean(config.get("due", ""))
    title = soft_clean(config.get("issue_title", "")) or os.path.splitext(
        os.path.basename(os.path.normpath(config.get("input_path", ""))))[0]
    lock_order, lock_names, files = [], {}, []
    for group, path, is_summary in saved:
        try:
            names = _mark_fill(path, group, fill, title, due, config, specs, log)
        except Exception as e:
            log(f"  ⚠️ 没能标出填写列：{os.path.basename(path)}：{e}")
            continue
        for sheet_names in names:
            for k, n in zip(col_keys(sheet_names), sheet_names):
                if k in lock_keys(sheet_names, fill) and k not in lock_names:
                    lock_order.append(k)
                    lock_names[k] = n
        files.append((group, path, is_summary))
    groups = {}
    for group, path, is_summary in files:
        if not is_summary:
            continue
        g = groups.setdefault(group, {"files": [], "sheets": {}})
        g["files"].append(os.path.relpath(path, output_path))
    for group, g in groups.items():
        if len(g["files"]) == 1:          # 一个分组只有一个文件时才记指纹（文件夹不合并时一组多个，不比对）
            try:
                g["sheets"] = fingerprint_file(os.path.join(output_path, g["files"][0]), config, lock_order)
            except Exception as e:
                log(f"  ⚠️ 没能记下「{group}」的指纹（查收时不比对改动）：{e}")
        g["file"] = g["files"][0]
    record = {"version": 1, "title": title, "due": due, "created": datetime.datetime.now().isoformat(timespec="seconds"),
              "split_column": soft_clean(config.get("split_column", "")), "fill_columns": fill,
              "lock_keys": lock_order, "lock_columns": [lock_names[k] for k in lock_order], "groups": groups}
    if rules or optional:            # 写了括号才有：不写括号时记录和 v2.10 一字不差
        record["required"] = [n for n in fill if n not in optional]
        record["rules"] = rules
    rec_path = os.path.join(output_path, RECORD_NAME)
    with open(rec_path, "w", encoding="utf-8") as f:
        json.dump(record, f, ensure_ascii=False, indent=1)
    _write_send_list(output_path, groups, saved)
    when = f"请于 {due} 前" if due else "请尽快"
    lines = [f"【{title}】{when}填写附件里黄色的列（{'、'.join(fill)}），其他列已锁定不用动；填好后把原文件交回。"]
    lines += [f"{g}：{'、'.join(d['files'])}" for g, d in groups.items()]
    with open(os.path.join(output_path, ISSUE_NOTE_NAME), "w", encoding="utf-8-sig") as f:
        f.write("\n".join(lines) + "\n")
    log(f"  ✍️ 已标出填写列 {'、'.join(fill)}（浅黄、可编辑，其余列锁定），写了下发记录和发送清单")
    return rec_path


def _mark_fill(path, group, fill, title, due, cfg, specs=None, log=None):
    """标黄填写列、锁定其余格子、加「填报说明」sheet。返回各 sheet 的列名（含新加的填写列）。

    specs（parse_fill_spec 的结果）里有规则的列再加数据验证：下拉框、日期、数字、整数、手机号；
    日期列显示成 yyyy-mm-dd，手机号 / 电话列设成文本格式。数据验证拦不住粘贴，所以查收时照样按规则再查。
    """
    from openpyxl.styles import Protection
    rules = dict(specs or [])
    noted = set()
    wb = openpyxl.load_workbook(path)
    yellow = PatternFill("solid", fgColor=FILL_COLOR)
    unlocked = Protection(locked=False)
    out = []
    try:
        for ws in wb.worksheets:
            head = [tuple(c.value for c in row) for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, HEAD_SCAN))]
            h = sp.resolve_header_row(head, cfg)
            if h == -1 or h > len(head):
                continue
            names = col_names(head, h)
            keys = col_keys(names)
            cols = []
            for want in fill:
                c = _key_index(keys, want)
                if c is None:                          # 表里没有这一列：加在最后
                    names.append(want)
                    keys = col_keys(names)
                    c = len(names) - 1
                    _add_col(ws, h, c + 1, want)
                cols.append(c + 1)
            last = ws.max_row
            for want, c in zip(fill, cols):
                rule = rules.get(want, {"type": "text"})
                fmt = {"date": "yyyy-mm-dd", "mobile": "@", "phone": "@"}.get(rule["type"])
                ws.cell(row=h, column=c).fill = yellow
                for r in range(h + 1, last + 1):
                    cell = ws.cell(row=r, column=c)
                    if isinstance(cell, MergedCell):
                        continue
                    cell.fill = yellow
                    cell.protection = unlocked
                    if fmt:
                        cell.number_format = fmt
                if rule["type"] == "text" or last <= h:
                    continue
                letter = get_column_letter(c)
                dv = _validation(rule, want, h + 1, letter)
                if dv is not None:
                    dv.add(f"{letter}{h + 1}:{letter}{last}")
                    ws.add_data_validation(dv)
                elif rule["type"] == "list" and want not in noted and log:
                    noted.add(want)
                    log(f"  ℹ️ 「{want}」的选项太长（或带英文逗号），没加下拉框；收回来照样按选项查")
            ws.protection.sheet = True
            ws.protection.formatColumns = False      # 允许调列宽、行高，看得清楚
            ws.protection.formatRows = False
            out.append(names)
        if GUIDE_SHEET in wb.sheetnames:
            wb.remove(wb[GUIDE_SHEET])
        g = wb.create_sheet(GUIDE_SHEET)
        for row in (["填报对象", group], ["任务", title], ["截止时间", due or "（未定）"],
                    ["要填的列", "、".join(fill)],
                    ["怎么填", "只在浅黄色的列里填写；其他列已锁定，不用改。没有设口令，确需修改可在「审阅 → 撤销工作表保护」解除"],
                    ["交回", "填好后把这个文件原样交回，文件名可以不改"]):
            g.append(row)
        g.column_dimensions["A"].width = 12
        g.column_dimensions["B"].width = 70
        for c in g["A"]:
            c.font = Font(bold=True)
        if any(r["type"] != "text" or r.get("optional") for r in rules.values()):
            g.append([])
            g.append(["列名", "要求", "可填的值 / 示例"])
            for c in g[g.max_row]:
                c.font = Font(bold=True)
            for want in fill:
                rule = rules.get(want, {"type": "text"})
                example = " / ".join(rule["options"]) if rule["type"] == "list" else RULE_EXAMPLES.get(rule["type"], "")
                g.append([want, _rule_text(rule), example])
            g.column_dimensions["C"].width = 40
        g.protection.sheet = True        # 防止无意中改了「填报对象」，查收时认不出是谁交的（不设口令）
        wb.save(path)
    finally:
        wb.close()
    return out


def _write_send_list(output_path, groups, saved):
    """发送清单.xlsx：每组一行，发给谁、发哪个文件，「接收人」「发送情况」留给使用者自己填。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "发送清单"
    ws.append(["序号", "分组", "文件", "接收人", "发送情况", "备注"])
    for c in ws[1]:
        c.font = Font(bold=True)
    n = 0
    for group, g in groups.items():
        n += 1
        ws.append([n, group, "、".join(g["files"]), "", "", ""])
    if not groups:                       # 只有到人文件时，按文件列
        for group, path, _s in saved:
            n += 1
            ws.append([n, group, os.path.relpath(path, output_path), "", "", ""])
    for col, w in zip("ABCDEF", (6, 16, 40, 14, 12, 20)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(os.path.join(output_path, SEND_LIST_NAME))
    wb.close()


# ---------- 对外入口：查收 / 汇总（签名与 run_split 一致） ----------

def returned_files(input_path, output_root):
    """回收文件夹里的表：[(绝对路径, 相对路径)]，以及名单表的路径。跳过结果目录、锁文件和名单表。

    .et / .csv 这类读不了的文件也算进来：它们是有人交了的表，扫描时按 unreadable() 记成读不了，
    通报里写成「打不开：是 WPS 表格自己的格式…」，而不是把交了的人写成未交。
    """
    if os.path.isfile(input_path):
        return [(input_path, os.path.basename(input_path))], []
    out_abs = os.path.normcase(os.path.abspath(output_root)) if output_root else ""
    files, lists = [], []
    for root, dirs, names in os.walk(input_path):
        dirs[:] = sorted(d for d in dirs if d != RESULT_DIR and not (
            out_abs and os.path.normcase(os.path.abspath(os.path.join(root, d))).startswith(out_abs)))
        for f in sorted(names):
            lf = f.lower()
            if (not lf.endswith(TABLE_EXTS + tuple(OTHER_EXTS)) or f.startswith("~$")
                    or lf.endswith("__tmp__.xlsx")):
                continue
            stem = os.path.splitext(f)[0]
            path = os.path.join(root, f)
            if any(w in stem for w in LIST_WORDS):
                lists.append(path)
            elif not stem.startswith(("回收进度", "汇总_")):
                files.append((path, os.path.relpath(path, input_path)))
    return files, lists


def _prepare(config, log):
    """读配置，定下名单、该填的列、下发记录、输出目录，并扫描所有表。"""
    input_path = config["input_path"]
    base = input_path if os.path.isdir(input_path) else os.path.dirname(input_path)
    out_dir = config.get("output_path") or os.path.join(base, RESULT_DIR)
    files, lists = returned_files(input_path, out_dir)
    rec_path = config.get("collect_record") or find_record(input_path)
    record = load_record(rec_path) if rec_path else None
    if rec_path and record is None:
        log(f"⚠️ 下发记录读不了，按没有下发记录处理：{rec_path}")
    expect = split_names(config.get("collect_expect") or [])
    source = "指定的名单" if expect else ""
    if not expect and record:
        expect, source = list(record["groups"]), "下发记录"
    if not expect and lists:
        for lp in lists:
            try:
                expect += [n for n in read_name_list(lp) if n not in expect]
            except Exception as e:
                log(f"⚠️ 名单表读不了：{os.path.basename(lp)}：{e}")
        source = "名单表 " + "、".join(os.path.basename(lp) for lp in lists) if expect else ""
    wanted = split_names(config.get("collect_required") or [])
    req_source = "指定的列" if wanted else ""
    if not wanted and record and record.get("fill_columns"):
        wanted, req_source = list(record.get("required") or record["fill_columns"]), "下发时的填写列"
    rules = (record or {}).get("rules") if isinstance((record or {}).get("rules"), dict) else None
    try:
        issued = datetime.date.fromisoformat(str((record or {}).get("created", ""))[:10])
    except ValueError:
        issued = datetime.date.today()
    by = soft_clean(config.get("collect_by", ""))
    rows_mode = bool(by) and os.path.isfile(input_path)
    title = soft_clean(config.get("collect_title")) or (record or {}).get("title") or (
        os.path.basename(os.path.normpath(base)) if not rows_mode else os.path.splitext(os.path.basename(input_path))[0])
    due = soft_clean(config.get("collect_due")) or (record or {}).get("due", "")
    return {"input": input_path, "out_dir": out_dir, "files": files, "record": record, "expect": expect,
            "source": source, "wanted": wanted, "req_source": req_source or "大多数表都填了的列",
            "by": by, "rows_mode": rows_mode, "title": title, "due": due,
            # 文件夹模式：没指定分人列时用下发时的拆分列认文件（只用来认，不按它分人）
            "claim_by": by or ((record or {}).get("split_column", "") if not rows_mode else ""),
            "fp_keys": (record or {}).get("lock_keys") if record and not rows_mode else None,
            "rules": rules, "base": issued}


def _inspect(config, log, progress, stop):
    """扫描并逐人判定。返回 (ctx, res, report)；中途停止时 report 为 None。"""
    ctx = _prepare(config, log)
    files = ctx["files"]
    log(f"共找到 {len(files)} 个表" + (f"，应交名单 {len(ctx['expect'])} 个（来自{ctx['source']}）" if ctx["expect"] else
                                   "，没有应交名单（可在「应交名单」里填，或放一张「应交名单」表进文件夹）"))
    if not files:
        log("⚠️ 没有找到交回来的表（.xlsx / .xls），请检查文件夹")
    res = scan(files, config, log=log, by=ctx["claim_by"] or None, fp_keys=ctx["fp_keys"], stop=stop,
               progress=progress, rules=ctx["rules"], base=ctx["base"])
    if not res["complete"]:
        return ctx, res, None
    report = analyze(files, res, ctx["expect"], ctx["wanted"], ctx["record"], ctx["by"], ctx["rows_mode"])
    if report["required"]:
        log(f"检查这些列有没有填（{ctx['req_source']}）：" + "、".join(report["required"]))
    else:
        log("⚠️ 没有定出该填的列，只查谁没交、谁交了空表；可在「必须填的列」里指定")
    if report["unknown_required"]:
        log("⚠️ 表头里找不到这些列，请核对列名：" + "、".join(report["unknown_required"]))
    if ctx["rules"]:
        log("按下发时的填写要求查格式：" + "；".join(f"{n} {_rule_text(r)}" for n, r in ctx["rules"].items()))
    for n in report["notes"]:
        log(f"⚠️ {n}")
    return ctx, res, report


def check_collect(config, log_fn=None, progress_fn=None, stop_flag=None):
    """查收：逐人查交表情况，写「回收通报.txt」「回收进度.xlsx」，返回输出目录。

    config：input_path（回收文件夹，或一张各填各行的总表）、output_path（留空 = 输入旁的「汇总结果」）、
    collect_expect（应交名单）、collect_required（必须填的列）、collect_by（按哪一列分人）、
    collect_record（下发记录.json，留空自动找）、collect_title（通报标题）、collect_due（截止时间）；
    表头识别沿用 header_mode / header_row。
    """
    log = log_fn or (lambda m: None)
    now = datetime.datetime.now()
    ctx, _res, report = _inspect(config, log, lambda f: progress_fn and progress_fn(f * 0.9), stop_flag)
    out_dir = ctx["out_dir"]
    summary = {"mode": "collect_check", "output": out_dir, "stopped": report is None}
    if report is None:
        log("⛔ 已停止，这一轮没有生成通报")
        log(f"[SUMMARY] {json.dumps(summary, ensure_ascii=False)}")
        return out_dir
    os.makedirs(out_dir, exist_ok=True)
    rounds = _load_rounds(out_dir)
    people = report["people"]
    submitted = [p["name"] for p in people if p["status"] != ST_MISSING]
    if rounds:
        before = set(rounds[-1].get("submitted", []))
        report["new"] = [n for n in submitted if n not in before]
        report["fixed"], report["still"] = _round_diff(rounds[-1], report)
    has_fill = bool(ctx["record"] and ctx["record"].get("fill_columns"))
    text = notice_text(report, ctx["title"], ctx["due"], now, has_fill)
    notice = os.path.join(out_dir, NOTICE_NAME)
    try:
        with open(notice, "w", encoding="utf-8-sig") as f:
            f.write(text)
    except OSError as e:
        log(f"⚠️ 通报文字没写进文件（{e}），请直接从下面复制")
        notice = ""
    if progress_fn:
        progress_fn(0.95)
    try:
        progress = write_progress(report, os.path.join(out_dir, PROGRESS_NAME), ctx["title"], now, ctx["due"],
                                  has_fill)
    except Exception as e:
        log(f"⚠️ 进度表没存成：{e}")
        progress = ""
    result = os.path.join(out_dir, RESULT_NAME)
    try:
        with open(result, "w", encoding="utf-8") as f:
            json.dump(result_data(report, ctx["title"], ctx["due"], now), f, ensure_ascii=False, indent=1)
    except OSError as e:
        log(f"⚠️ 回收结果.json 没写成：{e}")
        result = ""
    rounds.append({"time": now.isoformat(timespec="seconds"), "submitted": submitted,
                   "ok": [p["name"] for p in people if p["status"] == ST_OK], "issues": _round_issues(report)})
    _save_rounds(out_dir, rounds)
    if progress_fn:
        progress_fn(1.0)
    log("")
    for line in text.rstrip("\n").split("\n"):
        log(line)
    log(f"\n📁 进度表：{progress}" if progress else "")
    summary.update(counts(report), notice=notice, progress=progress, title=ctx["title"], result=result)
    log(f"[SUMMARY] {json.dumps(summary, ensure_ascii=False)}")
    return out_dir


def run_collect(config, log_fn=None, progress_fn=None, stop_flag=None):
    """汇总：查收一遍后把交回来的表合成一张（空表不合），最后一个 sheet「回收统计」逐人列出状态。返回结果文件路径。"""
    log = log_fn or (lambda m: None)
    t0 = time.time()
    ctx, res, report = _inspect(config, log, lambda f: progress_fn and progress_fn(f * 0.3), stop_flag)
    out_dir = ctx["out_dir"]
    summary = {"mode": "collect", "output": "", "stopped": report is None}
    if report is None:
        log("⛔ 已停止，没有生成汇总表")
        log(f"[SUMMARY] {json.dumps(summary, ensure_ascii=False)}")
        return ""
    files = ctx["files"]
    rel_idx = {rel: i for i, (_p, rel) in enumerate(files)}
    labels, take = {}, set()
    for p in report["people"]:
        if p["status"] in (ST_MISSING, ST_EMPTY):
            continue
        for rel in p["files"]:
            labels[rel] = p["name"]
            take.add(rel_idx[rel])
    if report["mode"] == "files":
        take |= {rel_idx[r] for r in report["extra"]}
    else:
        take = set(range(len(files)))          # 总表：整张表就是汇总对象
    sel = [files[i] for i in sorted(take)]
    if not sel:
        log("⚠️ 还没有可以汇总的表")
        log(f"[SUMMARY] {json.dumps(summary, ensure_ascii=False)}")
        return ""
    sel_idx = {files[i][1]: n for n, i in enumerate(sorted(take))}
    sel_sheets = [dict(s, file=sel_idx[s["rel"]]) for s in res["sheets"] if s["rel"] in sel_idx]
    # 只交回一份时也按「每人一个来源」合并，「来源」列写交表人而不是 sheet 名
    mode = pick_mode(len(sel) if report["mode"] == "rows" else max(2, len(sel)), sel_sheets)
    groups = compare(sel_sheets, mode)
    completeness(groups, ctx["wanted"])
    output = _result_path(out_dir, ctx["title"])
    log(f"\n开始汇总 {len(sel)} 个表…")
    st = merge_start(output)
    st["base"] = ctx["base"].isoformat()

    def _prog(frac, _text=""):
        if progress_fn:
            progress_fn(0.3 + 0.7 * frac)

    done = merge_step(st, sel, config, mode, {k: g["required"] for k, g in groups.items()}, log=log,
                      progress=_prog, labels=labels if report["mode"] == "files" else None,
                      people=report["people"] if report["mode"] == "files" else None, stop=stop_flag,
                      rules=ctx["rules"])
    summary.update(counts(report))
    return _finish_merge(st, done, output, ctx["title"], t0, summary, log, progress_fn)


def _result_path(out_dir, title):
    """汇总表的文件名：汇总_{标题}_{月日时分}.xlsx，已存在就加序号。"""
    stamp = datetime.datetime.now().strftime("%m%d%H%M")
    output = os.path.join(out_dir, f"汇总_{safe_filename(title)}_{stamp}.xlsx")
    n = 2
    while os.path.exists(output):
        output = os.path.join(out_dir, f"汇总_{safe_filename(title)}_{stamp}_{n}.xlsx")
        n += 1
    return output


def _finish_merge(st, done, output, title, t0, summary, log, progress_fn):
    """合并收尾：打完成日志和 [SUMMARY]，返回汇总表路径（没生成返回 ""）。"""
    if not done:
        log("⛔ 已停止，没有生成汇总表")
        summary["stopped"] = True
        log(f"[SUMMARY] {json.dumps(summary, ensure_ascii=False)}")
        return ""
    if progress_fn:
        progress_fn(1.0)
    merged = [s for s in st["sources"] if s["status"] == "合并"]
    rows = sum(s["rows"] for s in merged)
    mm, ss = divmod(int(time.time() - t0 + 0.5), 60)
    text = merge_text(st, title)
    if os.path.exists(output):
        log(f"\n✅ 汇总完成：{_merged_what(merged)}，共 {rows} 行，"
            f"耗时 {f'{mm} 分 {ss} 秒' if mm else f'{ss} 秒'}")
        log(f"📄 汇总表：{output}")
    else:
        log("⚠️ 这些表里没有能合并的数据")
        output = ""
    summary.update(output=output, files=len({s["file"] for s in merged}), rows=rows,
                   failed_files=len(st["errors"]), title=title, report_text=text)
    log(f"[SUMMARY] {json.dumps(summary, ensure_ascii=False)}")
    return output


def merge_tables(config, log_fn=None, progress_fn=None, stop_flag=None):
    """普通汇总：文件夹里同一套模板的表合成一张（选的是一个文件时，把它的各个 sheet 合成一张）。

    不查谁没交、不认人：每个文件都是一个来源，「来源」列写文件名（一个文件时写 sheet 名）。
    config：input_path（文件夹或一个文件）、output_path（留空 = 输入旁的「汇总结果」）、
    collect_title（结果文件名里的标题，留空用文件夹名 / 文件名）；表头识别沿用 header_mode / header_row。
    签名与 run_split 一致，收尾打 [SUMMARY]（mode 为 merge）。返回结果文件路径，没生成返回 ""。
    """
    log = log_fn or (lambda m: None)
    t0 = time.time()
    input_path = config["input_path"]
    is_dir = os.path.isdir(input_path)
    base = input_path if is_dir else os.path.dirname(input_path)
    out_dir = config.get("output_path") or os.path.join(base, RESULT_DIR)
    files, lists = returned_files(input_path, out_dir)
    if is_dir:                       # 普通汇总没有名单一说：名字像名单的表也照样合
        files = sorted(files + [(p, os.path.relpath(p, input_path)) for p in lists], key=lambda x: x[1])
    title = soft_clean(config.get("collect_title")) or (
        os.path.basename(os.path.normpath(base)) if is_dir else os.path.splitext(os.path.basename(input_path))[0])
    summary = {"mode": "merge", "output": "", "stopped": False}
    if not files:
        log("⚠️ 没有找到表格（.xlsx / .xls），请检查文件夹")
        log(f"[SUMMARY] {json.dumps(summary, ensure_ascii=False)}")
        return ""
    log(f"共找到 {len(files)} 个表" if is_dir else f"把「{os.path.basename(input_path)}」里的各个 sheet 合成一张")
    res = scan(files, config, log=log, stop=stop_flag, progress=lambda f: progress_fn and progress_fn(f * 0.3))
    if not res["complete"]:
        log("⛔ 已停止，没有生成汇总表")
        summary["stopped"] = True
        log(f"[SUMMARY] {json.dumps(summary, ensure_ascii=False)}")
        return ""
    mode = pick_mode(len(files), res["sheets"])
    output = _result_path(out_dir, title)
    log(f"\n开始汇总 {len(files)} 个表…" if is_dir else "\n开始合并…")
    st = merge_start(output)

    def _prog(frac, _text=""):
        if progress_fn:
            progress_fn(0.3 + 0.7 * frac)

    done = merge_step(st, files, config, mode, {}, log=log, progress=_prog, stop=stop_flag)
    return _finish_merge(st, done, output, title, t0, summary, log, progress_fn)


# ---------- 合并：分段执行（内部技能 70 秒一段），桌面版一次做完 ----------

def merge_start(output):
    """返回可续做的状态（全是能存成 JSON 的值）。之后反复调用 merge_step 直到做完。

    状态字段：
      output      结果文件路径
      next        下一个要合并的源文件序号
      sheets      {组键: {title, h, keys, names, next_row, src_col, base}}：结果 sheet 的表头行、列键、写到第几行；
                  base 是模板自带的列数，后面是「来源」列和多出的列
      sources     [{file, sheet, rows, totals, missing, extra, status}]（status：合并 / 表头对不上 / 隐藏 / …）
      hashes      {文件摘要: 文件}，内容相同的文件只合一份；dupes 被跳过的重复文件
      cells       已写进结果的格子数，用来估算存盘用时
    """
    return {"output": output, "next": 0, "sheets": {}, "sources": [], "hashes": {}, "dupes": [], "errors": [],
            "sec_per_mb": 0.0, "save_sec": 0.0, "cells": 0, "done": False}


def _open_output(st):
    """新建结果工作簿；续做时读回上次存的，去掉上次中断时多写、没记进度的 sheet、行和列。"""
    if os.path.isfile(st["output"]):
        wb = openpyxl.load_workbook(st["output"])
        titles = {s["title"] for s in st["sheets"].values()}
        for ws in list(wb.worksheets):
            if ws.title not in titles:
                wb.remove(ws)
        for s in st["sheets"].values():
            ws = wb[s["title"]]
            if ws.max_row >= s["next_row"]:
                ws.delete_rows(s["next_row"], ws.max_row - s["next_row"] + 1)
            ncols = len(s["keys"])
            if ws.max_column > ncols:
                ws.delete_cols(ncols + 1, ws.max_column - ncols)
        return wb
    if st["sheets"]:                     # 存过结果、结果文件却不在了：从头做
        st["sheets"] = {}
        st["next"], st["sources"], st["hashes"], st["dupes"], st["errors"] = 0, [], {}, [], []
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    return wb


def _new_sheet(wb, st, key, ws_src, h, names, keys, sheet):
    """以这个来源 sheet 为模板建结果 sheet：表头区域（值、格式、合并单元格、列宽）原样复制，再加「来源」列。"""
    used = {w.title.lower() for w in wb.worksheets} | {REPORT_SHEET.lower()}
    ws = wb.create_sheet(title=sp._safe_sheet_title(sheet, used))
    cells, merges, widths = sp._read_header_format(ws_src, h)
    sp._write_header(ws, cells, merges, widths)
    keys, names = list(keys), list(names)
    src_col = len(keys) + 1
    _add_col(ws, h, src_col, SOURCE_COL)
    keys.append("\0source")
    names.append(SOURCE_COL)
    ws.freeze_panes = ws.cell(row=h + 1, column=1)
    st["sheets"][key] = {"title": ws.title, "h": h, "keys": keys, "names": names, "next_row": h + 1,
                         "src_col": src_col, "base": src_col - 1}
    return ws


def _add_col(ws, h, col, name):
    """在结果 sheet 的表头行加一列，格式照前一列的表头格子。"""
    dst = ws.cell(row=h, column=col)
    dst.value = name
    if col > 1:
        left = ws.cell(row=h, column=col - 1)
        if not isinstance(left, MergedCell) and left.has_style:
            dst.font, dst.fill, dst.border, dst.alignment = (copy(left.font), copy(left.fill),
                                                             copy(left.border), copy(left.alignment))
    ws.column_dimensions[get_column_letter(col)].width = max(10, min(30, 2 * len(name) + 4))


def merge_step(st, files, cfg, mode, required=None, until=None, log=None, progress=None, commit=None,
               labels=None, people=None, stop=None, rules=None):
    """从 st['next'] 接着合并，过了 until（time.time() 时刻）就停在两个源文件之间。返回是否全部做完。

    required：{组键: [该填的列键]}，这些列空着的格子标浅红、计数。labels：{相对路径: 交表人}，写进「来源」列。
    people：查收结果，写进「回收统计」。stop() 为真时立刻停下，不存盘，返回 False。
    rules：{列名: 填报规则}（下发记录里的），有规则的列按 _coerce 规整写法（只改汇总结果），认不出的标浅橙；
    补年份用 st["base"]（下发日期，ISO 文字），没有时用今天。不传时行为不变。
    每次至少合一个文件；按已合文件的速度和上次存盘用时估算，来不及就先停。
    停下或做完时先把结果工作簿存盘，再调用 commit()（调用方在里面把 st 存盘）。
    """
    log = log or (lambda m: None)
    progress = progress or (lambda f, s: None)
    commit = commit or (lambda: None)
    wb = _open_output(st)
    total = max(1, len(files))
    done_here = 0
    while st["next"] < len(files):
        if stop and stop():
            wb.close()
            return False
        i = st["next"]
        path, rel = files[i]
        size_mb = max(0.01, os.path.getsize(path) / 1e6) if os.path.exists(path) else 0.01
        save_est = max(st["save_sec"], st["cells"] * SAVE_PER_CELL) * 1.3      # 存盘用时随结果变大而变长
        if done_here and until and time.time() + st["sec_per_mb"] * size_mb * 1.5 + save_est + 2 > until:
            break
        progress(i / total, f"合并第 {i + 1}/{len(files)} 个表")
        log(f"[{i + 1}/{len(files)}] {rel}")
        t0 = time.time()
        _merge_file(wb, st, path, rel, cfg, mode, required or {}, log, labels or {}, rules)
        st["sec_per_mb"] = max(st["sec_per_mb"], (time.time() - t0) / size_mb)
        st["next"] = i + 1
        done_here += 1
    if st["next"] >= len(files):
        progress(0.97, "写统计表")
        if wb.worksheets:
            _write_report(wb, st, people, labels or {})
        st["done"] = True
    progress(min(0.98, st["next"] / total), "保存结果")
    t0 = time.time()
    os.makedirs(os.path.dirname(st["output"]) or ".", exist_ok=True)
    if wb.worksheets:                    # 前面的表都没有可合并的数据时还没有 sheet，先不存
        _no_formulas(wb)
        wb.save(st["output"])
    wb.close()
    st["save_sec"] = time.time() - t0
    commit()
    return st["done"]


def _merge_file(wb, st, path, rel, cfg, mode, required, log, labels, rules=None):
    """把一个源文件的各 sheet 并进结果工作簿；读不了只记进 errors，不影响其余文件。"""
    why = unreadable(path)
    if why:
        st["errors"].append({"file": rel, "error": why})
        log(f"  读不了：{why}")
        return
    try:
        fh = file_hash(path)
    except OSError as e:
        st["errors"].append({"file": rel, "error": f"读取失败：{e}"})
        return
    if fh in st["hashes"]:
        st["dupes"].append({"file": rel, "same_as": st["hashes"][fh]})
        log(f"  与「{st['hashes'][fh]}」内容相同，只合一份")
        return
    st["hashes"][fh] = rel
    src_wb = None
    try:
        src_wb, sheets, is_xls = _load_sheets(path)
        states = {n: s for n, (s, _c) in hidden_info(path).items()} if is_xls else {}
        usable, header_only = [], False
        for sn, rows, ws in sheets:
            state = states.get(sn, "visible") if is_xls else ws.sheet_state
            if state != "visible":
                st["sources"].append({"file": rel, "sheet": sn, "rows": 0, "totals": 0, "status": "隐藏，未合并"})
                continue
            if _is_guide(sn):
                continue
            h = sp.resolve_header_row(rows[:_head_limit(cfg)], cfg)
            if h == -1 or h > len(rows):
                if any(_nonempty(r) for r in rows):
                    st["sources"].append({"file": rel, "sheet": sn, "rows": 0, "totals": 0,
                                          "status": "认不出表头，未合并"})
                continue
            if any(_nonempty(r) for r in rows[h:]):
                usable.append((sn, rows, ws, h))
            else:
                header_only = True
        if header_only and not usable:
            st["sources"].append({"file": rel, "sheet": "", "rows": 0, "totals": 0, "status": "只有表头，没填数据"})
        style_cache = {}     # 按源文件作用域，见 splitter._copy_cell_style
        for sn, rows, ws, h in usable:
            _merge_sheet(wb, st, rel, sn, rows, ws, h, mode, len(usable) > 1, style_cache,
                         set(required.get(group_key(mode, sn), ())), labels.get(rel), rules)
    except Exception as e:
        st["errors"].append({"file": rel, "error": f"处理失败：{e}"})
        log(f"  处理失败：{e}")
    finally:
        if src_wb is not None:
            try:
                src_wb.close()
            except Exception:
                pass


def _merge_sheet(wb, st, rel, sn, rows, ws_src, h, mode, multi_sheet, style_cache, required, who=None, rules=None):
    names = col_names(rows[:h], h)
    keys = col_keys(names)
    key = group_key(mode, sn)
    item = {"file": rel, "sheet": sn, "rows": 0, "totals": 0, "missing": [], "extra": [], "blanks": {},
            "status": "合并"}
    out = st["sheets"].get(key)
    ws = wb[out["title"]] if out else _new_sheet(wb, st, key, ws_src, h, names, keys, sn)
    out = st["sheets"][key]
    tk = out["keys"]
    _hit, bad = match(tk, keys)
    if bad:
        item["status"] = "表头对不上，未合并"
        item["rows"] = sum(1 for r in rows[h:] if _nonempty(r))
        st["sources"].append(item)
        return
    pos = {k: c for c, k in enumerate(tk)}
    dst = []
    for c, k in enumerate(keys):
        if k not in pos:                 # 多出的列：放在最后
            tk.append(k)
            out["names"].append(names[c])
            pos[k] = len(tk) - 1
            _add_col(ws, out["h"], len(tk), names[c] or f"第 {c + 1} 列")
            item["extra"].append(names[c] or f"第 {c + 1} 列")
        dst.append(pos[k] + 1)
    have = set(keys)
    item["missing"] = [out["names"][j] for j, k in enumerate(tk[:out["base"]])     # 只算模板自带的列
                       if not k.startswith("#") and k not in have]
    src_label = label(rel, sn, mode, multi_sheet, who)
    check = {c: names[c] or f"第 {c + 1} 列" for c, k in enumerate(keys) if k in required}
    blanks = dict.fromkeys(check.values(), 0)
    fix = {}                              # 列号 → (列名, 规则)：按填报规则规整写法
    for name, rule in (rules or {}).items():
        c = _key_index(keys, name)
        if c is not None and rule.get("type", "text") in BAD_CODE:
            fix[c] = (name, rule)
    try:
        base = datetime.date.fromisoformat(st.get("base") or "")
    except ValueError:
        base = datetime.date.today()
    r_out = out["next_row"]
    for idx in range(h, len(rows)):
        row = rows[idx]
        if not _nonempty(row):
            continue
        if is_total_row(row):
            item["totals"] += 1
            continue
        for c, d in enumerate(dst):
            src = ws_src.cell(row=idx + 1, column=c + 1)
            if isinstance(src, MergedCell):
                continue
            cell = ws.cell(row=r_out, column=d)
            cell.value = src.value
            if src.has_style:
                try:
                    sp._copy_cell_style(src, cell, style_cache)
                except Exception:
                    pass
            if c in check and soft_clean(src.value) == "":
                blanks[check[c]] += 1
                cell.fill = BLANK_FILL
            elif c in fix and soft_clean(src.value) != "":
                _fix_cell(st, cell, src.value, fix[c], base, src_label, idx + 1)
        ws.cell(row=r_out, column=out["src_col"]).value = src_label
        r_out += 1
        item["rows"] += 1
    st["cells"] = st.get("cells", 0) + item["rows"] * (len(dst) + 1)
    out["next_row"] = r_out
    item["blanks"] = {n: v for n, v in blanks.items() if v}
    st["sources"].append(item)


def _fix_cell(st, cell, value, name_rule, base, src_label, src_row):
    """有填报规则的格子：认得出就改成规范写法（记个数和两个例子），认不出就标浅橙。"""
    name, rule = name_rule
    ok, new = _coerce(value, rule, base)
    phone = rule["type"] in ("mobile", "phone")
    if not ok:
        cell.fill = INVALID_FILL
        u = st.setdefault("uncoerced", {"n": 0, "ex": []})
        u["n"] += 1
        if len(u["ex"]) < 3:
            u["ex"].append([src_label, src_row, _mask(value) if phone else str(value)])
        return
    if rule["type"] == "date":
        cell.number_format = "yyyy-mm-dd"
    elif phone:
        cell.number_format = "@"
    if new == value and type(new) is type(value):
        return
    cell.value = new
    d = st.setdefault("coerced", {}).setdefault(name, {"n": 0, "ex": []})
    d["n"] += 1
    after = new.strftime("%Y-%m-%d") if isinstance(new, datetime.datetime) else str(new)
    before = str(value)
    if phone:
        before, after = _mask(before), _mask(after)
    if len(d["ex"]) < 2 and [before, after] not in d["ex"]:
        d["ex"].append([before, after])


def _write_report(wb, st, people=None, labels=None):
    """结果工作簿最后一个 sheet「汇总说明」：逐人（或逐个来源）列出状态、合了几行、漏填、缺哪些列、多哪些列。

    没有查收结果（people 为 None，普通汇总）时不写「应交」「状态」两列。
    """
    labels = labels or {}
    if REPORT_SHEET in wb.sheetnames:
        wb.remove(wb[REPORT_SHEET])
    ws = wb.create_sheet(REPORT_SHEET)
    cut = 0 if people is not None else 2

    def put(row):
        ws.append(row[cut:])

    put(["应交", "状态", "来源文件", "sheet", "汇总行数", "未汇总的合计/小计行", "漏填的格子（标了浅红）",
         "缺少的列（留空）", "多出的列（放在最后）", "说明"])
    bold = Font(bold=True)
    for c in ws[1]:
        c.font = bold

    def src_row(who, status, s):
        ok = s["status"] == "合并"
        return [who, status, s["file"], s["sheet"], s["rows"] if ok else 0, s.get("totals", 0),
                "、".join(f"{n} {v} 格" for n, v in s.get("blanks", {}).items()),
                "、".join(s.get("missing", [])), "、".join(s.get("extra", [])),
                "" if ok else s["status"] + (f"（{s['rows']} 行）" if s["rows"] else "")]

    shown = set()
    for p in people or []:
        mine = [s for s in st["sources"] if s["file"] in p["files"]]
        status = p["status"] if len(p["flags"]) <= 1 else "、".join(p["flags"])
        if not mine:
            note = "空表，未汇总" if p["status"] == ST_EMPTY else "；".join(p["notes"])
            put([p["name"], status, "、".join(p["files"]), "", 0, 0, "", "", "", note])
        for s in mine:
            put(src_row(p["name"], status, s))
            ws.cell(row=ws.max_row, column=2).fill = PatternFill("solid", fgColor=_STATUS_COLOR[p["status"]])
            shown.add(id(s))
        if not mine:
            ws.cell(row=ws.max_row, column=2).fill = PatternFill("solid", fgColor=_STATUS_COLOR[p["status"]])
    for s in st["sources"]:
        if id(s) not in shown:
            who = labels.get(s["file"], EXTRA if people else "")
            put(src_row(who, "", s))
    for d in st["dupes"]:
        put(["", "", d["file"], "", 0, 0, "", "", "", f"与「{d['same_as']}」内容完全相同，只合了一份"])
    for e in st["errors"]:
        put(["", "", e["file"], "", 0, 0, "", "", "", e["error"]])
    if st.get("coerced") or st.get("uncoerced"):
        ws.append([])
        ws.append(["自动规整（只改了汇总表，原文件没动）"])
        ws.cell(row=ws.max_row, column=1).font = bold
        for name, d in st.get("coerced", {}).items():
            ws.append([name, f"{d['n']} 格", "例：" + "；".join(f"{a} → {b}" for a, b in d["ex"])])
        u = st.get("uncoerced")
        if u:
            ws.append(["没敢改的", f"{u['n']} 格", "；".join(f"{lab} 第 {r} 行「{v}」" for lab, r, v in u["ex"])
                       + ("…" if u["n"] > len(u["ex"]) else "") + "，已标浅橙"])
    for col, w in zip("ABCDEFGHIJ", (14, 12, 28, 14, 10, 18, 26, 24, 24, 36)[cut:]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


def _merged_what(merged):
    n = len({s["file"] for s in merged})
    return f"把 {len(merged)} 个 sheet 合成了一张" if n == 1 and len(merged) > 1 else f"汇总了 {n} 个表"


def merge_text(st, title):
    """汇总结果的文字说明：合了几个表几行，哪些没合进去、为什么。界面结果卡片直接显示。"""
    merged = [s for s in st["sources"] if s["status"] == "合并"]
    lines = [f"【{title}】{_merged_what(merged)}，共 {sum(s['rows'] for s in merged)} 行"]
    left = [s for s in st["sources"] if s["status"] != "合并"]
    if left:
        lines.append(f"没合进去的（{len(left)}）：")
        lines += [f"  {s['file']}" + (f"「{s['sheet']}」" if s["sheet"] else "") + f"：{s['status']}" for s in left]
    short = [s for s in merged if s.get("missing")]
    if short:
        lines.append(f"缺列的（{len(short)}，缺的列留空）：")
        lines += [f"  {s['file']}：缺「" + "」「".join(s["missing"]) + "」" for s in short]
    more = [s for s in merged if s.get("extra")]
    if more:
        lines.append(f"多出列的（{len(more)}，多的列放在最后）：")
        lines += [f"  {s['file']}：多「" + "」「".join(s["extra"]) + "」" for s in more]
    totals = sum(s.get("totals", 0) for s in st["sources"])
    if totals:
        lines.append(f"各表自己的合计 / 小计行没合进来（{totals} 行）")
    fixed = sum(d["n"] for d in st.get("coerced", {}).values())
    if fixed:
        lines.append(f"按填写要求规整了 {fixed} 格写法（日期、数字、号码、选项），只改了汇总表，例子见「{REPORT_SHEET}」")
    if st.get("uncoerced"):
        lines.append(f"有 {st['uncoerced']['n']} 格认不出写法、没敢改，标了浅橙")
    if st["dupes"]:
        lines.append(f"内容完全相同、只合了一份（{len(st['dupes'])}）：" + "、".join(d["file"] for d in st["dupes"]))
    if st["errors"]:
        lines.append(f"读不了的（{len(st['errors'])}）：")
        lines += [f"  {e['file']}：{e['error']}" for e in st["errors"]]
    lines.append(f"每行最后一列「{SOURCE_COL}」写着来自哪里；最后一个 sheet「{REPORT_SHEET}」逐个列出明细。")
    return "\n".join(lines) + "\n"
