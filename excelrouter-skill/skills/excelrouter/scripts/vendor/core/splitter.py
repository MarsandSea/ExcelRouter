"""
Excel 智能拆分工具 - 核心拆分逻辑

把一批 Excel 表格按「某一列的取值」拆成多个文件，并完整保留原始表头格式。
面向「真实世界脏数据」：表头不在第一行、多 sheet、xls、公式列、单元格样式各异等都尽量稳。

输入可为单个文件或整个目录；一次运行产出两套结果：
  · 汇总（始终）：按 split_column 拆。单文件/merge=True → 扁平合并；目录+merge=False → 原文件拆分。
  · 到人（可选）：to_person 开 + 有 person_column + 文件名命中 person_file_filter 时，
    再按 person_column 拆到人（同一人跨文件合并到一个文件）。
拆分列（通用，不绑定任何业务）：
  · split_column     —— 主拆分列，必填，按列名选择
  · person_column    —— 到人列，选填
  · selected_values  —— 选填，只拆这些主取值；留空则自动枚举该列所有取值

表头识别（由 config['header_mode'] 控制）：
  · auto    —— 通用启发式自动找表头行（默认）
  · row     —— 用户指定 header_row（1 基行号）
  · keyword —— 旧关键词法：表头需同时含 grid_keys + id_keys（专家兜底）

数据行处理（由 config['preserve_format'] 控制）：
  · 保留格式（默认）：逐行复制值 + 完整单元格格式（字体/颜色/边框/数字格式）
  · 快速模式：数据行只写值，不处理格式，最快

【跨文件合并】同一个取值若出现在多个源文件里，会**追加合并到同一个输出文件**
（config['merge_across_files']，默认 False，界面上叫「跨文件合并汇总」），从根本上修掉旧版「同名文件互相覆盖」的 bug。
实现：输出工作簿在内存里按取值累积、最后统一保存；第一个产生该取值的源文件定义表头/格式，
后续文件按列名对齐追加（v3.0 起；旧版按列位置，列顺序不同的表会错位）：列的顺序不同、多几列少几列都能对上，
对上的列名不到一半的不硬合，单独放进「{sheet}_2」，见 _OutputBook.place。

【读表】v3.0 起只用 openpyxl：每个源文件完整加载一次（data_only=True，值和格式都从这一份取），
.xls 用 xlrd 读成内存工作簿，不再落临时文件；每一行都按【原始 Excel 行号】处理，按取值一次分好桶。
不再依赖 pandas（体积、安装、某些嵌入式 Python 加载 pandas 会卡死）。

【重要】按【原始行号】从源表读取、加法式写入输出，绝不删行——
从根本上规避历史上「删行后用错位行号索引」的 bug。

Copyright (c) 2026 AbeLin
MIT License
"""

import os
import re
import json
import time
import threading
import warnings
import datetime
from copy import copy
from itertools import chain, islice

import openpyxl
from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.formula.translate import Translator

from core.utils import soft_clean, hard_clean, apply_alias, is_skip_value, safe_filename
# 表头识别在 core/xlio.py（汇总也用它）；这里照旧导出原名，collect / 技能 / 测试都按 splitter 的名字调用
from core.xlio import resolve_header_row
from core.xlio import _looks_numeric, detect_header_row_auto, find_header_row  # noqa: F401
from core.xlio import (_head_limit, _iter_sheets_ro, cell_value, col_keys, col_names, hidden_info, is_total_row,
                       key_value, match, open_book, put, zip_folder)

warnings.filterwarnings('ignore')

# 拆分字段没写（或只有「-」「/」这类符号）、但这一行有内容：放进输出根目录的这个文件让用户自己核对，不再悄悄丢掉。
# 内部版 er_lite 用同一个文件名。
UNASSIGNED = "请核对_无归属.xlsx"


# =====================================================
# 未计算公式检测（防止公式列静默变空白）
# =====================================================

def detect_uncalculated_formulas(work_path, max_rows=200):
    """检测文件是否含「未计算的公式」：有公式(=开头)但没有缓存值。

    真实 Excel 文件保存时会缓存公式计算结果，拆分读到的就是这个值；若文件由程序生成或
    公式从未被 Excel 计算过，缓存值为空，拆分后这些列会变成空白且用户不会察觉。
    只采样前 max_rows 行，避免对大文件做双倍全量加载。返回 True 表示存在风险。
    """
    try:
        wb_f = openpyxl.load_workbook(work_path, read_only=True, data_only=False)
        wb_v = openpyxl.load_workbook(work_path, read_only=True, data_only=True)
        risky = False
        for sn in wb_f.sheetnames:
            if sn not in wb_v.sheetnames:
                continue
            ws_f, ws_v = wb_f[sn], wb_v[sn]
            ws_f.reset_dimensions()     # 尺寸写错的表，不重置只读模式会少读行列
            ws_v.reset_dimensions()
            count = 0
            for row_f, row_v in zip(ws_f.iter_rows(), ws_v.iter_rows()):
                for cf, cv in zip(row_f, row_v):
                    if isinstance(cf.value, str) and cf.value.startswith('=') and cv.value is None:
                        risky = True
                        break
                if risky:
                    break
                count += 1
                if count >= max_rows:
                    break
            if risky:
                break
        wb_f.close()
        wb_v.close()
        return risky
    except Exception:
        return False


# =====================================================
# 长耗时阻塞调用的"仍在运行"心跳（openpyxl.load_workbook 无进度回调）
# =====================================================

class _Ticker:
    """在一段无法拆分、无进度回调的阻塞调用（整份 openpyxl 读取）期间，
    定时打一条"仍在处理"日志，避免宽表/大文件读取几十秒甚至更久时，界面在这段
    空窗期里毫无变化，被用户误判为卡死（实际只是单线程还在读，见 CLAUDE.md 用户反馈）。
    用独立守护线程纯计时打日志，不碰被计时的调用本身，调用结束后 stop 即可。
    """

    def __init__(self, log_fn, label, first_delay=4.0, interval=7.0):
        self._log_fn = log_fn
        self._label = label
        self._first_delay = first_delay
        self._interval = interval
        self._stop_evt = threading.Event()
        self._thread = None

    def __enter__(self):
        if self._log_fn:
            self._thread = threading.Thread(target=self._run, daemon=True)
            self._thread.start()
        return self

    def _run(self):
        elapsed = 0.0
        delay = self._first_delay
        while not self._stop_evt.wait(delay):
            elapsed += delay
            self._log_fn(f"    …仍在{self._label}，已等待 {int(elapsed)} 秒"
                          f"（文件较大/列较多时属正常现象，不是卡死，请继续等待）")
            delay = self._interval

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._stop_evt.set()
        if self._thread is not None:
            self._thread.join(timeout=0.5)
        return False


# =====================================================
# 样式缓存（保留格式模式提速的关键）
# =====================================================

def _copy_cell_style(src, dst, cache):
    """把 src 单元格的格式（字体、填充、边框、对齐、数字格式）复制到 dst，按「源单元格的样式编号」缓存。

    src._style 是源工作簿里的样式编号数组（字体第几个、填充第几个……），同一个源文件里
    编号相同 = 格式相同；dst 第一次拿到这套格式后，它在目标工作簿里的编号数组也就定了，
    之后同样格式的格子直接套用这个数组，不再逐项复制、比较样式对象（快一个数量级）。
    缓存键带上目标工作簿本身（同一个源文件会写进多个输出工作簿，编号各不相同）。
    缓存按「单个源文件」作用域使用：换源文件就换一个新 dict。

    ⚠️ 历史 bug（v2.10.0 及以前）：旧的 _cached_copy 用 id(src.font) 当键，但 src.font 每次访问
    都返回一个新的临时代理对象——大多数时候缓存不命中（慢），代理被回收后地址被下一个复用时
    又会命中别的格子的样式，实测 900 个格子里有 11～20 个填充色 / 字体被复制错。别改回按 id 缓存。
    """
    key = (dst.parent.parent, tuple(src._style))
    arr = cache.get(key)
    if arr is None:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        cache[key] = copy(dst._style)
    else:
        dst._style = copy(arr)


def _cached_copy(style_obj, cache):
    """（已弃用，保留给外部调用方）按 id 缓存复制样式对象。

    不要再用：style_obj 若是 cell.font 这类每次新建的代理对象，id 会被复用而取到别的样式，
    见 _copy_cell_style 的说明。
    """
    k = id(style_obj)
    cached = cache.get(k)
    if cached is None:
        cached = copy(style_obj)
        cache[k] = cached
    return cached


# =====================================================
# 公式智能保留（keep_formulas，仅同行公式安全平移）
# =====================================================

_SAFE_REF_RE = re.compile(r'^(\$?)([A-Z]{1,3})(\$?)(\d+)$')


def _is_row_safe_ref(ref, origin_row):
    """判断单个引用是否「同表、单格、相对/绝对行都指向 origin_row」。

    引用带 '!'（跨表）或 ':'（区域）或格式不是「列字母+行号」（命名区域/整列/整行）都不安全。
    行号本身若不等于 origin_row（跨行引用），也不安全——拆分后该行已被过滤或搬到别处。
    """
    if '!' in ref or ':' in ref:
        return False
    m = _SAFE_REF_RE.match(ref)
    if not m:
        return False
    row = int(m.group(4))
    return row == origin_row


def _safe_translate(formula, origin, dest):
    """公式安全性校验 + 行号平移。

    只有当公式里所有单元格引用都满足 _is_row_safe_ref（同行、同表、单格）时，才认为
    这个公式「整行搬到哪都还对」，用 openpyxl 的 Translator 把行号从 origin 平移到 dest
    并返回新公式；否则返回 None，表示不安全——调用方应保留已经写入的缓存数值，不覆盖成公式。
    任何解析异常都视为不安全，退回 None（宁可落成数值，不可能给出算错的公式）。
    """
    try:
        origin_row = int(re.match(r'^\$?[A-Za-z]{1,3}\$?(\d+)$', origin).group(1))
        tok = Tokenizer(formula)
        for t in tok.items:
            if t.type == 'OPERAND' and t.subtype == 'RANGE':
                if not _is_row_safe_ref(t.value, origin_row):
                    return None
        return Translator(formula, origin=origin).translate_formula(dest)
    except Exception:
        return None


# =====================================================
# 列名 / 取值枚举（供 GUI 下拉与多选，也给技能用）
# =====================================================

def _trim_tail(rows):
    """去掉尾部的空行：尾部只有格式、没有内容的行不算数（和旧版 pandas 读表的口径一样）。"""
    rows = list(rows)
    while rows and all(v is None or v == "" for v in rows[-1]):
        rows.pop()
    return rows


def _key_row(row):
    """一行的值 → 识别表头、比列名用的值：错误格当空，3.0 当 3。"""
    return tuple(key_value(v) for v in row)


def _header_names(row):
    """表头行 → 列名列表（soft_clean 后）。只看这一行，和旧版一致。"""
    return [soft_clean(key_value(v)) for v in row]


def _find_col(names, column):
    """按列名定位 0 基列号：先精确匹配，再包含匹配。找不到返回 None。"""
    target = soft_clean(column)
    if not target:
        return None
    for c, name in enumerate(names):
        if name == target:
            return c
    for c, name in enumerate(names):
        if name and (target in name or name in target):
            return c
    return None


def _hidden_sheets(file_path, config):
    """不发出的隐藏 sheet 名集合：keep_hidden 打开时为空。"""
    if config.get('keep_hidden', False):
        return set()
    return {sn for sn, (state, _cols) in hidden_info(file_path).items() if state != "visible"}


def _sheet_heads(file_path, config):
    """只读方式逐 sheet 产出 (sheet 名, 表头行号或 -1, 列名, 表头附近的行, 其余行的迭代器)。

    表头只在前几行里找；数据行先是「表头附近的行」里表头之后的部分，再接着从迭代器里取。
    隐藏的 sheet 不拆，这里也不列出来（keep_hidden 打开时照常列出）。
    调用方中途不读了（break / return）也会关掉源文件。
    """
    limit = _head_limit(config)
    hidden = _hidden_sheets(file_path, config)
    sheets = _iter_sheets_ro(file_path)
    try:
        for sn, it in sheets:
            if sn in hidden:
                continue
            head = _trim_tail(_key_row(r) for r in islice(it, limit))
            h = resolve_header_row(head, config)
            names = _header_names(head[h - 1]) if 0 < h <= len(head) else []
            yield sn, h, names, head, it
    finally:
        sheets.close()


def sheet_headers(file_path, config):
    """每个 sheet 的表头：[(sheet 名, 表头行号或 -1, [列名])]。只读每个 sheet 前几行，快（技能用）。"""
    return [(sn, h, [c for c in names if c] if 0 < h <= len(head) else [])
            for sn, h, names, head, _it in _sheet_heads(file_path, config)]


def list_columns(file_path, config, log_fn=None):
    """返回首个可识别 sheet 的列名列表（供 GUI 主/二级列下拉）。只读前若干行，快。"""
    for _sn, h, names, head, _it in _sheet_heads(file_path, config):
        if 0 < h <= len(head):
            return [c for c in names if c]
    return []


def list_values(file_path, config, column, log_fn=None):
    """返回某列去重、归并、去跳过值后的取值列表（供 GUI「只拆这些值」多选）。"""
    return scan_values(file_path, config, column)[0]


def scan_values(file_path, config, column):
    """读一遍文件，返回 (取值列表, 这个文件的模板)。取值同 list_values。

    模板是第一个含这一列的 sheet 的列名（多行表头时空着的往上找，同跨文件对齐），交给 template_groups
    看一批表是不是同一种模板；文件里没有这一列时为 None。GUI「扫描全部分组」用它，一遍读完两样。
    """
    alias = config.get('value_alias_map', {})
    skip = config.get('skip_values', [])
    values, tpl = set(), None
    for _sn, h, names, head, it in _sheet_heads(file_path, config):
        if not 0 < h <= len(head):
            continue
        ci = _find_col(names, column)
        if ci is None:
            continue
        if tpl is None:
            tpl = col_names(head, h)
        for row in chain(head[h:], it):
            nv = apply_alias(key_value(row[ci]) if ci < len(row) else None, alias)
            if not is_skip_value(nv, skip):
                values.add(nv)
    return sorted(values), tpl


def template_groups(pairs):
    """按表头给一批文件分组：[(文件, 列名或 None)] → [{'names': 列名, 'files': [文件, ...]}]，文件多的组在前。

    列名一样（不计顺序；比法同跨文件对齐，只看字母数字汉字）就是同一种模板。列名为 None 的文件
    （读不了、没有拆分字段）不计入。跨文件合并现在按列名对齐，几种模板也能合，但用户该知道表其实不一样。
    """
    groups = {}
    for f, names in pairs:
        if names is None:
            continue
        sig = frozenset(k for k in col_keys(names) if not k.startswith('#'))
        g = groups.get(sig)
        if g is None:
            groups[sig] = {'names': list(names), 'files': [f]}
        else:
            g['files'].append(f)
    return sorted(groups.values(), key=lambda g: -len(g['files']))


def _quote(names, k=4):
    return "".join(f"「{n}」" for n in names[:k]) + ("等" if len(names) > k else "")


def template_lines(groups):
    """几种模板 → 说明行：每种几个表、举一个表名，和最多的那种比多了 / 少了哪些列。"""
    main = {k: n for k, n in zip(col_keys(groups[0]['names']), groups[0]['names']) if not k.startswith('#')}
    lines = []
    for g in groups:
        mine = {k: n for k, n in zip(col_keys(g['names']), g['names']) if not k.startswith('#')}
        more = [n for k, n in mine.items() if k not in main]
        less = [n for k, n in main.items() if k not in mine]
        diff = "；".join(([f"多{_quote(more)}"] if more else []) + ([f"少{_quote(less)}"] if less else []))
        lines.append(f"  · {len(g['files'])} 个和「{os.path.basename(g['files'][0])}」一样"
                     + (f"：{diff}" if diff else ""))
    return lines


# =====================================================
# 表头格式读取 / 写入
# =====================================================

def _read_header_format(ws, h_idx):
    """从源 sheet 读取表头区域（1..h_idx 行）的值与格式。返回 (cells, merges, widths)。"""
    cells = []
    for r in range(1, h_idx + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            is_m = isinstance(cell, MergedCell)
            cells.append({
                'row': r, 'col': c,
                'value': None if is_m else cell.value,
                'is_merged': is_m,
                'font':          copy(cell.font)      if not is_m and cell.has_style else None,
                'fill':          copy(cell.fill)      if not is_m and cell.has_style else None,
                'border':        copy(cell.border)    if not is_m and cell.has_style else None,
                'alignment':     copy(cell.alignment) if not is_m and cell.has_style else None,
                'number_format': cell.number_format   if not is_m else None,
            })
    merges = [str(mr) for mr in ws.merged_cells.ranges if mr.max_row <= h_idx]
    widths = {col: ws.column_dimensions[col].width for col in ws.column_dimensions}
    return cells, merges, widths


def _write_header(ws_out, header_cells, merges, widths):
    """把表头的值 + 格式 + 合并单元格 + 列宽写入输出 sheet。"""
    for ci in header_cells:
        if ci['is_merged']:
            continue
        dst = ws_out.cell(row=ci['row'], column=ci['col'])
        put(dst, ci['value'])      # 表头里以「=」开头的文字照样是文字
        if ci['font']:          dst.font          = ci['font']
        if ci['fill']:          dst.fill          = ci['fill']
        if ci['border']:        dst.border        = ci['border']
        if ci['alignment']:     dst.alignment     = ci['alignment']
        if ci['number_format']: dst.number_format = ci['number_format']
    for mr_str in merges:
        ws_out.merge_cells(mr_str)
    for col_letter, width in widths.items():
        if width is not None:
            ws_out.column_dimensions[col_letter].width = width


def _copy_header_cell(ws_src, h_src, c, ws_out, h_out, d, name):
    """跨文件按列名对齐时，来源多出来的一列：在输出表头行第 d 列写上列名，格式和列宽照来源第 c 列。"""
    dst = ws_out.cell(row=h_out, column=d)
    put(dst, name)
    src = ws_src.cell(row=h_src, column=c)
    if not isinstance(src, MergedCell) and src.has_style:
        dst.font, dst.fill = copy(src.font), copy(src.fill)
        dst.border, dst.alignment = copy(src.border), copy(src.alignment)
        dst.number_format = src.number_format
    dim = ws_src.column_dimensions.get(get_column_letter(c))
    if dim is not None and dim.width is not None:
        ws_out.column_dimensions[get_column_letter(d)].width = dim.width


# =====================================================
# 下拉 / 数据验证：照源表带到每个产出
# =====================================================

# 判断两个数据验证是不是同一个时比的属性（公式另算）
_DV_ATTRS = ('type', 'operator', 'allow_blank', 'showDropDown', 'showErrorMessage', 'showInputMessage',
             'errorTitle', 'error', 'promptTitle', 'prompt', 'errorStyle', 'imeMode')
_DV_LIST_MAX = 255       # Excel 写死的下拉选项总长度上限（字面列表）
_DV_RANGE_MAX = 5000     # 引用的选项区域最多看这么多格，再大就不内联


def _formula_ref(f):
    """数据验证公式里引用了格子吗：返回 (有引用, 整个公式只是一个区域时的区域文字)。解析不了按「有引用」处理。"""
    try:
        items = Tokenizer(f if f.startswith('=') else '=' + f).items
    except Exception:
        return True, None
    refs = [t for t in items if t.type == 'OPERAND' and t.subtype == 'RANGE']
    if not refs:
        return False, None
    only = refs[0].value if len(items) == 1 else None
    return True, only


def _list_from_range(wb, ws, ref):
    """下拉选项引用的区域（同表、别的 sheet、或定义的名称）→ 字面列表 '"甲,乙"'；读不出、太长、含英文逗号时返回 None。

    只读格子不建格子（用 _cells），不改动源表的尺寸。
    """
    target, rng = ws, ref
    if '!' in ref:
        sheet, rng = ref.rsplit('!', 1)
        sheet = sheet.strip("'").replace("''", "'")
        target = wb[sheet] if sheet in wb.sheetnames else None
    else:
        dn = ws.defined_names.get(ref) if hasattr(ws, 'defined_names') else None
        dn = dn or wb.defined_names.get(ref)
        if dn is not None:
            dests = list(dn.destinations)
            if len(dests) != 1 or dests[0][0] not in wb.sheetnames:
                return None
            target, rng = wb[dests[0][0]], dests[0][1]
    if target is None:
        return None
    try:
        c1, r1, c2, r2 = range_boundaries(rng.replace('$', ''))
    except Exception:
        return None
    r1, c1 = r1 or 1, c1 or 1
    r2 = min(r2 or target.max_row, target.max_row)
    c2 = min(c2 or target.max_column, target.max_column)
    if r2 < r1 or c2 < c1 or (r2 - r1 + 1) * (c2 - c1 + 1) > _DV_RANGE_MAX:
        return None
    vals = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = target._cells.get((r, c))
            v = key_value(cell_value(cell)) if cell is not None else None
            t = "" if v is None else str(v).strip()
            if t and t not in vals:
                vals.append(t)
    if not vals or any(',' in t or '"' in t for t in vals):
        return None
    lit = ','.join(vals)
    return f'"{lit}"' if len(lit) <= _DV_LIST_MAX else None


def _source_validations(wb, ws, h):
    """源 sheet 的数据验证 → (表头区, 数据区, 去掉了几个, 内联了几个)。

    表头区 / 数据区各是 [(签名, 模板, [(起行, 止行, 起列, 止列)])]，表头区只含第 1～h 行，数据区只含 h 行以下。
    · 公式里不引用格子的（字面下拉、数值 / 日期范围）照搬；
    · 下拉选项引用本工作簿一片格子的：把选项直接写进下拉（不超过 255 个字、不含英文逗号），否则去掉；
    · 其余引用格子的（自定义公式、按别的格子算范围）去掉——拆出去以后引用的格子对不上了。
    """
    header, data, dropped, inlined = [], [], 0, 0
    for dv in ws.data_validations.dataValidation:
        f1 = dv.formula1
        if dv.formula2 and _formula_ref(dv.formula2)[0]:
            dropped += 1
            continue
        if f1:
            has_ref, only = _formula_ref(f1)
            if has_ref:
                lit = _list_from_range(wb, ws, only) if (dv.type == 'list' and only) else None
                if lit is None:
                    dropped += 1
                    continue
                f1 = lit
                inlined += 1
        tpl = copy(dv)
        tpl.formula1 = f1
        sig = (f1, dv.formula2) + tuple(getattr(dv, a, None) for a in _DV_ATTRS)
        hr, dr = [], []
        for cr in dv.sqref.ranges:
            c1, r1, c2, r2 = cr.bounds
            if r1 <= h:
                hr.append((r1, min(r2, h), c1, c2))
            if r2 > h:
                dr.append((max(r1, h + 1), r2, c1, c2))
        if hr:
            header.append((sig, tpl, hr))
        if dr:
            data.append((sig, tpl, dr))
    return header, data, dropped, inlined


def _note_validations(sinfo, entries, src_index, start, col_of, max_col):
    """把源表 0 基行号 src_index[i]（Excel 第 src_index[i]+1 行）的数据验证记到产出第 start+i 行
    （列按 col_of 换算），保存时统一写（见 _flush_validations）。"""
    if not entries:
        return
    store = sinfo.setdefault('dv', {})
    for sig, tpl, ranges in entries:
        cols = store.setdefault(sig, (tpl, {}))[1]
        for r1, r2, c1, c2 in ranges:
            hit = [start + i for i, si in enumerate(src_index) if r1 <= si + 1 <= r2]
            if not hit:
                continue
            for c in range(c1, min(c2, max_col) + 1):
                cols.setdefault(col_of(c), set()).update(hit)


def _flush_validations(info):
    """把一个产出 sheet 记下的数据验证写进去：同一个验证一列里连着的行合成一段。"""
    for _sig, (tpl, cols) in info.get('dv', {}).items():
        parts = []
        for d in sorted(cols):
            letter, rows = get_column_letter(d), sorted(cols[d])
            a = prev = rows[0]
            for r in rows[1:] + [None]:
                if r is not None and r == prev + 1:
                    prev = r
                    continue
                parts.append(f"{letter}{a}" if a == prev else f"{letter}{a}:{letter}{prev}")
                if r is not None:
                    a = prev = r
        if parts:
            dv = copy(tpl)
            dv.sqref = MultiCellRange(" ".join(parts))
            info['ws'].add_data_validation(dv)


# =====================================================
# 文件名 / sheet 名去重与净化
# =====================================================

_INVALID_SHEET_CHARS = re.compile(r'[\[\]:*?/\\]')


def _safe_sheet_title(name, used_titles):
    """把任意 sheet 名转成合法且在本工作簿内唯一的 Excel sheet 标题。

    Excel 规则：禁止 [ ] : * ? / \\，不能为空，最长 31 字符，且不能重复。
    used_titles 收集已用的小写标题；重名时在末尾追加 (2)/(3)... 并保证不超 31 字符。
    """
    t = _INVALID_SHEET_CHARS.sub('_', soft_clean(name))[:31] or "Sheet"
    base = t
    i = 2
    while t.lower() in used_titles:
        suffix = f"({i})"
        t = base[:31 - len(suffix)] + suffix
        i += 1
    used_titles.add(t.lower())
    return t


def _dedupe_path(save_path, used_paths):
    """若目标路径已被占用，则在扩展名前追加 (2)/(3)... 直到不冲突。

    用于避免「两个不同取值净化后同名 → 同一 save_path → 后者覆盖前者」的数据丢失。
    used_paths 是已占用路径的规范化集合（os.path.normcase + abspath）。
    """
    norm = os.path.normcase(os.path.abspath(save_path))
    if norm not in used_paths:
        return save_path
    root, ext = os.path.splitext(save_path)
    i = 2
    while True:
        cand = f"{root}({i}){ext}"
        if os.path.normcase(os.path.abspath(cand)) not in used_paths:
            return cand
        i += 1


# =====================================================
# 输出工作簿（按取值在内存中累积，最后统一保存）
# =====================================================

class _OutputBook:
    """一个输出文件对应的内存工作簿，支持跨源文件追加。"""

    __slots__ = ('save_path', 'wb', 'sheets', '_used_titles')

    def __init__(self, save_path):
        self.save_path = save_path
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.sheets = {}        # sheet_name -> {'ws':, 'next_row':, 'keys':, 'names':, 'h':, 'base':}（后四项见 place）
        self._used_titles = set()

    def get_or_create_sheet(self, sheet_name, h_idx, header_cells, merges, widths):
        """取得（或首次建立并写入表头的）输出 sheet 信息。"""
        info = self.sheets.get(sheet_name)
        if info is None:
            ws = self.wb.create_sheet(title=_safe_sheet_title(sheet_name, self._used_titles))
            _write_header(ws, header_cells, merges, widths)
            info = {'ws': ws, 'next_row': h_idx + 1}
            self.sheets[sheet_name] = info
        return info

    def place(self, sheet_name, h_idx, header, keys, names, ws_src):
        """给一个来源 sheet 找输出 sheet，返回 (sheet 信息, 列映射或 None, 换到的 sheet 名或 None, 对齐说明或 None)。

        · 还没有同名 sheet：用这个来源的表头建一个；
        · 同名 sheet 的列键和这个来源完全一样：照原样按位置追加（列映射为 None）；
        · 列不完全一样、但这个来源的列名有一半以上在输出里：按列名对齐，缺的列留空，多出的列加在最后；
          各列位置其实没变（只是有的表多几列、少几列）时列映射仍为 None，照原样按位置写；
        · 对上的列名不到一半：不硬合，换到「{sheet}_2」（被占用又对不上就 _3…）。拆分丢行就是丢数据，
          所以不能像汇总那样整个不合。
        keys / names 是这个来源 sheet 每一列的列键和列名（没有列名的列，列键按位置记成「#列号」）。
        """
        name, i = sheet_name, 1
        while True:
            info = self.sheets.get(name)
            moved = name if i > 1 else None
            if info is None:
                info = self.get_or_create_sheet(name, h_idx, *header)
                info.update(keys=list(keys), names=list(names), h=h_idx, base=len(keys))
                return info, None, moved, None
            out_keys = info.get('keys')
            if out_keys is None or out_keys == list(keys) or all(k.startswith('#') for k in keys):
                return info, None, moved, None
            if not match(out_keys, keys)[1]:
                col_map, detail = self._align(info, keys, names, ws_src, h_idx)
                if all(d == c for c, d in enumerate(col_map, start=1)):
                    col_map = None
                return info, col_map, moved, detail
            i += 1
            name = f"{sheet_name}_{i}"

    @staticmethod
    def _align(info, keys, names, ws_src, h_src):
        """按列名把来源的每一列对到输出 sheet 的列上：返回 (列映射, 说明或 None)。

        列映射第 c-1 项是来源第 c 列要写到的输出列号。来源多出的列加在输出的最后，
        有列名的照来源写上表头（格式、列宽也照来源）。说明里是：比建这个 sheet 的那个表少了哪些列、
        多了哪些列，已有的列顺序变没变；都没有时为 None。
        """
        out_keys, out_names = info['keys'], info['names']
        pos = {k: c for c, k in enumerate(out_keys, start=1)}
        old = len(out_keys)
        ws_out, h_out = info['ws'], info['h']
        col_map, extra = [], []
        for c, k in enumerate(keys, start=1):
            d = pos.get(k)
            if d is None:
                out_keys.append(k)
                out_names.append(names[c - 1])
                d = pos[k] = len(out_keys)
                if not k.startswith('#'):
                    _copy_header_cell(ws_src, h_src, c, ws_out, h_out, d, names[c - 1])
                    extra.append(names[c - 1])
            col_map.append(d)
        have, base = set(keys), info.get('base', old)
        # 少的列只和建 sheet 的那个表比：别的表多出来、加在后面的列，这个表没有是自然的，不算少
        missing = [n for k, n in zip(out_keys[:base], out_names[:base]) if not k.startswith('#') and k not in have]
        named = [d for k, d in zip(keys, col_map) if not k.startswith('#') and d <= old]
        reordered = named != sorted(named)
        if not (missing or extra or reordered):
            return col_map, None
        return col_map, {'missing': missing, 'extra': extra, 'reordered': reordered}

    def save(self):
        """落盘。无任何 sheet 时不保存，返回 False。"""
        if not self.sheets:
            self.wb.close()
            return False
        for info in self.sheets.values():
            _flush_validations(info)
        os.makedirs(os.path.dirname(self.save_path), exist_ok=True)
        self.wb.save(self.save_path)
        self.wb.close()
        return True


class _Rows:
    """给 _append_rows 的行容器：index 是这些行在源表里的 0 基行号（Excel 行号减 1）。

    _append_rows 只用 .index 和 len()；内部 er_lite 也传同款容器，别改这个约定。
    """

    __slots__ = ('index',)

    def __init__(self, index):
        self.index = index

    def __len__(self):
        return len(self.index)


def _append_rows(sheet_info, rows_df, preserve, ws_src, src_max_col, style_cache,
                 heartbeat=None, formula_ws=None, col_map=None):
    """把选中的数据行追加到输出 sheet 的当前游标下。

    rows_df.index 是这些行在源表里的【原始 0 基行号】，excel_row = index + 1（只用 .index 和 len()，
    splitter 传 _Rows，内部 er_lite 传同款容器）。值和格式都从 ws_src 按原始行号取：
    preserve=True 时逐格复制「值 + 格式」；否则只写值（错误格和空格子不写）。
    写值一律走 put：源表里以「=」开头的文字照样是文字，不会变成能执行的公式。
    formula_ws 非 None 时（keep_formulas 开启且 preserve 有效），对每个公式单元格尝试
    「同行安全平移」：能平移则覆盖为新公式，不能则保留已写入的缓存数值（见 _safe_translate）。
    col_map 非 None 时（跨文件按列名对齐），源表第 c 列写到输出第 col_map[c-1] 列；这时调用方不传 formula_ws。
    heartbeat(n) 每写入一批行回调一次，用于向界面报进度并让出 GIL（保持界面响应）。
    """
    ws_out = sheet_info['ws']
    start = sheet_info['next_row']
    n = len(rows_df)
    # 心跳按「已处理单元格数」折算成行数阈值，而非固定行数：宽表（列数多）单元格
    # 级操作更贵，若仍按固定 200 行让一次 GIL，两次让出之间工作线程霸占 GIL 的时间会随列数
    # 线性变长，表现为列数越多界面越卡（甚至处理详情面板都打不开）。目标约 4000 格/次让出。
    batch = max(20, 4000 // max(1, src_max_col))
    cols = [(c, c if col_map is None else col_map[c - 1]) for c in range(1, src_max_col + 1)]
    if col_map is not None:
        formula_ws = None       # 列换了位置：同行公式按列字母引用，平移过去会算错，落成数值

    if ws_src is not None:
        for offset, src_idx in enumerate(rows_df.index):
            excel_row = int(src_idx) + 1
            out_row = start + offset
            if preserve:
                for c, d in cols:
                    src_cell = ws_src.cell(row=excel_row, column=c)
                    dst = ws_out.cell(row=out_row, column=d)
                    put(dst, src_cell.value)
                    if not isinstance(src_cell, MergedCell):
                        try:
                            if src_cell.has_style:
                                _copy_cell_style(src_cell, dst, style_cache)
                        except Exception:
                            pass
                    if formula_ws is not None:
                        fcell = formula_ws.cell(row=excel_row, column=c)
                        # 按单元格类型认公式：以「=」开头的文字不是公式，不能拿去平移
                        if fcell.data_type == 'f' and isinstance(fcell.value, str):
                            col_letter = get_column_letter(c)
                            translated = _safe_translate(
                                fcell.value, f"{col_letter}{excel_row}", f"{col_letter}{out_row}")
                            if translated is not None:
                                dst.value = translated
                            # 不安全：dst.value 保持上面已写入的缓存数值，不覆盖
            else:
                for c, d in cols:
                    v = cell_value(ws_src.cell(row=excel_row, column=c))
                    if v is not None and v != "":
                        put(ws_out.cell(row=out_row, column=d), v)
            if heartbeat and (offset + 1) % batch == 0:
                heartbeat(batch)
    if heartbeat and n % batch:
        heartbeat(n % batch)

    sheet_info['next_row'] = start + n


def _matches_filter(name, keywords):
    """文件名是否命中任一关键词；keywords 为空表示全部命中。"""
    if not keywords:
        return True
    n = soft_clean(name)
    return any(soft_clean(k) and soft_clean(k) in n for k in keywords)


def _summary_key_path(output_root, primary, src_stem, single_file, merge):
    """『汇总』输出键与路径。

    单文件 / merge=True → 扁平 {主取值}.xlsx（跨文件合并到一个）。
    文件夹 + merge=False → {主取值}/汇总/{主取值}_{原文件名}.xlsx（原文件拆分，不合并）。
    文件名带上主取值前缀：同一份源表拆到不同主取值的文件夹里，原本各自都叫
    「{原文件名}.xlsx」，一旦脱离文件夹（比如直接从各文件夹里把文件拖出来一起发）
    就分不清谁是谁，前缀是让文件名本身自带身份。
    """
    p = safe_filename(primary)
    if single_file or merge:
        return ('汇总', primary), os.path.join(output_root, f"{p}.xlsx")
    return ('汇总', primary, src_stem), os.path.join(output_root, p, "汇总", f"{p}_{src_stem}.xlsx")


def _person_key_path(output_root, primary, person, src_stem, merge):
    """『到人』输出键与路径，与『汇总』的 merge 语义对称（v2.7.2 修名实不符）：

    · merge=True（跨文件合并，把整个文件夹当一个整体）→ 键 ('到人', primary, person)，
      路径 {主取值}/到人/{姓名}.xlsx：同一人的全部行合并进这一张表，文件名不带
      源文件名——旧版在 merge 模式下文件名取第一个源文件（张三_7月.xlsx 里装着
      8/9 月的行），名字与内容不符，用户以为其他月份被覆盖丢了。
    · merge=False（按原表拆分）→ 键 ('到人', primary, person, src_stem)，
      路径 {主取值}/到人/{姓名}_{原文件名}.xlsx：每人每源一张，与汇总的
      {主取值}_{原文件名}.xlsx 对称；不同月份表结构独立时也避免按列位置合并错位。

    键含不含 src_stem 决定合并还是分表；src_stem 不再只影响文件名。
    """
    p = safe_filename(primary)
    s = safe_filename(person)
    if merge:
        return ('到人', primary, person), os.path.join(output_root, p, "到人", f"{s}.xlsx")
    return ('到人', primary, person, src_stem), os.path.join(
        output_root, p, "到人", f"{s}_{src_stem}.xlsx")


# =====================================================
# 单文件处理
# =====================================================

def _align_lines(sn, notes):
    """一个来源 sheet 跨文件按列名对齐的情况 → 日志行。notes：{情况: 涉及几个输出文件}。"""
    lines = []
    for nk, n in notes.items():
        if nk == 'formula':
            lines.append(f"  ⚠️ sheet「{sn}」的列和前面的表位置不一样，这部分行的公式落成了数值"
                         f"（公式按列字母引用，换了位置会算错；涉及 {n} 个输出文件）")
            continue
        moved, missing, extra, reordered = nk
        parts = (["列的顺序不一样"] if reordered else []) + \
                ([f"少{_quote(missing)}（留空）"] if missing else []) + \
                ([f"多{_quote(extra)}（加在最后）"] if extra else [])
        if moved:
            head = f"  ⚠️ sheet「{sn}」和前面同名 sheet 的列名大半对不上，没有硬合，单独放进「{moved}」"
        else:
            head = f"  ℹ️ sheet「{sn}」和前面同名 sheet 的列不完全一样，已按列名对齐"
        lines.append(head + (f"：{'；'.join(parts)}" if parts else "") + f"（涉及 {n} 个输出文件）")
    return lines


def _clear_hidden(ws, h, hcols):
    """把内存里源 sheet 隐藏列从表头行起的值清空（源文件不落盘，不受影响），复制出去的产出里就没有这些内容。

    表头上方的标题行不动；跨到可见列的合并单元格（如大标题）不清，免得把看得见的内容一起清掉。
    与内部版 er_lite 的同名函数同一规则。
    """
    keep = {(mr.min_row, mr.min_col) for mr in ws.merged_cells.ranges
            if any(c not in hcols for c in range(mr.min_col, mr.max_col + 1))}
    for r in range(h, ws.max_row + 1):
        for c in hcols:
            cell = ws._cells.get((r, c))
            if cell is not None and not isinstance(cell, MergedCell) and (r, c) not in keep:
                cell.value = None


def _hidden_data_rows(ws, h, last):
    """表头以下、有内容的隐藏行（多半是筛选掉的）有几行。.xls 读成的内存工作簿没有行属性，算 0。"""
    hid = {r for r, dim in list(ws.row_dimensions.items()) if dim.hidden and h < r <= last}
    if not hid:
        return 0
    return len({r for (r, _c), cell in ws._cells.items() if r in hid and cell.value not in (None, "")})


def _content_rows(ws, h, last):
    """表头以下（第 h+1 ~ last 行）有内容的行号集合：只有空格的不算。"""
    rows = set()
    for (r, _c), cell in ws._cells.items():
        if h < r <= last and r not in rows:
            v = cell.value
            if v is not None and (not isinstance(v, str) or soft_clean(v) != ""):
                rows.add(r)
    return rows


def _looks_like_total(ws, r, end):
    """没写分组的这一行其实是合计行：前几格写着「合计 / 小计 / 总计」，或者是表里最后一行、只有数字没有文字。

    这类行算进「合计等」，不放进「请核对_无归属」——不然几乎每张带合计行的表都要多一个请核对文件。
    """
    row = next(ws.iter_rows(min_row=r, max_row=r, values_only=True), ())
    if is_total_row(row):
        return True
    vals = [v for v in row if v is not None and soft_clean(v) != ""]
    return r == end and bool(vals) and all(_looks_numeric(v) for v in vals)


def _review_sheet(review, rel_path, sn, title, h, header):
    """「请核对_无归属」里给一个源 sheet 的那一页（第一次用时建，表头照源表）。"""
    info = review.sheets.get((rel_path, sn))
    if info is None:
        ws = review.wb.create_sheet(title=_safe_sheet_title(title, review._used_titles))
        _write_header(ws, *header)
        info = review.sheets[(rel_path, sn)] = {'ws': ws, 'next_row': h + 1}
    return info


def _last_data_row(ws):
    """最后一行有内容的行号；尾部只有格式、没有内容的行不算（和旧版 pandas 读表的口径一样）。整页空返回 0。"""
    last = 0
    for (r, _c), cell in ws._cells.items():
        if r > last and cell.value is not None and cell.value != "":
            last = r
    return last


def _column_values(ws, col, first_row, last_row):
    """第 col 列（1 基）从 first_row 到 last_row 的值。"""
    return [row[0] for row in ws.iter_rows(min_row=first_row, max_row=last_row,
                                           min_col=col, max_col=col, values_only=True)]


def process_file(file_path, rel_path, output_root, config, outputs,
                 single_file=False, log_fn=None, stop_flag=None, tick_fn=None, stats=None, review=None):
    """处理单个文件：始终产出『汇总』；满足条件时附加产出『到人』。返回写入总行数。

    stats 为可选的可变 dict（由 run_split 传入），用于跨文件累计「跳过 sheet 数 /
    失败文件数」，供收尾的机器可读摘要行使用；为 None 时只记日志不累计。

    · 汇总：按 split_column 拆。单文件/merge=True → 扁平合并；文件夹+merge=False → 原文件拆分。
    · 到人：仅当 to_person 开、有 person_column、且文件名命中 person_file_filter 时，
      按 split_column → person_column 拆，同一人跨文件合并到一个文件。
    · tick_fn(0..1) 报告本文件内部进度（读取→拆分各阶段），供界面进度条使用。
    · review 为「请核对_无归属」的 _OutputBook（run_split 建）：拆分字段没写、但有内容的行写进去，
      一个源 sheet 一页；为 None 时这些行只计数。
    · stats 里累计行数对账：source_rows（有内容的数据行）= routed（进了至少一个分组）+ skip_rows（合计等）
      + unassigned（没写分组）+ not_selected（选了只拆部分取值时没选上的）。隐藏的、跳过的 sheet 不计。

    读表：源文件只完整加载一次（open_book），值和格式都从这一份取；每个 sheet 先把拆分列读一遍、
    按取值分好桶（取值 → 原始行号），再按桶输出。输出顺序与旧版逐取值筛选完全一致。
    """
    preserve_format = config.get('preserve_format', True)
    keep_formulas   = config.get('keep_formulas', False)
    split_column    = soft_clean(config.get('split_column', ''))
    person_column   = soft_clean(config.get('person_column', ''))
    exact           = config.get('exact_match', True)
    alias           = config.get('value_alias_map', {})
    skip            = config.get('skip_values', [])
    merge           = config.get('merge_across_files', True)
    keep_hidden     = config.get('keep_hidden', False)
    selected = [v for v in (apply_alias(x, alias) for x in config.get('selected_values', [])) if v]
    selected_set = set(selected)

    if not split_column:
        if log_fn:
            log_fn("  ⚠️ 未指定拆分字段，跳过")
        return 0

    base_name = os.path.basename(file_path)
    do_person = (config.get('to_person', False) and not single_file and bool(person_column)
                 and _matches_filter(base_name, config.get('person_file_filter', [])))

    is_xls = base_name.lower().endswith('.xls')
    src_stem = safe_filename(os.path.splitext(base_name)[0], "source")

    wb_src = None
    wb_formula = None

    def tick(frac):
        """报告本文件内部进度（0~1），供界面进度条平滑推进。"""
        if tick_fn:
            tick_fn(min(max(frac, 0.0), 1.0))

    try:
        tick(0.02)
        if log_fn:
            log_fn("  ⏳ 正在读取数据…（行数多的大文件这一步最慢，请耐心等待）")
        if not is_xls and detect_uncalculated_formulas(file_path):
            if log_fn:
                log_fn("  ⚠️ 注意：该文件含未计算的公式，相关列可能为空白，"
                       "建议用 Excel 打开另存后再处理")
        tick(0.10)
        preserve_ok = preserve_format and not is_xls
        if preserve_format and is_xls and log_fn:
            log_fn("  ⚠️ .xls 无法保留原始格式（仅保留数据，表头格式也会丢失）")
        keep_formulas_ok = keep_formulas and preserve_ok
        if keep_formulas and not preserve_ok and log_fn:
            log_fn("  ⚠️ 保留公式需同时开启保留格式（且不是 .xls 文件），本次已忽略该选项")

        # ---------- 读表：每个文件只完整加载一次，值和格式都从这一份取 ----------
        try:
            with _Ticker(log_fn, "读取数据"):
                wb_src, _ = open_book(file_path)
                if keep_formulas_ok:
                    # 额外打开一份 data_only=False 的源工作簿，专门取公式文本（不影响 wb_src 的缓存值读取）
                    wb_formula = openpyxl.load_workbook(file_path, read_only=False, data_only=False)
        except Exception as e:
            if stats is not None:
                stats["failed_files"] += 1
            if log_fn:
                log_fn(f"  ❌ 读取失败：{e}")
            return 0
        tick(0.30)

        # ---------- 逐 sheet 识别表头 ----------
        limit = _head_limit(config)
        hid = hidden_info(file_path)     # {sheet 名: (状态, {隐藏列号})}，xlsx 和 xls 同一套规则
        sheets = []        # [(sheet 名, ws, 表头行号, 最后一行有内容的行号, 表头行的列名)]
        cols = {}          # sheet 名 -> (每一列的列键, 列名)：跨文件合并时按列名对齐（见 _OutputBook.place）
        tpl_names = {}     # sheet 名 -> 完整列名（含隐藏列）：给 template_groups 比模板用
        hidden_cols = {}   # sheet 名 -> {隐藏列号}：产出里这些列照样隐藏；没开 keep_hidden 时内容清空
        hidden_skipped = 0
        for ws in wb_src.worksheets:
            state, hcols = hid.get(ws.title, ("visible", set()))
            if state != "visible":
                if stats is not None:
                    stats["hidden_sheets"] = stats.get("hidden_sheets", 0) + 1
                if not keep_hidden:
                    hidden_skipped += 1
                    if log_fn:
                        log_fn(f"  🙈 sheet「{ws.title}」是隐藏的，没有发出"
                               f"（要发出，打开「隐藏的 sheet 和列也照样发出」）")
                    continue
                if log_fn:
                    log_fn(f"  🙈 sheet「{ws.title}」是隐藏的，照样发出（产出里能看见）")
            last = _last_data_row(ws)
            head = [_key_row(r) for r in ws.iter_rows(min_row=1, max_row=min(last, limit),
                                                       values_only=True)] if last else []
            h = resolve_header_row(head, config)
            if h != -1 and h <= last and h <= len(head):
                sheets.append((ws.title, ws, h, last, _header_names(head[h - 1])))
                names_full = col_names(head, h)
                names_full += [""] * (ws.max_column - len(names_full))
                tpl_names[ws.title] = list(names_full)
                hcols = {c for c in hcols if c <= ws.max_column}
                hidden_cols[ws.title] = hcols
                keys_full = col_keys(names_full)
                if not keep_hidden:
                    # 列键照常按列名算（跨文件对齐不受影响），但隐藏列的列名不往产出里写
                    names_full = ["" if c in hcols else n for c, n in enumerate(names_full, start=1)]
                cols[ws.title] = (keys_full, names_full)
            else:
                if stats is not None:
                    stats["skipped_sheets"] += 1
                if log_fn:
                    log_fn(f"  ⚠️ sheet「{ws.title}」没识别出表头行，已跳过该 sheet")
        if not sheets:
            if hidden_skipped:      # 看得见的 sheet 都认不出表头，有内容的在隐藏 sheet 里：按设置不发，不算失败
                if log_fn:
                    log_fn("  ⚠️ 这个文件看得见的 sheet 里没有能拆的表，隐藏的 sheet 按设置没有发出")
                return 0
            if stats is not None:
                stats["failed_files"] += 1
            if log_fn:
                log_fn("  ⚠️ 未找到有效表头，跳过")
            return 0

        # ---------- 先按拆分列分好组，再清空隐藏列，最后读表头格式 ----------
        # 顺序不能反：隐藏的拆分列照样要能分组；表头格式里带着列名，要在清空之后读。
        plan = []          # [(sheet 名, ws, 表头行号, 最后一行, 拆分列号, 到人列号, 分组键, 到人键, 有内容的行)]
        validations = {}   # sheet 名 -> (表头区, 数据区) 的数据验证，见 _source_validations
        for sn, ws, h, last, names in sheets:
            col_idx = _find_col(names, split_column)
            if col_idx is None:
                if stats is not None:
                    stats["skipped_sheets"] += 1
                if log_fn:
                    log_fn(f"  ⚠️ sheet「{sn}」第 {h} 行（识别出的表头行）里没找到拆分字段"
                           f"「{split_column}」，已跳过该 sheet")
                continue
            pcol_idx = _find_col(names, person_column) if do_person else None
            # 数据行是 Excel 第 h+1 ~ last 行；第 j 个数据行的 0 基源行号是 h + j
            keys = ([apply_alias(key_value(v), alias) for v in _column_values(ws, col_idx + 1, h + 1, last)]
                    if last > h else [])
            pkeys = ([apply_alias(key_value(v), alias) for v in _column_values(ws, pcol_idx + 1, h + 1, last)]
                     if pcol_idx is not None and last > h else None)
            content = _content_rows(ws, h, last)        # 清空隐藏列之前算：只在隐藏列里有内容的行也算一行
            # 下拉选项可能就放在这张表的隐藏列里：也要在清空之前读
            try:
                dv_head, dv_data, dv_drop, dv_inline = _source_validations(wb_src, ws, h)
            except Exception:
                dv_head, dv_data, dv_drop, dv_inline = [], [], 0, 0
            validations[sn] = (dv_head, dv_data)
            if log_fn and (dv_drop or dv_inline):
                bits = ([f"{dv_inline} 个下拉的选项引用了别处的格子，已把选项直接写进下拉"] if dv_inline else []) + \
                       ([f"{dv_drop} 个数据验证引用了别处的格子（如自定义公式），拆出去会对不上，产出里去掉了"]
                        if dv_drop else [])
                log_fn(f"  ℹ️ sheet「{sn}」" + "；".join(bits))
            hcols = hidden_cols.get(sn, set())
            if hcols:
                if stats is not None:
                    stats["hidden_cols"] = stats.get("hidden_cols", 0) + len(hcols)
                shown = "、".join(f"「{tpl_names[sn][c - 1]}」" if tpl_names[sn][c - 1] else f"第 {c} 列"
                                 for c in sorted(hcols)[:5]) + ("等" if len(hcols) > 5 else "")
                if keep_hidden:
                    msg = f"有 {len(hcols)} 列是隐藏的（{shown}），照样发出、仍然隐藏"
                else:
                    _clear_hidden(ws, h, hcols)
                    msg = f"有 {len(hcols)} 列是隐藏的（{shown}），产出里这些列的内容已清空、仍然隐藏"
                if log_fn:
                    log_fn(f"  🙈 sheet「{sn}」{msg}")
            n_hidden_rows = _hidden_data_rows(ws, h, last)
            if n_hidden_rows:
                if stats is not None:
                    stats["hidden_rows"] = stats.get("hidden_rows", 0) + n_hidden_rows
                if log_fn:
                    log_fn(f"  ℹ️ sheet「{sn}」有 {n_hidden_rows} 行是隐藏的（可能是筛选掉的），"
                           f"照常发出，产出里不再隐藏")
            plan.append((sn, ws, h, last, col_idx, pcol_idx, keys, pkeys, content))

        try:
            header_meta = {p[0]: _read_header_format(p[1], p[2]) for p in plan}
        except Exception as e:
            if stats is not None:
                stats["failed_files"] += 1
            if log_fn:
                log_fn(f"  ❌ 读取格式失败：{e}")
            return 0
        est_rows = sum(max(0, last - h) for _sn, _ws, h, last, _names in sheets)
        if log_fn:
            log_fn(f"  ⏳ 读取完成（{len(sheets)} 个 sheet，约 {est_rows} 行数据），开始拆分…")
        tick(0.55)

        # 每个源文件一套样式缓存（源工作簿存活期间 id 稳定）
        style_cache = {}   # 按源文件作用域，见 _copy_cell_style
        total_rows = 0
        touched_keys = set()
        rows_written = 0
        last_hb_log = 0

        def heartbeat(n):
            """每批行让出一次 GIL 给界面线程，并周期性在日志里报心跳。"""
            nonlocal rows_written, last_hb_log
            rows_written += n
            time.sleep(0.001)
            if log_fn and rows_written - last_hb_log >= 2000:
                last_hb_log = rows_written
                log_fn(f"    …已拆分写入 {rows_written} 行")

        notes = {}         # 这个 sheet 跨文件对齐的情况 -> 涉及几个输出文件；拆完一个 sheet 合成几行日志

        def emit(key, save_path, sheet_name, h_idx, rows, ws_src, src_max_col, formula_ws):
            nonlocal total_rows
            if len(rows) == 0 or sheet_name not in header_meta:
                return
            ob = outputs.get(key)
            if ob is None:
                # 防撞名覆盖：不同取值净化后同名时，自动追加 (2)/(3)... 后缀（F1）
                used = {os.path.normcase(os.path.abspath(o.save_path)) for o in outputs.values()}
                if review is not None:      # 真有一个叫「请核对_无归属」的分组时，它拿 (2) 后缀
                    used.add(os.path.normcase(os.path.abspath(review.save_path)))
                ob = _OutputBook(_dedupe_path(save_path, used))
                outputs[key] = ob
            keys_, names_ = cols[sheet_name]
            sinfo, col_map, moved, detail = ob.place(sheet_name, h_idx, header_meta[sheet_name],
                                                     keys_, names_, ws_src)
            if moved or detail:
                d = detail or {}
                nk = (moved, tuple(d.get('missing', ())), tuple(d.get('extra', ())), bool(d.get('reordered')))
                notes[nk] = notes.get(nk, 0) + 1
            if col_map is not None and formula_ws is not None:
                notes['formula'] = notes.get('formula', 0) + 1     # _append_rows 会把这批行的公式落成数值
            for c in hidden_cols.get(sheet_name, ()):              # 源表隐藏的列，产出里照样隐藏
                d = c if col_map is None else col_map[c - 1]
                sinfo['ws'].column_dimensions[get_column_letter(d)].hidden = True
            dv_head, dv_data = validations.get(sheet_name, ((), ()))
            if 'dv' not in sinfo:      # 第一次往这个产出 sheet 写：表头区的数据验证照原位置带上
                sinfo['dv'] = {}
                for sig, tpl, ranges in dv_head:
                    cols_ = sinfo['dv'].setdefault(sig, (tpl, {}))[1]
                    for r1, r2, c1, c2 in ranges:
                        for c in range(c1, min(c2, src_max_col) + 1):
                            cols_.setdefault(c, set()).update(range(r1, r2 + 1))
            _note_validations(sinfo, dv_data, rows.index, sinfo['next_row'],
                              (lambda c: c) if col_map is None else (lambda c: col_map[c - 1]), src_max_col)
            _append_rows(sinfo, rows, preserve_ok, ws_src, src_max_col, style_cache,
                         heartbeat=heartbeat, formula_ws=formula_ws, col_map=col_map)
            total_rows += len(rows)
            touched_keys.add(key)

        # ---------- 逐 sheet 拆分 ----------
        contains = bool(selected_set) and not exact       # 包含匹配：一行可以落进几个分组
        n_sheets = max(1, len(plan))
        tpl = None         # 这个文件的模板：第一个含拆分字段的 sheet 的列名（见 template_groups）
        kinds = {}         # 分组键 -> 这一行算哪类（对账用；同一个键只判一次）

        def kind(k):
            c = kinds.get(k)
            if c is None:
                if hard_clean(k) == "":
                    c = 'unassigned'
                elif is_skip_value(k, skip):
                    c = 'skip_rows'
                elif not selected_set:
                    c = 'routed'
                elif contains:
                    c = 'routed' if any(p in k for p in selected if not is_skip_value(p, skip)) else 'not_selected'
                else:
                    c = 'routed' if k in selected_set else 'not_selected'
                kinds[k] = c
            return c

        for si, (sn, ws, h, last, col_idx, pcol_idx, keys, pkeys, content) in enumerate(plan):
            if stop_flag and stop_flag():
                break
            if tpl is None:
                tpl = tpl_names[sn]
            if last <= h:
                continue
            # 行数对账：只数有内容的行
            tally = {'source_rows': 0, 'routed': 0, 'skip_rows': 0, 'unassigned': 0, 'not_selected': 0}
            blank = []         # 拆分字段没写、但有内容的行（0 基数据行序号）
            end = max(content) if content else 0
            for j, k in enumerate(keys):
                if h + 1 + j in content:
                    c = kind(k)
                    if c == 'unassigned' and _looks_like_total(ws, h + 1 + j, end):
                        c = 'skip_rows'
                    tally['source_rows'] += 1
                    tally[c] += 1
                    if c == 'unassigned':
                        blank.append(j)
            if stats is not None:
                for tk, tv in tally.items():
                    stats[tk] = stats.get(tk, 0) + tv
            targets = selected if selected_set else sorted({k for k in keys if not is_skip_value(k, skip)})
            buckets = {}
            if not contains:
                for j, k in enumerate(keys):
                    buckets.setdefault(k, []).append(j)
            src_max_col = ws.max_column
            formula_ws = (wb_formula[sn] if (keep_formulas_ok and wb_formula is not None
                          and sn in wb_formula.sheetnames) else None)

            for ti, pval in enumerate(targets):
                if stop_flag and stop_flag():
                    break
                # 进度按「分组」推进：拆分阶段占本文件的 0.55~0.99
                tick(0.55 + 0.44 * ((si + ti / max(1, len(targets))) / n_sheets))
                if is_skip_value(pval, skip):
                    continue
                idxs = [j for j, k in enumerate(keys) if pval in k] if contains else buckets.get(pval)
                if not idxs:
                    continue

                # 汇总（始终产出）
                k, p = _summary_key_path(output_root, pval, src_stem, single_file, merge)
                emit(k, p, sn, h, _Rows([h + j for j in idxs]), ws, src_max_col, formula_ws)

                # 到人（可选附加产出）
                if pkeys is not None:
                    per = {}
                    for j in idxs:
                        per.setdefault(pkeys[j], []).append(j)
                    for person in sorted(v for v in per if not is_skip_value(v, skip)):
                        k2, p2 = _person_key_path(output_root, pval, person, src_stem, merge)
                        emit(k2, p2, sn, h, _Rows([h + j for j in per[person]]), ws, src_max_col, formula_ws)

            if blank and not (stop_flag and stop_flag()):
                if review is not None:
                    rinfo = _review_sheet(review, rel_path, sn, f"{src_stem}·{sn}", h, header_meta[sn])
                    for c in hidden_cols.get(sn, ()):
                        rinfo['ws'].column_dimensions[get_column_letter(c)].hidden = True
                    _append_rows(rinfo, _Rows([h + j for j in blank]), preserve_ok, ws, src_max_col,
                                 style_cache, heartbeat=heartbeat)
                if log_fn:
                    log_fn(f"  ⚠️ sheet「{sn}」有 {len(blank)} 行没写「{split_column}」，没分进任何一组"
                           + (f"，放进了「{UNASSIGNED}」请核对" if review is not None else ""))

            if notes:
                if stats is not None and any(nk != 'formula' and nk[0] for nk in notes):
                    stats["unaligned"] = stats.get("unaligned", 0) + 1
                if log_fn:
                    for line in _align_lines(sn, notes):
                        log_fn(line)
                notes.clear()

        if stats is not None and tpl is not None:
            stats.setdefault("templates", []).append((rel_path, tpl))
        if log_fn and total_rows:
            log_fn(f"  ✅ 本文件命中 {total_rows} 行 → 分入 {len(touched_keys)} 个输出文件")
        return total_rows
    finally:
        # 无论正常结束还是中途异常，都释放源工作簿
        if wb_src is not None:
            try:
                wb_src.close()
            except Exception:
                pass
        if wb_formula is not None:
            try:
                wb_formula.close()
            except Exception:
                pass


# =====================================================
# 主流程
# =====================================================

def _copy_file(src, dst):
    """按字节复制一个文件。"""
    with open(src, 'rb') as fi, open(dst, 'wb') as fo:
        for chunk in iter(lambda: fi.read(1 << 20), b''):
            fo.write(chunk)


def run_split(config, log_fn=None, progress_fn=None, stop_flag=None):
    """拆分入口：input_path 可为单文件或目录。一次运行产出『汇总』+ 可选『到人』，
    最后统一保存，并对每个生成了文件夹的主取值打包 ZIP。返回输出目录。

    progress_fn(0~1) 反映整体真实进度：处理各文件占 0~0.75（含文件内部各阶段），
    保存输出占 0.75~0.98，打包收尾到 1.0。"""
    t0 = time.time()
    input_path  = config["input_path"]
    output_root = config["output_path"]
    make_zip    = config.get('make_zip', True)
    single_file = os.path.isfile(input_path)

    timestamp   = datetime.datetime.now().strftime('%m%d%H%M')
    output_path = os.path.join(output_root, f"{timestamp}结果")
    # 时间戳只精确到分钟：一分钟内跑两次（选错字段马上重跑很常见）会把两次的产出
    # 写进同一个目录、混在一起且无从分辨。已存在就加序号，绝不复用旧目录。
    n = 2
    while os.path.exists(output_path):
        output_path = os.path.join(output_root, f"{timestamp}结果_{n}")
        n += 1
    os.makedirs(output_path, exist_ok=True)

    log_file = os.path.join(output_path, "运行日志.txt")

    def _log(msg):
        try:
            with open(log_file, 'a', encoding='utf-8') as lf:
                lf.write(msg + "\n")
        except Exception:
            pass
        if log_fn:
            log_fn(msg)

    # 收集任务（单文件 / 目录递归）
    tasks = []
    if single_file:
        tasks.append((input_path, os.path.basename(input_path)))
    else:
        for root, dirs, files in os.walk(input_path):
            if os.path.abspath(root).startswith(os.path.abspath(output_root)):
                continue
            # 按名字排序：Linux 上 os.walk 的顺序不固定，而跨文件合并时第一个文件定表头 /
            # 列顺序——不排序的话麒麟上同一批表拆出来的列顺序会和 Windows 不一样
            dirs.sort()
            for f in sorted(files):
                # 大小写统一用 lf 判断：区分大小写的文件系统（Linux）上，残留的
                # `X__TMP__.XLSX` 会绕过原来那条区分大小写的判断被当数据重新吃进去
                lf = f.lower()
                if (lf.endswith(('.xlsx', '.xls')) and not f.startswith('~$')
                        and not lf.endswith('__tmp__.xlsx')):   # 跳过旧版可能残留的临时文件
                    fp = os.path.join(root, f)
                    tasks.append((fp, os.path.relpath(fp, input_path)))

    do_person = config.get('to_person', False) and not single_file and bool(
        soft_clean(config.get('person_column', '')))

    total = len(tasks)
    _log(f"共找到 {total} 个文件，开始处理...")
    if total == 0:
        _log("⚠️ 输入里没有找到任何 Excel 文件（.xlsx / .xls），请检查输入路径")
    _log(f"拆分字段：{config.get('split_column', '')}"
         + (f"｜到人：{config.get('person_column', '')}" if do_person else "")
         + f"｜跨文件合并：{'是' if config.get('merge_across_files', True) else '否'}\n")

    outputs = {}
    review = _OutputBook(os.path.join(output_path, UNASSIGNED))   # 没写分组的行；有才保存
    grand_total = 0
    stopped = False
    stats = {"skipped_sheets": 0, "failed_files": 0, "unaligned": 0,     # 供收尾 [SUMMARY] 行汇总
             "hidden_sheets": 0, "hidden_cols": 0, "hidden_rows": 0}
    for i, (fp, rp) in enumerate(tasks):
        if stop_flag and stop_flag():
            _log("⛔ 已停止")
            stopped = True
            break
        try:
            size_mb = os.path.getsize(fp) / 1048576
            _log(f"[{i + 1}/{total}] {rp}（{size_mb:.1f} MB）")
        except OSError:
            _log(f"[{i + 1}/{total}] {rp}")

        # 整体进度 = (已完成文件数 + 当前文件内部进度) / 总文件数，处理阶段占 0~0.75。
        # 旧版按 (i+1)/total 报进度，单文件场景一开始就 100%，误导用户以为卡死。
        def _tick(frac, _i=i):
            if progress_fn:
                progress_fn((_i + frac) / total * 0.75)

        _tick(0.0)
        try:
            grand_total += process_file(fp, rp, output_path, config, outputs,
                                        single_file, _log, stop_flag, tick_fn=_tick,
                                        stats=stats, review=review)
        except Exception as e:
            stats["failed_files"] += 1
            _log(f"  ❌ 处理失败：{e}")
        _tick(1.0)

    # ---------- 跨文件合并时，提醒这批表其实有几种模板 ----------
    merge = config.get('merge_across_files', True)
    groups = template_groups(stats.get("templates", [])) if merge and not single_file else []
    if len(groups) > 1:
        _log(f"\n📋 这 {sum(len(g['files']) for g in groups)} 个表其实是 {len(groups)} 种模板"
             f"（含拆分字段的 sheet 列名不完全一样），跨文件合并已按列名对齐：")
        for line in template_lines(groups):
            _log(line)
        if stats["unaligned"]:
            _log(f"  其中 {stats['unaligned']} 个 sheet 和同名 sheet 的列名大半对不上，没有硬合，"
                 f"各自单独放进了「sheet 名_2」，见上面的 ⚠️")

    # ---------- 统一保存所有输出文件 ----------
    n_out = len(outputs)
    _log(f"\n正在保存 {n_out} 个输出文件...")
    saved = 0
    failed_saves = 0
    primaries = set()
    for j, (key, ob) in enumerate(outputs.items()):
        try:
            if ob.save():
                saved += 1
                primaries.add(key[1])   # key = (tree, primary, ...)
        except Exception as e:
            failed_saves += 1
            _log(f"  ❌ 保存失败 {os.path.basename(ob.save_path)}：{e}"
                 f"（若该文件正被 Excel/WPS 打开，请先关闭再重跑）")
        if progress_fn:
            progress_fn(0.75 + 0.23 * (j + 1) / n_out)
        if (j + 1) % 10 == 0 and (j + 1) < n_out:
            _log(f"  …已保存 {j + 1}/{n_out}")
        time.sleep(0.002)   # 让出 GIL，保存阶段保持界面响应

    # ---------- 没写分组的行：单独一个文件请用户核对（不算组数、不进 ZIP、不进下发清单） ----------
    unassigned_file = ""
    try:
        if review.save():
            unassigned_file = review.save_path
            _log(f"  📋 有 {stats.get('unassigned', 0)} 行没写「{config.get('split_column', '')}」，"
                 f"放进了「{UNASSIGNED}」（一个源 sheet 一页），请核对后补上分组再拆，或手动发给对应的人")
    except Exception as e:
        failed_saves += 1
        _log(f"  ❌ 保存失败 {UNASSIGNED}：{e}（若该文件正被 Excel/WPS 打开，请先关闭再重跑）")

    # ---------- 让对方填写的列（v2.10，选填）：标黄、锁定其余列、写下发记录 ----------
    # 必须在打包之前做，ZIP 里才是可以直接发出去的「待填写」版本。
    issued = False
    if config.get('fill_columns') and saved and not stopped:
        try:
            from core.collect import prepare_issue
            done = [((key[2] if key[0] == '到人' else key[1]), ob.save_path, key[0] == '汇总')
                    for key, ob in outputs.items() if os.path.exists(ob.save_path)]
            issued = bool(prepare_issue(output_path, done, config, _log))
        except Exception as e:
            _log(f"  ⚠️ 标出填写列失败（拆分结果不受影响）：{e}")

    # ---------- 对每个生成了文件夹的主取值打包 ZIP（方便分发） ----------
    zips = 0
    if make_zip and not single_file:
        for primary in primaries:
            folder = os.path.join(output_path, safe_filename(primary))
            if not os.path.isdir(folder):
                continue
            # 若该取值的汇总是扁平文件（跨文件合并模式），先复制进 {取值}/汇总/ 再打包，
            # 保证每个分组的 ZIP 自带「汇总 + 到人」，方便整包发负责人。
            sb = outputs.get(('汇总', primary))
            if sb is not None and os.path.exists(sb.save_path):
                try:
                    dst_dir = os.path.join(folder, "汇总")
                    os.makedirs(dst_dir, exist_ok=True)
                    _copy_file(sb.save_path, os.path.join(dst_dir, os.path.basename(sb.save_path)))
                except Exception:
                    pass
            if any(os.scandir(folder)):
                try:
                    zip_folder(folder, folder + ".zip")
                    zips += 1
                    _log(f"  📦 {safe_filename(primary)}.zip")
                except Exception as e:
                    _log(f"  ⚠️ 打包失败 {primary}：{e}")

    if progress_fn:
        progress_fn(1.0)
    mm, ss = divmod(int(time.time() - t0 + 0.5), 60)
    _log(f"\n✅ 全部完成！生成 {saved} 个文件，共写入 {grand_total} 行，"
         f"耗时 {f'{mm} 分 {ss} 秒' if mm else f'{ss} 秒'}")
    _log(f"📁 输出目录：{output_path}")
    src_rows = stats.get("source_rows", 0)
    if src_rows and not stopped:
        parts = [f"拆出 {stats.get('routed', 0)}"]
        if stats.get("skip_rows"):
            parts.append(f"合计等 {stats['skip_rows']}")
        if stats.get("not_selected"):
            parts.append(f"没选的分组 {stats['not_selected']}")
        if stats.get("unassigned"):
            parts.append(f"没写分组 {stats['unassigned']}（见 {UNASSIGNED}）")
        _log(f"🧮 行数核对：源表 {src_rows} 行 = " + " + ".join(parts))
    if len(primaries) == 1 and not config.get('selected_values') and not stopped:
        _log(f"⚠️ 只拆出了 1 组「{next(iter(primaries))}」：每一行的「{config.get('split_column', '')}」都一样，"
             f"可能选错了拆分字段")

    # 机器可读收尾行（v2.7）：GUI 提取后渲染完成摘要。追加式输出，不改变返回值与
    # 既有日志内容，下游 excelrouter-skill 只是多打一行，无契约变化。
    summary = {
        "mode": "excel",
        "groups": len(primaries),
        "files": saved,
        "rows": grand_total,
        "zips": zips,
        "skipped_sheets": stats["skipped_sheets"],
        "failed_files": stats["failed_files"],
        "failed_saves": failed_saves,
        "stopped": stopped,
        "output": output_path,
        "unaligned": stats["unaligned"],      # v3.0：列名对不上、单独放进「{sheet}_2」的来源 sheet 数
        # v3.0：隐藏的 sheet / 列 / 行各有几个（keep_hidden 关着时：sheet 没发、列清空了、行照常发）
        "hidden_sheets": stats["hidden_sheets"],
        "hidden_cols": stats["hidden_cols"],
        "hidden_rows": stats["hidden_rows"],
        # v3.0：行数对账（隐藏的、跳过的 sheet 不计）。source_rows = routed + skip_rows + unassigned + not_selected
        "source_rows": stats.get("source_rows", 0),
        "routed": stats.get("routed", 0),
        "skip_rows": stats.get("skip_rows", 0),
        "unassigned": stats.get("unassigned", 0),
        "not_selected": stats.get("not_selected", 0),
        "unassigned_file": unassigned_file,
    }
    if groups:
        summary["templates"] = len(groups)    # v3.0：跨文件合并时，这批表有几种模板
    if issued:
        summary["issued"] = True      # 写了下发记录：GUI 据此提示「收回来后用回收汇总查」
    _log(f"[SUMMARY] {json.dumps(summary, ensure_ascii=False)}")

    return output_path
