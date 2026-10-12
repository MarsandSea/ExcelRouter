"""core/splitter.py 的集成测试：表头识别、列枚举、跨文件合并、二级拆分、格式保留。"""
import datetime
import os

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import pytest

from core.splitter import (
    detect_header_row_auto, find_header_row, resolve_header_row,
    list_columns, list_values, run_split,
    _dedupe_path, _safe_sheet_title,
)


# ---------- 样本构造 ----------

HEADERS = ["工号", "姓名", "部门", "城市", "金额"]


def _make_book(path, rows, title="月度报表"):
    """表头在第2行、第1行是大标题（用于测自动表头识别）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = title
    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    for i, r in enumerate(rows, start=3):
        for c, v in enumerate(r, 1):
            ws.cell(row=i, column=c, value=v)
    wb.save(path)


def _count_rows(path, start=3):
    wb = openpyxl.load_workbook(path)
    total = 0
    for ws in wb.worksheets:
        for r in range(start, ws.max_row + 1):
            if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, ws.max_column + 1)):
                total += 1
    return total


@pytest.fixture
def two_books(tmp_path):
    inp = tmp_path / "in"
    out = tmp_path / "out"
    inp.mkdir()
    out.mkdir()
    _make_book(inp / "A.xlsx", [
        ["001", "张三", "销售部", "北京", 100],
        ["002", "李四", "销售部", "北京", 200],
        ["003", "王五", "销售部", "上海", 150],
        ["004", "赵六", "技术部", "北京", 300],
        ["", "合计", "", "", 750],
    ])
    _make_book(inp / "B.xlsx", [
        ["005", "钱七", "销售部", "广州", 120],
        ["006", "孙八", "技术部", "深圳", 220],
    ])
    return str(inp), str(out)


def _base_cfg(inp, out, **over):
    cfg = {
        "input_path": inp, "output_path": out,
        "header_mode": "auto", "header_row": 1,
        "grid_keys": [], "id_keys": [],
        "split_column": "部门", "selected_values": [],
        "person_column": "", "to_person": False, "person_file_filter": [],
        "value_alias_map": {},
        "skip_values": ["合计", "小计", "总计", ""],
        "merge_across_files": True, "make_zip": True, "exact_match": True,
        "preserve_format": True, "auto_open_output": False,
    }
    cfg.update(over)
    return cfg


# ---------- 表头识别 ----------

def test_detect_header_row_auto():
    rows = [["月度报表", None, None], ["工号", "姓名", "部门"], ["001", "张三", "销售部"]]
    assert detect_header_row_auto(rows) == 2


_D1, _D2 = datetime.datetime(2026, 10, 1), datetime.datetime(2026, 10, 2)
_DATED_HEAD = ["日期", "姓名", "部门", "备注", "金额"]


@pytest.mark.parametrize("rows, want", [
    # 合计行只填了金额：最后一条数据行下面「全是数字」，曾被当成表头
    ([_DATED_HEAD, [_D1, "张三", "销售部", "走访", 100], [_D2, "李四", "技术部", "回访", 200],
      [None, None, None, None, 300]], 1),
    # 表头上面有大标题
    ([["2026年10月走访明细"], _DATED_HEAD, [_D1, "张三", "销售部", "走访", 100],
      [None, None, None, None, 100]], 2),
    # 日期存成文字
    ([_DATED_HEAD, ["2026-10-01", "张三", "销售部", "走访", 100], ["2026-10-02", "李四", "技术部", "回访", 200],
      [None, None, None, None, 300]], 1),
    # 合计行带「合计」二字（原来就对，守住）
    ([_DATED_HEAD, [_D1, "张三", "销售部", "走访", 100], [_D2, "李四", "技术部", "回访", 200],
      ["合计", None, None, None, 300]], 1),
])
def test_detect_header_with_dates_and_sparse_total(rows, want):
    assert detect_header_row_auto(rows) == want


def test_detect_header_still_skips_blank_row_after_header():
    """表头下面隔一行空行再接数据，照样认得出（原有行为）。"""
    rows = [["工号", "姓名", "部门", "金额"], [None, None, None, None],
            ["001", "张三", "销售部", 100], ["002", "李四", "技术部", 200]]
    assert detect_header_row_auto(rows) == 1


def test_split_dated_sheet_with_sparse_total(tmp_path):
    """日期列 + 底部只填金额的合计行：整张表曾因认错表头被跳过，一行都没拆出来。"""
    inp, out = tmp_path / "走访.xlsx", tmp_path / "out"
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in [_DATED_HEAD, [_D1, "张三", "销售部", "走访", 100], [_D2, "李四", "技术部", "回访", 200],
              [None, None, None, None, 300]]:
        ws.append(r)
    wb.save(inp)
    logs = []
    res = run_split(_base_cfg(str(inp), str(out), make_zip=False), log_fn=logs.append)
    assert sorted(f for f in os.listdir(res) if f.endswith(".xlsx")) == ["技术部.xlsx", "销售部.xlsx"], logs
    ws_out = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx")).active
    assert [c.value for c in ws_out[1]] == _DATED_HEAD
    assert ws_out.cell(row=2, column=2).value == "张三"


def test_resolve_header_modes():
    rows = [["标题"], ["网格", "工号", "姓名"], ["A", "001", "张三"]]
    assert resolve_header_row(rows, {"header_mode": "auto"}) == 2
    assert resolve_header_row(rows, {"header_mode": "row", "header_row": 2}) == 2
    assert resolve_header_row(
        rows, {"header_mode": "keyword", "grid_keys": ["网格"], "id_keys": ["工号"]}) == 2
    assert find_header_row(rows, ["不存在"], ["工号"]) == -1


# ---------- 列名 / 取值枚举 ----------

def test_list_columns_and_values(two_books):
    inp, _ = two_books
    cfg = {"header_mode": "auto"}
    a = os.path.join(inp, "A.xlsx")
    assert list_columns(a, cfg) == HEADERS
    assert list_values(a, cfg, "部门") == ["技术部", "销售部"]   # 合计行被跳过


# ---------- 跨文件合并（核心）----------

def test_merge_across_files(two_books):
    inp, out = two_books
    res = run_split(_base_cfg(inp, out), log_fn=lambda m: None)
    sales = os.path.join(res, "销售部.xlsx")
    tech = os.path.join(res, "技术部.xlsx")
    assert os.path.exists(sales) and os.path.exists(tech)
    # A 的 3 行销售部 + B 的 1 行销售部 = 4（证明合并而非覆盖）
    assert _count_rows(sales) == 4
    assert _count_rows(tech) == 2


def test_no_merge_keeps_separate(two_books):
    inp, out = two_books
    res = run_split(_base_cfg(inp, out, merge_across_files=False), log_fn=lambda m: None)
    # 不合并：销售部/汇总 下按源文件分文件
    folder = os.path.join(res, "销售部", "汇总")
    assert os.path.isdir(folder)
    files = [f for f in os.listdir(folder) if f.endswith(".xlsx")]
    assert len(files) == 2   # A、B 各一个
    assert os.path.exists(os.path.join(res, "销售部.zip"))   # 每主取值打包


def test_single_file_input(two_books):
    inp, _ = two_books
    # 输入单个文件 -> 扁平输出，无文件夹、无 ZIP
    one = os.path.join(inp, "A.xlsx")
    out = os.path.join(os.path.dirname(inp), "single_out")
    os.makedirs(out, exist_ok=True)
    res = run_split(_base_cfg(one, out), log_fn=lambda m: None)
    assert _count_rows(os.path.join(res, "销售部.xlsx")) == 3   # A 表销售部 3 行
    assert _count_rows(os.path.join(res, "技术部.xlsx")) == 1
    assert not os.path.isdir(os.path.join(res, "销售部"))
    assert not os.path.exists(os.path.join(res, "销售部.zip"))


def test_rerun_same_minute_gets_fresh_dir(two_books):
    """时间戳只到分钟：连跑两次（选错字段马上重跑）不能写进同一个目录、混在一起。"""
    inp, out = two_books
    one = os.path.join(inp, "A.xlsx")
    first = run_split(_base_cfg(one, out), log_fn=lambda m: None)
    before = sorted(os.listdir(first))
    second = run_split(_base_cfg(one, out, split_column="城市"), log_fn=lambda m: None)
    assert second != first
    assert sorted(os.listdir(first)) == before          # 第一次的产出原样未被掺入
    assert os.path.exists(os.path.join(second, "北京.xlsx"))
    assert not os.path.exists(os.path.join(second, "销售部.xlsx"))


# ---------- 格式保留 ----------

def test_header_format_preserved(two_books):
    inp, out = two_books
    res = run_split(_base_cfg(inp, out), log_fn=lambda m: None)
    wb = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx"))
    ws = wb.active
    assert ws["A2"].value == "工号"
    assert ws["A2"].font.bold is True


# ---------- 保留公式（keep_formulas，仅同行公式安全平移）----------

def _make_formula_book(path):
    """单文件、表头第1行；D 列是同行公式（可安全平移），E 列是跨行区域公式（不安全）。

    销售部记录放在第3行，拆分后会上移到输出第2行，制造行号平移场景。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for c, h in enumerate(["部门", "单价", "数量", "金额", "合计"], 1):
        ws.cell(row=1, column=c, value=h)
    ws.append(["技术部", 10, 2, "=B2*C2", "=SUM(D2:D3)"])   # row2
    ws.append(["销售部", 5, 4, "=B3*C3", "=SUM(D2:D3)"])    # row3
    wb.save(path)


def test_keep_formulas_single_file(tmp_path):
    inp = tmp_path / "结算表.xlsx"
    out = tmp_path / "out"
    out.mkdir()
    _make_formula_book(inp)

    cfg = _base_cfg(str(inp), str(out), split_column="部门",
                    header_mode="row", header_row=1, keep_formulas=True)
    res = run_split(cfg, log_fn=lambda m: None)
    wb = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx"))
    ws = wb.active
    # 同行公式：安全，平移行号后仍是活公式
    assert ws["D2"].value == "=B2*C2"
    # 跨行区域公式：不安全，落成缓存数值（不是公式字符串）
    assert not (isinstance(ws["E2"].value, str) and ws["E2"].value.startswith("="))


def test_keep_formulas_default_off(tmp_path):
    inp = tmp_path / "结算表.xlsx"
    out = tmp_path / "out"
    out.mkdir()
    _make_formula_book(inp)

    cfg = _base_cfg(str(inp), str(out), split_column="部门",
                    header_mode="row", header_row=1)   # keep_formulas 不传，走默认 False
    res = run_split(cfg, log_fn=lambda m: None)
    wb = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx"))
    ws = wb.active
    assert not (isinstance(ws["D2"].value, str) and ws["D2"].value.startswith("="))


# ---------- 到人（双产出 + 范围过滤）----------

def test_to_person_all(two_books):
    inp, out = two_books
    res = run_split(_base_cfg(inp, out, to_person=True, person_column="姓名",
                              merge_across_files=False), log_fn=lambda m: None)
    # 汇总照常产出
    assert os.path.isdir(os.path.join(res, "销售部", "汇总"))
    # 到人：销售部里 A 的张三/李四/王五 + B 的钱七 都各一个；文件名带来源文件名后缀
    for name, stem in (("张三", "A"), ("李四", "A"), ("王五", "A"), ("钱七", "B")):
        assert os.path.exists(os.path.join(res, "销售部", "到人", f"{name}_{stem}.xlsx")), name
    # 技术部到人：A 赵六 + B 孙八
    assert os.path.exists(os.path.join(res, "技术部", "到人", "赵六_A.xlsx"))
    assert os.path.exists(os.path.join(res, "技术部", "到人", "孙八_B.xlsx"))


def test_to_person_scope_filter(two_books):
    inp, out = two_books
    # 仅文件名含 "A" 的表参与到人
    res = run_split(_base_cfg(inp, out, to_person=True, person_column="姓名",
                              person_file_filter=["A"], merge_across_files=False),
                    log_fn=lambda m: None)
    assert os.path.exists(os.path.join(res, "销售部", "到人", "张三_A.xlsx"))   # A 表
    assert not os.path.exists(os.path.join(res, "销售部", "到人", "钱七_B.xlsx"))  # B 表，不在范围
    # 但 B 表仍进了汇总
    assert os.path.exists(os.path.join(res, "销售部", "汇总", "销售部_B.xlsx"))


# ---------- 只拆部分取值 ----------

def test_selected_values_subset(two_books):
    inp, out = two_books
    res = run_split(_base_cfg(inp, out, selected_values=["销售部"]), log_fn=lambda m: None)
    assert os.path.exists(os.path.join(res, "销售部.xlsx"))
    assert not os.path.exists(os.path.join(res, "技术部.xlsx"))


# ---------- 取值归并映射 ----------

def test_value_alias_map(tmp_path):
    inp = tmp_path / "in"; out = tmp_path / "out"
    inp.mkdir(); out.mkdir()
    _make_book(inp / "C.xlsx", [
        ["001", "张三", "销售部", "北京", 100],
        ["002", "李四", "销售一部", "上海", 200],
    ])
    cfg = _base_cfg(str(inp), str(out),
                    value_alias_map={"销售": ["销售部", "销售一部"]})
    res = run_split(cfg, log_fn=lambda m: None)
    merged = os.path.join(res, "销售.xlsx")
    assert os.path.exists(merged)
    assert _count_rows(merged) == 2                   # 两个别名归并成一组


# ---------- 对抗式审查回归用例 ----------

def _make_book_custom(path, headers, rows, title="报表", sheet="Sheet1"):
    """可指定列头/行/sheet 名的样本（表头在第2行）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws["A1"] = title
    for c, h in enumerate(headers, 1):
        ws.cell(row=2, column=c, value=h)
    for i, r in enumerate(rows, start=3):
        for c, v in enumerate(r, 1):
            ws.cell(row=i, column=c, value=v)
    wb.save(path)


def test_filename_collision_no_overwrite(tmp_path):
    """F1：两个净化后同名（销售_部）但不同的取值不能互相覆盖。"""
    import glob
    inp = tmp_path / "in"; out = tmp_path / "out"
    inp.mkdir(); out.mkdir()
    _make_book_custom(inp / "X.xlsx", HEADERS, [
        ["001", "张三", "销售/部", "北京", 100],
        ["002", "李四", "销售/部", "北京", 200],
        ["003", "王五", "销售:部", "上海", 150],   # 净化后同为「销售_部」
    ])
    res = run_split(_base_cfg(str(inp / "X.xlsx"), str(out)), log_fn=lambda m: None)
    files = glob.glob(os.path.join(res, "销售_部*.xlsx"))
    assert len(files) == 2, files                       # 两份并存（销售_部.xlsx + 销售_部(2).xlsx）
    assert sum(_count_rows(f) for f in files) == 3       # 3 行数据一行不丢


def test_skip_leftover_temp_file(two_books):
    """F2：旧版可能残留在输入目录的 *__tmp__.xlsx 不应被当成正式输入处理。"""
    inp, out = two_books
    _make_book_custom(os.path.join(inp, "ghost__tmp__.xlsx"), HEADERS, [
        ["009", "幽灵", "幽灵部", "火星", 999],
    ])
    res = run_split(_base_cfg(inp, out), log_fn=lambda m: None)
    assert not os.path.exists(os.path.join(res, "幽灵部.xlsx"))   # 残留临时文件被跳过
    assert os.path.exists(os.path.join(res, "销售部.xlsx"))       # 正常文件照常处理


def test_column_mismatch_aligns_by_name(tmp_path):
    """F7（v3.0 起按列名对齐）：B 多一列「备注」加在最后、表头照抄；C 少「城市」，金额仍进金额列，城市留空。"""
    inp = tmp_path / "in"; out = tmp_path / "out"
    inp.mkdir(); out.mkdir()
    _make_book_custom(inp / "A.xlsx", HEADERS, [["001", "张三", "销售部", "北京", 100]])
    _make_book_custom(inp / "B.xlsx", HEADERS + ["备注"],
                      [["002", "李四", "销售部", "上海", 200, "x"]])      # 多一列
    _make_book_custom(inp / "C.xlsx", ["工号", "姓名", "部门", "金额"],
                      [["003", "王五", "销售部", 300]])                   # 少一列
    logs = []
    res = run_split(_base_cfg(str(inp), str(out)), log_fn=logs.append)
    ws = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx")).active
    assert [c.value for c in ws[2]] == HEADERS + ["备注"]
    assert [c.value for c in ws[3]] == ["001", "张三", "销售部", "北京", 100, None]
    assert [c.value for c in ws[4]] == ["002", "李四", "销售部", "上海", 200, "x"]
    assert [c.value for c in ws[5]] == ["003", "王五", "销售部", None, 300, None]
    assert any("多「备注」（加在最后）" in m for m in logs), logs
    assert any("少「城市」（留空）" in m for m in logs), logs
    assert not any("列数" in m for m in logs)                             # 旧的「按列位置合并可能错位」不再出现


def test_zip_bundles_summary_with_person(two_books):
    """F5：merge=True + 到人 时，每个分组的 ZIP 内应同时含『汇总』与『到人』。"""
    import zipfile
    inp, out = two_books
    res = run_split(_base_cfg(inp, out, to_person=True, person_column="姓名"),
                    log_fn=lambda m: None)   # merge 默认 True → 汇总走扁平文件
    zpath = os.path.join(res, "销售部.zip")
    assert os.path.exists(zpath)
    names = [n.replace("\\", "/") for n in zipfile.ZipFile(zpath).namelist()]
    assert any(n.startswith("汇总/") and n.endswith("销售部.xlsx") for n in names), names   # 汇总进包
    assert any(n.startswith("到人/") and n.endswith("张三.xlsx") for n in names), names     # 到人也在包里（v2.7.2：merge 模式不带源文件名）


def test_dedupe_path_helper():
    """F1 工具：撞名时追加 (2)/(3)... 后缀。"""
    base = os.path.join("out", "a.xlsx")
    used = set()
    p1 = _dedupe_path(base, used)
    used.add(os.path.normcase(os.path.abspath(p1)))
    p2 = _dedupe_path(base, used)
    assert os.path.basename(p1) == "a.xlsx"
    assert os.path.basename(p2) == "a(2).xlsx"


def test_safe_sheet_title_helper():
    """F3 工具：非法字符替换、超长截断、重名去重。"""
    used = set()
    assert _safe_sheet_title("数据:报表*1", used) == "数据_报表_1"
    assert len(_safe_sheet_title("标" * 40, used)) <= 31
    used2 = set()
    assert _safe_sheet_title("X", used2) == "X"
    assert _safe_sheet_title("X", used2) == "X(2)"

# ---------- 收尾摘要行（v2.7，GUI 完成摘要的数据源）----------

def test_summary_line(two_books):
    """run_split 收尾应打恰好一行 [SUMMARY] JSON，含分组/文件/行数/跳过统计。"""
    import json
    inp, out = two_books
    logs = []
    result = run_split(_base_cfg(inp, out), log_fn=logs.append)
    summary = [m for m in logs if m.startswith("[SUMMARY] ")]
    assert len(summary) == 1
    s = json.loads(summary[0][len("[SUMMARY] "):])
    assert s["mode"] == "excel"
    assert s["groups"] == 2            # 销售部 + 技术部
    assert s["rows"] == 6              # 4 + 2 行数据（合计行按跳过值忽略）
    assert s["files"] >= 2
    assert s["skipped_sheets"] == 0
    assert s["failed_files"] == 0
    assert s["stopped"] is False
    assert s["output"] == result

def test_summary_line_counts_skipped_sheets(tmp_path):
    """拆分字段找不到而跳过的 sheet 要计入 skipped_sheets。"""
    import json
    inp = tmp_path / "in"
    out = tmp_path / "out"
    inp.mkdir()
    # 两个 sheet：一个有「部门」列，一个没有 → 后者被跳过
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "有部门"
    for c, h in enumerate(["工号", "姓名", "部门"], 1):
        ws1.cell(row=1, column=c, value=h)
    ws1.cell(row=2, column=1, value="001")
    ws1.cell(row=2, column=2, value="张三")
    ws1.cell(row=2, column=3, value="销售部")
    ws2 = wb.create_sheet("没有部门")
    for c, h in enumerate(["编号", "说明"], 1):
        ws2.cell(row=1, column=c, value=h)
    ws2.cell(row=2, column=1, value="x1")
    ws2.cell(row=2, column=2, value="备注")
    wb.save(inp / "multi.xlsx")
    logs = []
    run_split(_base_cfg(str(inp), str(out)), log_fn=logs.append)
    s = json.loads([m for m in logs if m.startswith("[SUMMARY] ")][0][len("[SUMMARY] "):])
    assert s["skipped_sheets"] == 1
    assert s["groups"] == 1


# ---------- 到人：同一人跨多个源文件（v2.7.2 回归，名实不符 bug）----------

def _make_month_books(inp):
    """三张月度数据源：张三在 7/8/9 月都有行，李四只在 7 月。"""
    for mon, rows in [("7月明细", [("张三", 100), ("李四", 200)]),
                      ("8月明细", [("张三", 300)]),
                      ("9月明细", [("张三", 400)])]:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["姓名", "部门", "金额"])
        for r in rows:
            ws.append([r[0], "销售部", r[1]])
        wb.save(os.path.join(inp, f"{mon}.xlsx"))


def test_to_person_cross_files_merge(tmp_path):
    """merge=True：同一人跨源的全部行进一张「姓名.xlsx」，不再叫「姓名_第一个源.xlsx」。"""
    inp, out = tmp_path / "in", tmp_path / "out"
    inp.mkdir()
    _make_month_books(str(inp))
    res = run_split(_base_cfg(str(inp), str(out), to_person=True,
                              person_column="姓名"), log_fn=lambda m: None)
    d = os.path.join(res, "销售部", "到人")
    # 名实相符：文件名不带源文件名，且只有一张
    assert sorted(os.listdir(d)) == ["张三.xlsx", "李四.xlsx"]
    wb = openpyxl.load_workbook(os.path.join(d, "张三.xlsx"))
    ws = wb.active
    amounts = sorted(ws.cell(row=i, column=3).value for i in range(2, ws.max_row + 1))
    wb.close()
    assert amounts == [100, 300, 400]        # 7/8/9 月的行都在，一行不少
    # 旧命名不应再出现
    assert not os.path.exists(os.path.join(d, "张三_7月明细.xlsx"))


def test_to_person_cross_files_no_merge(tmp_path):
    """merge=False：每人每源一张，与汇总的按原表拆分对称；行不串表。"""
    inp, out = tmp_path / "in", tmp_path / "out"
    inp.mkdir()
    _make_month_books(str(inp))
    res = run_split(_base_cfg(str(inp), str(out), to_person=True, person_column="姓名",
                              merge_across_files=False), log_fn=lambda m: None)
    d = os.path.join(res, "销售部", "到人")
    assert set(os.listdir(d)) == {
        "李四_7月明细.xlsx", "张三_7月明细.xlsx", "张三_8月明细.xlsx", "张三_9月明细.xlsx"}
    for fname, expect in [("张三_7月明细.xlsx", [100]), ("张三_8月明细.xlsx", [300]),
                          ("张三_9月明细.xlsx", [400]), ("李四_7月明细.xlsx", [200])]:
        wb = openpyxl.load_workbook(os.path.join(d, fname))
        ws = wb.active
        amounts = [ws.cell(row=i, column=3).value for i in range(2, ws.max_row + 1)]
        wb.close()
        assert amounts == expect, (fname, amounts)


def test_walk_skips_tmp_file_case_insensitively(tmp_path):
    """残留的临时文件不论大小写都要跳过。

    旧实现里扩展名判断用的是 f.lower()，而 __tmp__ 判断用的是原始 f —— 在区分
    大小写的文件系统（Linux / 麒麟）上，`X__TMP__.XLSX` 会绕过过滤被当成数据表
    重新吃进去，导致同一批数据被重复拆分。
    """
    inp = tmp_path / "in"
    inp.mkdir()
    out = tmp_path / "out"
    _make_book(inp / "正常表.xlsx", [
        ["001", "张三", "技术部", "北京", 100],
        ["002", "李四", "销售部", "上海", 200],
    ])
    # 大写命名的残留临时文件：内容里的部门取值刻意与正常表不同，便于断言
    _make_book(inp / "旧残留__TMP__.XLSX", [
        ["901", "王五", "不该出现的部门", "广州", 300],
    ])

    run_split(_base_cfg(str(inp), str(out)), log_fn=lambda *_: None)

    names = {p.name for p in out.rglob("*.xlsx")}
    assert not any("不该出现的部门" in n for n in names), names
    assert not (out / "不该出现的部门").exists()


def _many_styles_book(path, n=300):
    """每行一种颜色组合、隔行换部门：专测「格式照原样」不是碰巧对的。"""
    colors = ["FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["部门", "姓名", "金额"])
    for i in range(n):
        ws.append(["A" if i % 2 else "B", f"人{i}", i])
        for c in ws[ws.max_row]:
            c.fill = PatternFill("solid", fgColor=colors[i % 6])
            c.font = Font(bold=i % 3 == 0, color=colors[(i + 2) % 6])
    wb.save(path)


def _style_sig(c):
    return c.fill.fgColor.rgb, c.font.bold, c.font.color.rgb if c.font.color else None


def test_preserve_format_copies_each_cells_own_style(tmp_path):
    """回归：旧版按 id(cell.font) 缓存样式，cell.font 每次都是新的临时对象，地址被复用时会
    取到别的格子的样式——实测 900 个格子里约 20 个填充色 / 字体错了。"""
    src = tmp_path / "src.xlsx"
    _many_styles_book(src)
    cfg = {"input_path": str(src), "output_path": str(tmp_path / "out"), "split_column": "部门",
           "header_mode": "auto", "preserve_format": True, "make_zip": False, "merge_across_files": False,
           "skip_values": [""], "exact_match": True, "selected_values": []}
    out = run_split(cfg, log_fn=lambda m: None)
    ws_src = openpyxl.load_workbook(src).active
    for dept, parity in (("A", 1), ("B", 0)):
        path = next(os.path.join(d, f) for d, _, fs in os.walk(out) for f in fs if f.startswith(dept))
        ws = openpyxl.load_workbook(path).active
        src_rows = [r for r in range(2, 302) if (r - 2) % 2 == parity]
        for k, r in enumerate(src_rows):
            for c in (1, 2, 3):
                assert _style_sig(ws.cell(k + 2, c)) == _style_sig(ws_src.cell(r, c)), (dept, r, c)


# ---------- v3.0：纯 openpyxl 读表 ----------

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _cfg_for(inp, out, **over):
    return _base_cfg(str(inp), str(out), make_zip=False, merge_across_files=False, **over)


def test_split_does_not_need_pandas(two_books):
    """拆分全程不导入 pandas：子进程里把 pandas 设成导入就报错，列名、取值、拆分照常。"""
    import json
    import subprocess
    import sys
    inp, out = two_books
    code = ("import json, os, sys\n"
            "sys.modules['pandas'] = None\n"
            "from core.splitter import list_columns, list_values, run_split\n"
            "cfg = json.loads(sys.argv[1])\n"
            "a = os.path.join(cfg['input_path'], 'A.xlsx')\n"
            "res = run_split(cfg, log_fn=lambda m: None)\n"
            "print(json.dumps([list_columns(a, cfg), list_values(a, cfg, '部门'), sorted(os.listdir(res))]))\n")
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    p = subprocess.run([sys.executable, "-c", code, json.dumps(_base_cfg(inp, out))], cwd=ROOT, env=env,
                       capture_output=True, text=True, encoding="utf-8")
    assert p.returncode == 0, p.stderr
    cols, vals, files = json.loads(p.stdout)
    assert cols == HEADERS
    assert vals == ["技术部", "销售部"]
    assert "销售部.xlsx" in files and "技术部.xlsx" in files


def _injection_book(path):
    """表头上方的说明行、备注列里都有以「=」开头的文字；金额列是真公式（同行）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "结算"
    ws.append(["=说明：本表按部门拆"])
    ws["A1"].data_type = "s"
    ws.append(["部门", "备注", "单价", "数量", "金额"])
    ws.append(["销售部", "=单价×数量", 10, 3, "=C3*D3"])
    ws["B3"].data_type = "s"
    ws.append(["销售部", "正常", 5, 2, "=C4*D4"])
    wb.save(path)


@pytest.mark.parametrize("keep, preserve", [(True, True), (False, True), (False, False)])
def test_equals_text_never_becomes_formula(tmp_path, keep, preserve):
    """安全修复：源表里以「=」开头的文字，拆出来仍是文字；真公式开了「保留公式」才是活公式。"""
    src = tmp_path / "结算.xlsx"
    _injection_book(src)
    res = run_split(_cfg_for(src, tmp_path / "out", keep_formulas=keep, preserve_format=preserve),
                    log_fn=lambda m: None)
    ws = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx")).active
    assert (ws["A1"].value, ws["A1"].data_type) == ("=说明：本表按部门拆", "s")
    assert (ws["B3"].value, ws["B3"].data_type) == ("=单价×数量", "s")
    if keep:
        assert (ws["E3"].value, ws["E3"].data_type) == ("=C3*D3", "f")
        assert (ws["E4"].value, ws["E4"].data_type) == ("=C4*D4", "f")
    else:
        # 造表时公式没有计算结果，不保留公式时落成空值；关键是不能是公式
        assert ws["E3"].data_type != "f" and ws["E4"].data_type != "f"


def test_na_like_words_are_real_groups(tmp_path):
    """「NA」「N/A」「null」是用户写的文字：旧版（pandas）当成空格子，整行被丢掉；现在照常成组、照常输出。"""
    src = tmp_path / "na.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["部门", "编号", "备注"])
    for row in (["NA", 1, None], ["N/A", 2, None], ["null", 3, None], ["销售部", 4, "N/A"]):
        ws.append(row)
    wb.save(src)
    cfg = _cfg_for(src, tmp_path / "out", preserve_format=False)
    assert list_values(str(src), cfg, "部门") == sorted(["NA", "N/A", "null", "销售部"])
    res = run_split(cfg, log_fn=lambda m: None)
    assert {"NA.xlsx", "N_A.xlsx", "null.xlsx", "销售部.xlsx"} <= set(os.listdir(res))
    sales = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx")).active
    assert sales["C2"].value == "N/A"           # 快速模式下也不再变成空白


def test_error_cells_are_not_groups(tmp_path):
    """错误格（#N/A、#DIV/0!）和文字「#N/A」都不成组，取值预览里也看不到（预览和拆分口径一致）。"""
    src = tmp_path / "err.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["部门", "金额"])
    ws.append(["销售部", 1])
    ws.append(["#N/A", 2])                      # 错误格
    ws.append(["#N/A", 3])
    ws["A4"].data_type = "s"                    # 写成文字的「#N/A」
    ws.append(["技术部", "#DIV/0!"])
    wb.save(src)
    cfg = _cfg_for(src, tmp_path / "out")
    assert list_values(str(src), cfg, "部门") == ["技术部", "销售部"]
    res = run_split(cfg, log_fn=lambda m: None)
    groups = {f for f in os.listdir(res) if f.endswith(".xlsx")} - {"请核对_无归属.xlsx"}
    assert groups == {"技术部.xlsx", "销售部.xlsx"}


def test_numeric_header_names_read_like_text():
    """表头格存成 2026.0 时列名仍叫「2026」（和旧版 pandas 一致）；错误格的列名当空。"""
    from core.splitter import _find_col, _header_names
    names = _header_names([2026.0, "部门", None, "#N/A"])
    assert names == ["2026", "部门", "", ""]
    assert _find_col(names, "2026") == 0
    assert _find_col(names, "部") == 1


# ---------- v3.0：跨文件合并按列名对齐 ----------

REORDERED = ["姓名", "工号", "部门", "金额", "城市"]


def _summary(logs):
    import json
    return json.loads([m for m in logs if m.startswith("[SUMMARY] ")][0][len("[SUMMARY] "):])


def _values(ws, row):
    return [c.value for c in ws[row]]


@pytest.mark.parametrize("preserve", [True, False])
def test_merge_aligns_reordered_columns(tmp_path, preserve):
    """B 的列顺序和 A 不一样：旧版按位置追加，姓名落进工号列；现在每个值进它自己的列。"""
    inp = tmp_path / "in"
    inp.mkdir()
    _make_book_custom(inp / "A.xlsx", HEADERS, [["001", "张三", "销售部", "北京", 100]])
    _make_book_custom(inp / "B.xlsx", REORDERED, [["李四", "002", "销售部", 200, "上海"]])
    logs = []
    res = run_split(_base_cfg(str(inp), str(tmp_path / "out"), preserve_format=preserve), log_fn=logs.append)
    ws = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx")).active
    assert _values(ws, 2) == HEADERS
    assert _values(ws, 3) == ["001", "张三", "销售部", "北京", 100]
    assert _values(ws, 4) == ["002", "李四", "销售部", "上海", 200]
    assert any("已按列名对齐：列的顺序不一样" in m for m in logs), logs
    s = _summary(logs)
    assert s["unaligned"] == 0 and s["templates"] == 1         # 列一样、只是顺序不同，算同一种模板


def test_dir_files_processed_in_name_order(tmp_path, monkeypatch):
    """Linux 上 os.walk 不按名字排序：v3.0.0 的麒麟构建里 B 先被处理、B 的列顺序成了表头。
    这里把 os.walk 的结果倒过来模拟，第一个文件仍须是 A。"""
    real_walk = os.walk

    def shuffled_walk(top, *a, **kw):
        for root, dirs, files in real_walk(top, *a, **kw):
            dirs.reverse()
            yield root, dirs, files[::-1]

    monkeypatch.setattr(os, "walk", shuffled_walk)
    inp = tmp_path / "in"
    inp.mkdir()
    _make_book_custom(inp / "A.xlsx", HEADERS, [["001", "张三", "销售部", "北京", 100]])
    _make_book_custom(inp / "B.xlsx", REORDERED, [["李四", "002", "销售部", 200, "上海"]])
    res = run_split(_base_cfg(str(inp), str(tmp_path / "out")), log_fn=lambda m: None)
    ws = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx")).active
    assert _values(ws, 2) == HEADERS
    assert _values(ws, 3) == ["001", "张三", "销售部", "北京", 100]


def test_merge_unaligned_goes_to_own_sheet(tmp_path):
    """对上的列名不到一半：不硬合，这个来源的行放进「Sheet1_2」，用它自己的表头；一行不丢。"""
    inp = tmp_path / "in"
    inp.mkdir()
    _make_book_custom(inp / "A.xlsx", HEADERS, [["001", "张三", "销售部", "北京", 100]])
    _make_book_custom(inp / "B.xlsx", ["编号", "名称", "部门", "地址", "数值"],
                      [["X1", "甲", "销售部", "广州", 7], ["X2", "乙", "技术部", "深圳", 8]])
    logs = []
    res = run_split(_base_cfg(str(inp), str(tmp_path / "out")), log_fn=logs.append)
    wb = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx"))
    assert wb.sheetnames == ["Sheet1", "Sheet1_2"]
    assert _values(wb["Sheet1"], 3) == ["001", "张三", "销售部", "北京", 100]
    assert wb["Sheet1"].max_row == 3
    assert _values(wb["Sheet1_2"], 2) == ["编号", "名称", "部门", "地址", "数值"]
    assert _values(wb["Sheet1_2"], 3) == ["X1", "甲", "销售部", "广州", 7]
    tech = openpyxl.load_workbook(os.path.join(res, "技术部.xlsx"))
    assert tech.sheetnames == ["Sheet1"]                      # 技术部只有 B 的行：用 B 的表头建，不用挪
    assert any("单独放进「Sheet1_2」" in m for m in logs), logs
    assert any("2 种模板" in m for m in logs), logs
    s = _summary(logs)
    assert s["unaligned"] == 1 and s["templates"] == 2 and s["rows"] == 3


def test_merge_same_template_reports_nothing(two_books):
    """同一套模板：没有对齐日志，摘要里 1 种模板、0 个没对上。"""
    inp, out = two_books
    logs = []
    run_split(_base_cfg(inp, out), log_fn=logs.append)
    assert not any("按列名对齐" in m or "模板" in m for m in logs), logs
    s = _summary(logs)
    assert s["unaligned"] == 0 and s["templates"] == 1


def test_no_merge_has_no_template_note(tmp_path):
    """不跨文件合并时各表各拆，模板不一样也不用提醒。"""
    inp = tmp_path / "in"
    inp.mkdir()
    _make_book_custom(inp / "A.xlsx", HEADERS, [["001", "张三", "销售部", "北京", 100]])
    _make_book_custom(inp / "B.xlsx", ["编号", "部门"], [["X1", "销售部"]])
    logs = []
    run_split(_base_cfg(str(inp), str(tmp_path / "out"), merge_across_files=False), log_fn=logs.append)
    assert not any("模板" in m for m in logs)
    assert "templates" not in _summary(logs)


def test_to_person_merge_aligns_columns(tmp_path):
    """到人（merge=True）也按列名对齐：张三在两个列顺序不同的表里，合进一张表、各列对得上。"""
    inp = tmp_path / "in"
    inp.mkdir()
    _make_book_custom(inp / "A.xlsx", HEADERS, [["001", "张三", "销售部", "北京", 100]])
    _make_book_custom(inp / "B.xlsx", REORDERED, [["张三", "001", "销售部", 200, "上海"]])
    res = run_split(_base_cfg(str(inp), str(tmp_path / "out"), to_person=True, person_column="姓名"),
                    log_fn=lambda m: None)
    ws = openpyxl.load_workbook(os.path.join(res, "销售部", "到人", "张三.xlsx")).active
    assert _values(ws, 3) == ["001", "张三", "销售部", "北京", 100]
    assert _values(ws, 4) == ["001", "张三", "销售部", "上海", 200]


def _formula_book(path, headers, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(path)


def test_keep_formulas_dropped_only_when_columns_move(tmp_path):
    """保留公式 + 跨文件合并：列位置没变的表照样保留同行公式；列换了位置的表落成数值，并在日志里说明。"""
    inp = tmp_path / "in"
    inp.mkdir()
    _formula_book(inp / "A.xlsx", ["部门", "单价", "数量", "金额"], [["销售部", 2, 3, "=B2*C2"]])
    _formula_book(inp / "B.xlsx", ["部门", "单价", "数量", "金额", "备注"], [["销售部", 4, 5, "=B2*C2", "多一列"]])
    _formula_book(inp / "C.xlsx", ["部门", "数量", "单价", "金额"], [["销售部", 6, 7, "=C2*B2"]])
    logs = []
    res = run_split(_base_cfg(str(inp), str(tmp_path / "out"), header_mode="row", header_row=1,
                              keep_formulas=True), log_fn=logs.append)
    ws = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx")).active
    assert _values(ws, 1) == ["部门", "单价", "数量", "金额", "备注"]
    assert (ws["D2"].value, ws["D2"].data_type) == ("=B2*C2", "f")    # A：原位
    assert (ws["D3"].value, ws["D3"].data_type) == ("=B3*C3", "f")    # B：只是多一列，位置没变
    assert ws["E3"].value == "多一列"
    assert (ws["B4"].value, ws["C4"].value) == (7, 6)                # C：单价、数量按列名换回来
    assert ws["D4"].data_type != "f"                                 # 公式按列字母引用，换了位置不能留
    assert any("公式落成了数值" in m for m in logs), logs


def test_scan_values_and_template_groups(tmp_path):
    """scan_values 一遍读出取值和模板；template_groups 不计列顺序分组，template_lines 说出多了 / 少了哪些列。"""
    from core.splitter import scan_values, template_groups, template_lines
    a, b, c = tmp_path / "A.xlsx", tmp_path / "B.xlsx", tmp_path / "C.xlsx"
    _make_book_custom(a, HEADERS, [["001", "张三", "销售部", "北京", 100]])
    _make_book_custom(b, REORDERED, [["李四", "002", "技术部", 200, "上海"]])
    _make_book_custom(c, HEADERS[:4] + ["备注"], [["003", "王五", "销售部", "广州", "x"]])
    cfg = {"header_mode": "auto"}
    vals, tpl = scan_values(str(a), cfg, "部门")
    assert vals == ["销售部"] and tpl == HEADERS
    assert scan_values(str(a), cfg, "不存在的列") == ([], None)
    pairs = [(str(p), scan_values(str(p), cfg, "部门")[1]) for p in (a, b, c)] + [("坏文件.xlsx", None)]
    groups = template_groups(pairs)
    assert [len(g["files"]) for g in groups] == [2, 1]               # A、B 只是顺序不同，算一种
    lines = template_lines(groups)
    assert lines[0] == "  · 2 个和「A.xlsx」一样"
    assert lines[1] == "  · 1 个和「C.xlsx」一样：多「备注」；少「金额」"


# ---------- 隐藏内容（v3.0）：隐藏的 sheet 不发、隐藏列清空、隐藏行照常发 ----------

def _summary(logs):
    import json
    return json.loads([m for m in logs if m.startswith("[SUMMARY] ")][0][len("[SUMMARY] "):])


def _make_hidden_book(path, hide_split=False):
    """明细：D 列「身份证」隐藏（hide_split 时连「部门」C 列也藏）、第 4 行隐藏；另有一张隐藏的「底稿」。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.append(["2026年10月明细"])
    ws.merge_cells("A1:E1")
    ws.append(["工号", "姓名", "部门", "身份证", "金额"])
    ws.append(["001", "张三", "销售部", "110101X", 100])
    ws.append(["002", "李四", "销售部", "110102X", 200])
    ws.append(["003", "王五", "技术部", "110103X", 300])
    ws.column_dimensions["D"].hidden = True
    if hide_split:
        ws.column_dimensions["C"].hidden = True
    ws.row_dimensions[4].hidden = True
    ws2 = wb.create_sheet("底稿")
    ws2.append(["部门", "说明"])
    ws2.append(["销售部", "不该发出去"])
    ws2.sheet_state = "hidden"
    wb.save(path)


def test_hidden_content_default(tmp_path):
    src = tmp_path / "明细.xlsx"
    _make_hidden_book(src)
    logs = []
    res = run_split(_base_cfg(str(src), str(tmp_path / "out"), make_zip=False), log_fn=logs.append)
    wb = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx"))
    assert wb.sheetnames == ["明细"]                                  # 隐藏的「底稿」没发
    ws = wb.active
    assert ws["A1"].value == "2026年10月明细"                         # 跨到可见列的大标题不清
    assert [ws.cell(row=r, column=4).value for r in range(2, 5)] == [None, None, None]   # 列名和内容都清空
    assert ws.column_dimensions["D"].hidden                          # 仍然隐藏
    assert [ws.cell(row=r, column=2).value for r in (3, 4)] == ["张三", "李四"]   # 隐藏的那行照常发
    assert not any(d.hidden for d in ws.row_dimensions.values())     # 产出里不再隐藏
    s = _summary(logs)
    assert (s["hidden_sheets"], s["hidden_cols"], s["hidden_rows"]) == (1, 1, 1)
    assert s["failed_files"] == 0 and s["skipped_sheets"] == 0
    text = "\n".join(logs)
    assert "「底稿」是隐藏的，没有发出" in text and "「身份证」" in text and "1 行是隐藏的" in text


def test_hidden_split_column_still_groups(tmp_path):
    """拆分字段本身被隐藏：先分组再清空，照样拆得开。"""
    src = tmp_path / "明细.xlsx"
    _make_hidden_book(src, hide_split=True)
    res = run_split(_base_cfg(str(src), str(tmp_path / "out"), make_zip=False), log_fn=lambda m: None)
    assert sorted(f for f in os.listdir(res) if f.endswith(".xlsx")) == ["技术部.xlsx", "销售部.xlsx"]
    ws = openpyxl.load_workbook(os.path.join(res, "技术部.xlsx")).active
    assert ws["B3"].value == "王五" and ws["C3"].value is None and ws.column_dimensions["C"].hidden


def test_keep_hidden_sends_everything(tmp_path):
    src = tmp_path / "明细.xlsx"
    _make_hidden_book(src)
    logs = []
    res = run_split(_base_cfg(str(src), str(tmp_path / "out"), make_zip=False, keep_hidden=True),
                    log_fn=logs.append)
    wb = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx"))
    assert wb.sheetnames == ["明细", "底稿"]
    assert all(ws.sheet_state == "visible" for ws in wb.worksheets)
    ws = wb["明细"]
    assert [ws.cell(row=r, column=4).value for r in range(2, 5)] == ["身份证", "110101X", "110102X"]
    assert ws.column_dimensions["D"].hidden
    assert wb["底稿"]["B2"].value == "不该发出去"
    assert _summary(logs)["hidden_sheets"] == 1


def test_list_columns_skip_hidden_sheet(tmp_path):
    src = tmp_path / "a.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "底稿"
    ws.append(["网格", "备注"])
    ws.append(["城东", "x"])
    ws.sheet_state = "hidden"
    ws2 = wb.create_sheet("明细")
    ws2.append(["部门", "姓名"])
    ws2.append(["销售部", "张三"])
    wb.save(src)
    assert list_columns(str(src), {}) == ["部门", "姓名"]
    assert list_values(str(src), {}, "网格") == []
    assert list_columns(str(src), {"keep_hidden": True}) == ["网格", "备注"]


def test_file_with_data_only_in_hidden_sheet_is_not_a_failure(tmp_path):
    """看得见的 sheet 是空的、数据在隐藏 sheet 里：按设置不发，不算「失败文件」。"""
    inp = tmp_path / "in"
    inp.mkdir()
    _make_book(inp / "A.xlsx", [["001", "张三", "销售部", "北京", 100]])
    wb = openpyxl.Workbook()
    wb.active.title = "空白"
    ws = wb.create_sheet("数据")
    ws.append(["部门"])
    ws.append(["销售部"])
    ws.sheet_state = "hidden"
    wb.save(inp / "B.xlsx")
    logs = []
    run_split(_base_cfg(str(inp), str(tmp_path / "out")), log_fn=logs.append)
    s = _summary(logs)
    assert s["failed_files"] == 0 and s["hidden_sheets"] == 1 and s["rows"] == 1


# ---------- 行数对账与「请核对_无归属」（v3.0）----------

def _recon_ok(s):
    return s["source_rows"] == s["routed"] + s["skip_rows"] + s["unassigned"] + s["not_selected"]


def _make_rows_book(path, rows, header=("工号", "姓名", "部门", "金额")):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(list(header))
    for r in rows:
        ws.append(r)
    wb.save(path)


def test_unassigned_rows_go_to_review_file(tmp_path):
    src = tmp_path / "名单.xlsx"
    _make_rows_book(src, [
        ["001", "张三", "销售部", 100],
        ["002", "李四", None, 200],          # 没写部门 → 请核对
        [None, None, None, None],            # 空行：不计
        ["003", "王五", "-", 50],             # 只有符号 → 请核对
        ["004", "赵六", "技术部", 300],
        [None, "合计", None, 650],            # 写着合计 → 合计等
        [None, None, None, 650],             # 最后一行只有数字 → 合计等
    ])
    logs = []
    res = run_split(_base_cfg(str(src), str(tmp_path / "out"), make_zip=False), log_fn=logs.append)
    s = _summary(logs)
    assert (s["source_rows"], s["routed"], s["skip_rows"], s["unassigned"], s["not_selected"]) == (6, 2, 2, 2, 0)
    assert _recon_ok(s)
    assert s["files"] == 2 and s["groups"] == 2                     # 请核对不算组数、不算文件
    assert s["unassigned_file"] == os.path.join(res, "请核对_无归属.xlsx")
    wb = openpyxl.load_workbook(s["unassigned_file"])
    assert wb.sheetnames == ["名单·Sheet"]
    ws = wb.active
    assert [c.value for c in ws[1]] == ["工号", "姓名", "部门", "金额"]
    assert [ws.cell(row=r, column=2).value for r in (2, 3)] == ["李四", "王五"]
    assert any("行数核对：源表 6 行 = 拆出 2 + 合计等 2 + 没写分组 2" in m for m in logs), logs


def test_no_review_file_when_every_row_has_a_group(two_books):
    inp, out = two_books
    logs = []
    res = run_split(_base_cfg(inp, out), log_fn=logs.append)
    s = _summary(logs)
    assert s["unassigned"] == 0 and s["unassigned_file"] == "" and _recon_ok(s)
    assert s["skip_rows"] == 1                                       # A 表那行「合计」（部门是空的）
    assert not os.path.exists(os.path.join(res, "请核对_无归属.xlsx"))


def test_review_file_one_page_per_source_sheet_and_not_zipped(tmp_path):
    inp = tmp_path / "in"
    (inp / "子").mkdir(parents=True)
    _make_rows_book(inp / "A.xlsx", [["1", "张三", "销售部", 1], ["2", "李四", None, 2]])
    _make_rows_book(inp / "子" / "A.xlsx", [["3", "王五", None, 3], ["4", "赵六", "技术部", 4]])
    logs = []
    res = run_split(_base_cfg(str(inp), str(tmp_path / "out"), merge_across_files=False), log_fn=logs.append)
    wb = openpyxl.load_workbook(os.path.join(res, "请核对_无归属.xlsx"))
    assert len(wb.sheetnames) == 2                                   # 两个源文件同名，页名自动去重
    import zipfile
    for z in ("销售部.zip", "技术部.zip"):
        with zipfile.ZipFile(os.path.join(res, z)) as zf:
            assert not any("请核对" in n for n in zf.namelist())
    assert _recon_ok(_summary(logs))


def test_group_named_like_review_file_gets_suffix(tmp_path):
    src = tmp_path / "a.xlsx"
    _make_rows_book(src, [["1", "张三", "请核对_无归属", 1], ["2", "李四", None, 2]])
    res = run_split(_base_cfg(str(src), str(tmp_path / "out"), make_zip=False), log_fn=lambda m: None)
    names = sorted(f for f in os.listdir(res) if f.endswith(".xlsx"))
    assert names == ["请核对_无归属(2).xlsx", "请核对_无归属.xlsx"]
    ws = openpyxl.load_workbook(os.path.join(res, "请核对_无归属(2).xlsx")).active
    assert ws["B2"].value == "张三"


def test_reconcile_with_selected_values(tmp_path):
    src = tmp_path / "a.xlsx"
    _make_rows_book(src, [["1", "a", "销售一部", 1], ["2", "b", "销售二部", 2], ["3", "c", "技术部", 3],
                          ["4", "d", "销售一部、销售二部", 4]])
    logs = []
    run_split(_base_cfg(str(src), str(tmp_path / "o1"), make_zip=False, selected_values=["销售一部"]),
              log_fn=logs.append)
    s = _summary(logs)
    assert (s["routed"], s["not_selected"]) == (1, 3) and _recon_ok(s)
    # 包含匹配：一行进了两个分组，对账里只算一行
    logs = []
    run_split(_base_cfg(str(src), str(tmp_path / "o2"), make_zip=False, exact_match=False,
                        selected_values=["销售一部", "销售二部"]), log_fn=logs.append)
    s = _summary(logs)
    assert (s["source_rows"], s["routed"], s["not_selected"]) == (4, 3, 1) and _recon_ok(s)
    assert s["rows"] == 4                                            # 写出的行数：第 4 行两组各一份


def test_hidden_sheet_rows_not_in_reconciliation(tmp_path):
    src = tmp_path / "明细.xlsx"
    _make_hidden_book(src)
    logs = []
    run_split(_base_cfg(str(src), str(tmp_path / "out"), make_zip=False), log_fn=logs.append)
    s = _summary(logs)
    assert s["source_rows"] == 3 and _recon_ok(s)                     # 「底稿」那一行不计


def test_single_group_warns(tmp_path):
    src = tmp_path / "a.xlsx"
    _make_rows_book(src, [["1", "张三", "销售部", 1], ["2", "李四", "销售部", 2]])
    logs = []
    run_split(_base_cfg(str(src), str(tmp_path / "out"), make_zip=False), log_fn=logs.append)
    assert any("只拆出了 1 组「销售部」" in m for m in logs)


# ---------- 下拉 / 数据验证（v3.0）----------

def _dvs(ws):
    """产出 sheet 的数据验证 → {(类型, 公式): 区域}。"""
    return {(dv.type, dv.formula1): str(dv.sqref) for dv in ws.data_validations.dataValidation}


def _dv_book(path, header=("工号", "姓名", "部门", "结果", "分数"), rows=None):
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.append(list(header))
    for r in rows or [["1", "张三", "销售部", None, None], ["2", "李四", "技术部", None, None],
                      ["3", "王五", "销售部", None, None]]:
        ws.append(r)
    opt = wb.create_sheet("选项")
    for v in ("已走访", "未走访", "待定"):
        opt.append([v])
    opt.sheet_state = "hidden"
    col = {h: get_column_letter(i) for i, h in enumerate(header, start=1)}
    lit = DataValidation(type="list", formula1='"是,否"', showErrorMessage=True)
    lit.add(f"{col['部门']}1:{col['部门']}200")              # 连表头那格也有（表头区照原位置带上）
    ref = DataValidation(type="list", formula1="选项!$A$1:$A$3")
    ref.add(f"{col['结果']}2:{col['结果']}200")
    whole = DataValidation(type="whole", operator="between", formula1="0", formula2="100")
    whole.add(f"{col['分数']}2:{col['分数']}200")
    custom = DataValidation(type="custom", formula1=f"LEN({col['工号']}2)<5")
    custom.add(f"{col['工号']}2:{col['工号']}200")
    for dv in (lit, ref, whole, custom):
        ws.add_data_validation(dv)
    wb.save(path)


def test_validations_follow_rows(tmp_path):
    src = tmp_path / "走访.xlsx"
    _dv_book(src)
    logs = []
    res = run_split(_base_cfg(str(src), str(tmp_path / "out"), make_zip=False), log_fn=logs.append)
    sales = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx"))["明细"]
    assert _dvs(sales) == {
        ("list", '"是,否"'): "C1:C3",                          # 表头格 + 两行数据
        ("list", '"已走访,未走访,待定"'): "D2:D3",              # 引用隐藏「选项」表的，选项直接写进下拉
        ("whole", "0"): "E2:E3",
    }                                                          # 自定义公式那个去掉了
    tech = openpyxl.load_workbook(os.path.join(res, "技术部.xlsx"))["明细"]
    assert _dvs(tech)[("list", '"是,否"')] == "C1:C2"
    text = "\n".join(logs)
    assert "1 个下拉的选项引用了别处的格子" in text and "1 个数据验证引用了别处的格子" in text


def test_validations_follow_columns_when_merging(tmp_path):
    inp = tmp_path / "in"
    inp.mkdir()
    _dv_book(inp / "A.xlsx")
    _dv_book(inp / "B.xlsx", header=("姓名", "工号", "结果", "部门", "分数"),
             rows=[["赵六", "4", None, "销售部", None]])
    res = run_split(_base_cfg(str(inp), str(tmp_path / "out"), make_zip=False), log_fn=lambda m: None)
    ws = openpyxl.load_workbook(os.path.join(res, "销售部.xlsx"))["明细"]
    assert ws["B4"].value == "赵六"                             # B 表的行按列名对齐，落在第 4 行
    dvs = _dvs(ws)
    assert dvs[("list", '"是,否"')] == "C1:C4"                  # B 表「部门」在 D 列，跟着列名回到 C 列
    assert dvs[("list", '"已走访,未走访,待定"')] == "D2:D4"


def test_list_from_range_limits():
    from core.splitter import _list_from_range
    wb = openpyxl.Workbook()
    ws = wb.active
    for v in ("甲", "乙", "甲", None, 3.0):
        ws.append([v])
    ws["B1"] = "有,逗号"
    assert _list_from_range(wb, ws, "$A$1:$A$5") == '"甲,乙,3"'   # 去重、跳过空格、3.0 当 3
    assert _list_from_range(wb, ws, "B1") is None                # 含英文逗号：没法写成字面下拉
    assert _list_from_range(wb, ws, "不存在!A1:A3") is None
    for i in range(1, 300):
        ws.cell(row=i, column=3, value=f"选项{i}")
    assert _list_from_range(wb, ws, "C1:C299") is None           # 超过 255 个字
    assert ws.max_column == 3                                    # 只读格子，没把源表撑大
