"""通报表两个选项（v3.2）：拆完「序号」每组从 1 重新编号（renumber_seq）、每组重算合计行（group_totals）。

两个都默认关：关着时产出和以前一模一样（序号照抄、源表合计行不进任何一组）。

Copyright (c) 2026 AbeLin
MIT License
"""
import glob
import os

import openpyxl
from openpyxl.styles import Font
import pytest

from core.splitter import report_hints, run_split

HEAD = ["序号", "镇区", "网格", "目标", "完成", "完成率", "负责人"]
TOWNS = ["城区", "东镇", "西镇"]


def _rows(n=12, start=1):
    out = []
    for i in range(start, start + n):
        g, d = 50 + i * 3, 10 + i * 2
        out.append([i, TOWNS[i % 3], f"网格{i}", g, d, d / g, f"人{i}"])
    return out


def _total(rows, label="合计"):
    g = sum(r[3] for r in rows)
    d = sum(r[4] for r in rows)
    return [label, None, None, g, d, d / g, len(rows)]


def _book(path, body, title="10月业务通报"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "通报"
    ws.append([title])
    ws.append(HEAD)
    for r in body:
        ws.append(r)
        if r and r[0] in ("合计", "小计") or (isinstance(r[0], str) and r[0].endswith("小计")):
            for c in ws[ws.max_row]:
                c.font = Font(bold=True)
    wb.save(path)


def _split(src, out, **opts):
    cfg = {"input_path": str(src), "output_path": str(out), "split_column": "镇区", **opts}
    logs = []
    res = run_split(cfg, log_fn=logs.append)
    books = {}
    for f in glob.glob(os.path.join(res, "**", "*.xlsx"), recursive=True):
        if "请核对" in f:
            continue
        ws = openpyxl.load_workbook(f).active
        books[os.path.splitext(os.path.basename(f))[0]] = (
            ws, [[c.value for c in row] for row in ws.iter_rows(min_row=3)])
    return books, "\n".join(logs)


@pytest.fixture
def report(tmp_path):
    body = _rows()
    src = tmp_path / "通报.xlsx"
    _book(src, body + [_total(body)])
    return src, body


def test_default_off_keeps_old_behavior(report, tmp_path):
    src, body = report
    books, _ = _split(src, tmp_path / "out")
    ws, vals = books["城区"]
    assert [v[0] for v in vals] == [r[0] for r in body if r[1] == "城区"]    # 序号照抄
    assert all(v[0] != "合计" for v in vals)                                 # 不补合计


def test_renumber_and_totals(report, tmp_path):
    src, body = report
    books, log = _split(src, tmp_path / "out", renumber_seq=True, group_totals=True)
    assert set(books) == set(TOWNS)
    for town in TOWNS:
        ws, vals = books[town]
        mine = [r for r in body if r[1] == town]
        assert [v[0] for v in vals[:-1]] == list(range(1, len(mine) + 1))
        last = vals[-1]
        assert last[0] == "合计"
        assert last[3] == sum(r[3] for r in mine) and last[4] == sum(r[4] for r in mine)
        # 完成率是「总完成 ÷ 总目标」，不是各行完成率的平均
        assert last[5] == pytest.approx(last[4] / last[3])
        assert last[6] == len(mine)                    # 负责人那一格是人数
        assert ws.cell(ws.max_row, 4).font.b           # 合计行格式照源表
        assert [v[2] for v in vals[:-1]] == [r[2] for r in mine]   # 其余列照抄
    assert "重新编号" in log and "合计行" in log


def test_only_renumber(report, tmp_path):
    src, body = report
    books, _ = _split(src, tmp_path / "out", renumber_seq=True)
    _ws, vals = books["东镇"]
    assert [v[0] for v in vals] == list(range(1, len(vals) + 1))
    assert all(v[0] != "合计" for v in vals)


def test_total_on_top(tmp_path):
    body = _rows()
    src = tmp_path / "通报.xlsx"
    _book(src, [_total(body)] + body)
    books, _ = _split(src, tmp_path / "out", group_totals=True)
    _ws, vals = books["西镇"]
    mine = [r for r in body if r[1] == "西镇"]
    assert vals[0][0] == "合计" and vals[0][3] == sum(r[3] for r in mine)
    assert len(vals) == len(mine) + 1


def test_subtotals_replaced_by_one_total(tmp_path):
    """源表每个镇区后面跟一行小计、最后一行总计：每组只留一行重算的合计，小计不重复。"""
    body = []
    allrows = _rows(15)
    for town in TOWNS:
        mine = [r for r in allrows if r[1] == town]
        body += mine + [[None, town, f"{town}小计"] + _total(mine)[3:]]
    body.append(_total(allrows, "总计"))
    src = tmp_path / "通报.xlsx"
    _book(src, body)
    books, _ = _split(src, tmp_path / "out", renumber_seq=True, group_totals=True)
    for town in TOWNS:
        _ws, vals = books[town]
        mine = [r for r in allrows if r[1] == town]
        assert len(vals) == len(mine) + 1
        assert vals[-1][0] == "总计" and vals[-1][4] == sum(r[4] for r in mine)


def test_subtotal_rows_keep_blank_seq_when_only_renumber(tmp_path):
    """只开重新编号：分进组里的小计行原来没写序号，照旧空着、不占号。"""
    allrows = _rows(9)
    body = []
    for town in TOWNS:
        mine = [r for r in allrows if r[1] == town]
        body += mine + [[None, town, "小计"] + _total(mine)[3:]]
    src = tmp_path / "通报.xlsx"
    _book(src, body)
    books, _ = _split(src, tmp_path / "out", renumber_seq=True)
    _ws, vals = books["城区"]
    assert [v[0] for v in vals] == [1, 2, 3, None]


def test_merge_across_files_numbers_continue(tmp_path):
    inp = tmp_path / "in"
    inp.mkdir()
    a, b = _rows(6), _rows(6, start=7)
    _book(inp / "1月.xlsx", a + [_total(a)])
    _book(inp / "2月.xlsx", b + [_total(b)])
    books, _ = _split(inp, tmp_path / "out", merge_across_files=True, make_zip=False,
                      renumber_seq=True, group_totals=True)
    _ws, vals = books["城区"]
    mine = [r for r in a + b if r[1] == "城区"]
    assert [v[0] for v in vals[:-1]] == list(range(1, len(mine) + 1))
    assert [v[0] for v in vals].count("合计") == 1
    assert vals[-1][3] == sum(r[3] for r in mine)


def test_unknown_total_column_left_blank(tmp_path):
    body = _rows()
    tot = _total(body)
    tot[2] = None
    tot.append(None)
    rows = [r + [r[3] * 2] for r in body]                 # 第 8 列「最高」：合计写的是最大值
    src = tmp_path / "通报.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["标题"])
    ws.append(HEAD + ["最高"])
    for r in rows:
        ws.append(r)
    ws.append(tot[:7] + [max(r[7] for r in rows)])
    wb.save(src)
    books, log = _split(src, tmp_path / "out", group_totals=True)
    _ws, vals = books["城区"]
    assert vals[-1][7] is None
    assert "「最高」" in log and "留空" in log


def test_fast_mode(report, tmp_path):
    src, body = report
    books, _ = _split(src, tmp_path / "out", preserve_format=False, renumber_seq=True, group_totals=True)
    _ws, vals = books["城区"]
    mine = [r for r in body if r[1] == "城区"]
    assert vals[-1][3] == sum(r[3] for r in mine)
    assert [v[0] for v in vals[:-1]] == list(range(1, len(mine) + 1))


def test_no_seq_name_no_renumber(tmp_path):
    """「编号」是编码不是序号：不重新编。"""
    body = _rows()
    src = tmp_path / "通报.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["编号"] + HEAD[1:])
    for r in body:
        ws.append(r)
    wb.save(src)
    assert report_hints(str(src), {}) == {"seq_column": "", "total_rows": 0}
    books, _ = _split(src, tmp_path / "out", renumber_seq=True, group_totals=True)
    wsb = openpyxl.load_workbook(next(iter(glob.glob(str(tmp_path / "out" / "*" / "城区.xlsx"))))).active
    assert [r[0] for r in wsb.iter_rows(min_row=2, values_only=True)] == [r[0] for r in body if r[1] == "城区"]


def test_report_hints(report):
    src, _ = report
    assert report_hints(str(src), {}) == {"seq_column": "序号", "total_rows": 1}


def test_merged_total_label(tmp_path):
    """合计行前三格合并写「合计」：每组的合计行同样合并、边框照源表。"""
    from openpyxl.styles import Border, Side
    body = _rows()
    src = tmp_path / "通报.xlsx"
    _book(src, body + [_total(body)])
    wb = openpyxl.load_workbook(src)
    ws = wb.active
    r = ws.max_row
    thin = Side(style="thin")
    for c in range(1, 8):
        ws.cell(r, c).border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    wb.save(src)
    books, _ = _split(src, tmp_path / "out", renumber_seq=True, group_totals=True)
    for town in TOWNS:
        wsb, vals = books[town]
        last = wsb.max_row
        assert [str(m) for m in wsb.merged_cells.ranges] == [f"A{last}:C{last}"]
        assert wsb.cell(last, 1).value == "合计"
        assert wsb.cell(last, 3).border.right.style == "thin"
        assert vals[-1][3] == sum(x[3] for x in body if x[1] == town)


def test_merged_total_default_off_unchanged(tmp_path):
    """不勾：合并的合计行照旧不发，产出里没有合并区。"""
    body = _rows()
    src = tmp_path / "通报.xlsx"
    _book(src, body + [_total(body)])
    wb = openpyxl.load_workbook(src)
    ws = wb.active
    ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=3)
    wb.save(src)
    books, _ = _split(src, tmp_path / "out")
    wsb, vals = books["城区"]
    assert not wsb.merged_cells.ranges and all(v[0] != "合计" for v in vals)


def test_vertically_merged_split_column(tmp_path):
    """「镇区」一列几行合并成一格：每一行都分进对应的组、产出里每行都写着镇区，不进「请核对」。"""
    src = tmp_path / "通报.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["标题"])
    ws.append(HEAD)
    i = 0
    for town in TOWNS:
        start = ws.max_row + 1
        for _ in range(4):
            i += 1
            ws.append([i, town, f"网格{i}", 10 * i, i, 0.1, f"人{i}"])
        ws.merge_cells(start_row=start, start_column=2, end_row=start + 3, end_column=2)
        ws.cell(start, 2).font = Font(bold=True)
    ws.append(["备注：本表为样例"])
    ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=7)   # 横着的不动
    wb.save(src)
    out = tmp_path / "out"
    books, log = _split(src, out)
    assert "3 处上下合并" in log
    for town in TOWNS:
        wsb, vals = books[town]
        assert len(vals) == 4 and all(v[1] == town for v in vals)
        assert all(wsb.cell(r, 2).font.b for r in range(3, 7))
    # 横着合并的备注行：没写分组，照旧进「请核对」，没被抄成 7 格
    review = glob.glob(str(out / "*" / "请核对_无归属.xlsx"))
    assert review
    row = next(openpyxl.load_workbook(review[0]).active.iter_rows(min_row=3, values_only=True))
    assert row[0] == "备注：本表为样例" and all(v is None for v in row[1:])
