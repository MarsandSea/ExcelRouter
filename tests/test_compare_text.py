"""gui/compare_text.py 的测试：关键列下拉选项、结果卡片文案（不需要显示器）。

Copyright (c) 2026 AbeLin · MIT License
"""
from gui import compare_text as CT

INFO = {"columns": ["序号", "工号", "姓名", "部门"],
        "keys": [{"cols": ["工号"], "matched": 120}, {"cols": ["姓名", "部门"], "matched": 118}]}


def test_key_options_puts_suggestions_first():
    labels, mapping, default = CT.key_options(INFO)
    assert labels[0] == "工号（推荐，能对上 120 行）"
    assert labels[1] == "姓名 + 部门（也可以，能对上 118 行）"
    assert labels[2:] == ["序号", "姓名", "部门"]          # 已推荐过的「工号」不重复列出
    assert mapping[labels[1]] == ["姓名", "部门"]
    assert default == labels[0]


def test_key_options_remembers_saved_key():
    _labels, mapping, default = CT.key_options(INFO, saved=["部门"])
    assert mapping[default] == ["部门"]
    _l, _m, default = CT.key_options(INFO, saved=["早就没了的列"])
    assert default.startswith("工号")


def test_key_options_without_suggestion():
    labels, _m, default = CT.key_options({"columns": ["姓名", "金额"], "keys": []})
    assert labels == ["姓名", "金额"] and default == ""


SUMMARY = {"key": ["工号"], "matched": 5, "only_a": 1, "only_b": 12, "changed_rows": 9, "changed_cells": 11,
           "unpaired": 2, "per_col": {"金额": 7, "部门": 4}, "only_a_cols": [], "only_b_cols": ["备注"],
           "preview": {"only_a": ["E006"], "only_b": [f"E1{i:02d}" for i in range(8)],
                       "changed": [["E002", ["金额"]], ["E003", ["金额", "部门"]]]}}


def test_result_lines():
    lines = CT.result_lines(SUMMARY)
    texts = [t for t, _k in lines]
    assert texts[0].startswith("➕ 新增 12 行") and texts[0].endswith("E107……")
    assert texts[1] == "➖ 没了 1 行（原表有、新表没有）：E006"
    assert "金额 7 行、部门 4 行" in texts[2]
    assert "    · E003：改了 金额、部门" in texts
    assert any("其余见核对结果" in t for t in texts)             # 改了 9 行只示例 2 行
    assert any(k == "warn" and "2 行关键列重复" in t for t, k in lines)
    assert any("只有新表有的列不比较：备注" in t for t in texts)


def test_result_lines_all_same():
    s = {"key": ["工号"], "matched": 3, "only_a": 0, "only_b": 0, "changed_rows": 0, "changed_cells": 0}
    lines = CT.result_lines(s)
    assert [k for _t, k in lines] == ["info", "info", "info"]
    text, warn = CT.banner_lines(s, "/x/核对结果")
    assert "内容一样" in text and not warn and text.endswith("📁 /x/核对结果")


def test_banner_warns_on_unpaired():
    text, warn = CT.banner_lines(SUMMARY, "/x")
    assert warn and "新增 12 · 没了 1 · 改过 9 行" in text
