"""core/recipe.py：方案的存、读、套用、列检查；方案文件按外来数据处理。

Copyright (c) 2026 AbeLin · MIT License
"""
import json

import pytest

from core import recipe as R

CFG = {"input_path": "C:/私人目录/九月.xlsx", "output_path": "D:/out", "due": "10月15日", "ui_mode": "excel",
       "split_column": "网格", "to_person": True, "person_column": "客户经理", "make_zip": False,
       "fill_columns": "走访日期（日期）、走访结果（已走访/未走访）", "skip_values": ["合计", ""],
       "value_alias_map": {"东区": ["东城"]}, "header_row": 3, "header_mode": "row", "window_geometry": "1x1",
       "collect_expect": "东区\n西区"}


def test_round_trip_keeps_settings_not_paths(tmp_path):
    r = R.make_recipe(CFG, "excel", "九月网格走访", note="15 号收", columns=["网格", "客户名称"],
                      input_name="C:/私人目录/九月.xlsx", app="3.0.0")
    path = R.save_recipe(r, str(tmp_path / "方案" / R.recipe_filename("九月/网格:走访")))
    text = open(path, encoding="utf-8").read()
    assert "私人目录" not in text and "D:/out" not in text and "10月15日" not in text and "window" not in text
    assert R.is_recipe(path)
    loaded, warns = R.load_recipe(path)
    assert not warns and loaded["hints"] == {"input_name": "九月.xlsx", "columns": ["网格", "客户名称"]}
    new, changed, _notes = R.apply_recipe({"input_path": "E:/十月.xlsx", "due": "旧的", "split_column": "部门"},
                                          loaded)
    for k in R.RECIPE_KEYS["excel"]:
        if k in CFG:
            assert new[k] == CFG[k], k
    assert new["input_path"] == "E:/十月.xlsx" and new["due"] == "" and new["ui_mode"] == "excel"
    assert "split_column" in changed and "collect_expect" not in new


def test_collect_recipe_keeps_name_list_lines(tmp_path):
    r = R.make_recipe(dict(CFG, collect_check=True, collect_by="网格"), "collect", "走访查收")
    assert r["settings"]["collect_expect"] == "东区\n西区"
    path = R.save_recipe(r, str(tmp_path / "a.json"))
    loaded, _ = R.load_recipe(path, mode="collect")
    new, _changed, notes = R.apply_recipe({"collect_record": "x.json", "collect_due": "y"}, loaded)
    assert new["collect_record"] == "" and new["collect_due"] == "" and notes == ["方案里有应交名单 2 个"]
    assert "应交名单 2 个" in R.summary(loaded)


def test_foreign_data_is_sanitized(tmp_path):
    data = {R.RECIPE_KEY: 1, "name": "x" * 500, "mode": "excel",
            "settings": {"split_column": ["网格", "部门"], "make_zip": "yes", "header_row": 99999,
                         "header_mode": "rm -rf", "skip_values": "合计", "fill_columns": "长" * 9000,
                         "value_alias_map": {"a": ["b"], "c": "d"}, "input_path": "C:/Windows", "future_key": 1},
            "hints": {"columns": "不是列表"}}
    p = tmp_path / "外来.json"
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    r, warns = R.load_recipe(str(p))
    s = r["settings"]
    assert "make_zip" not in s and "header_row" not in s and "header_mode" not in s and "skip_values" not in s
    assert s["split_column"] == "网格、部门" and len(s["fill_columns"]) == R.MAX_TEXT
    assert s["value_alias_map"] == {"a": ["b"]} and "input_path" not in s
    assert len(r["name"]) == R.MAX_NAME and r["hints"]["columns"] == []
    assert any("不支持" in w for w in warns) and any("make_zip" in w for w in warns)


def test_bad_files_raise_plain_errors(tmp_path):
    bad = tmp_path / "坏.json"
    bad.write_text("{不是 json", encoding="utf-8")
    with pytest.raises(R.RecipeError, match="坏了"):
        R.load_recipe(str(bad))
    rec = tmp_path / "下发记录.json"
    rec.write_text(json.dumps({"version": 1, "groups": {}}), encoding="utf-8")
    assert not R.is_recipe(str(rec))
    with pytest.raises(R.RecipeError, match="不是"):
        R.load_recipe(str(rec))
    ok = tmp_path / "拆分.json"
    R.save_recipe(R.make_recipe(CFG, "excel", "拆分"), str(ok))
    with pytest.raises(R.RecipeError, match="不能用来做"):
        R.load_recipe(str(ok), mode="collect")
    newer = tmp_path / "新版.json"
    newer.write_text(json.dumps({R.RECIPE_KEY: 9, "mode": "pdf", "settings": {}}), encoding="utf-8")
    assert any("新版本" in w for w in R.load_recipe(str(newer))[1])


def test_check_columns():
    r = R.make_recipe(CFG, "excel", "x", columns=["网格", "客户名称", "客户经理", "走访日期"])
    assert R.check_columns(r, ["网格", "客户名称", "客户经理", "电话"]) == []
    w = R.check_columns(r, ["部门", "姓名", "金额"])
    assert any("拆分字段「网格」这张表里没有" in x for x in w) and any("到人字段" in x for x in w)
    assert any("不太一样" in x for x in w)
    no_person = R.make_recipe(dict(CFG, to_person=False), "excel", "x")
    assert not any("到人" in x for x in R.check_columns(no_person, ["网格"]))
