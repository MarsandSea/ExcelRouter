# Copyright (c) 2026 AbeLin · MIT License
"""下游直接调用的「私有」函数契约。

内网版技能会原样带上 core/，其中的轻量拆分路径不经过 process_file，
而是直接调用下面这些下划线开头的函数；PDF 部分直接复用字体候选表。
改这些函数时：已有参数的名字和顺序别动，新增参数必须带默认值——
否则主仓测试照样全绿，同步到下游才炸。真要改，先改下游再改这里。
"""
import inspect

import pytest

from core import pdf_dist, splitter

# 名字 → 现有参数（按顺序）。新增参数可以，但只能加在后面且带默认值。
CONTRACT = {
    (splitter, "_read_header_format"): ["ws", "h_idx"],
    (splitter, "_dedupe_path"): ["save_path", "used_paths"],
    (splitter, "_append_rows"): ["sheet_info", "rows_df", "preserve", "ws_src",
                                 "src_max_col", "style_cache", "heartbeat",
                                 "formula_ws", "col_map"],
    (splitter, "_summary_key_path"): ["output_root", "primary", "src_stem",
                                      "single_file", "merge"],
    (splitter, "_person_key_path"): ["output_root", "primary", "person",
                                     "src_stem", "merge"],
    (splitter, "_OutputBook"): ["save_path"],
    (pdf_dist, "_find_cjk_font"): [],
    (pdf_dist, "_font_usable"): ["path"],
    (pdf_dist, "_make_watermark_pdf"): ["text", "w_pt", "h_pt"],
}

OUTPUTBOOK_METHODS = {
    "get_or_create_sheet": ["sheet_name", "h_idx", "header_cells", "merges", "widths"],
    "save": [],
}


def _check(fn, expected):
    params = [p for p in inspect.signature(fn).parameters.values()
              if p.name != "self"]
    names = [p.name for p in params]
    assert names[:len(expected)] == expected, f"{fn.__qualname__} 参数变了: {names}"
    for p in params[len(expected):]:
        assert p.default is not inspect.Parameter.empty or p.kind in (
            p.VAR_POSITIONAL, p.VAR_KEYWORD), \
            f"{fn.__qualname__} 新增参数 {p.name} 没有默认值，下游会调用失败"


@pytest.mark.parametrize("key", list(CONTRACT), ids=lambda k: k[1])
def test_private_signature_stable(key):
    mod, name = key
    assert hasattr(mod, name), f"{mod.__name__}.{name} 被删了，下游还在用"
    _check(getattr(mod, name), CONTRACT[key])


@pytest.mark.parametrize("name", list(OUTPUTBOOK_METHODS))
def test_outputbook_methods_stable(name):
    assert hasattr(splitter._OutputBook, name)
    _check(getattr(splitter._OutputBook, name), OUTPUTBOOK_METHODS[name])


def test_font_candidates_are_shared_lists():
    """下游水印直接读这两张表，别手抄一份。"""
    for name in ("_FONT_CANDIDATES", "_LINUX_FONT_CANDIDATES"):
        val = getattr(pdf_dist, name)
        assert isinstance(val, list) and val and all(isinstance(x, str) for x in val)
