"""把拆分输出目录整理成可比较的数据：文件清单、每个工作簿的 sheet、单元格的值与类型、格式、合并单元格、列宽，ZIP 成员。

不比较的：运行日志.txt（日志措辞会变）、请核对_无归属.xlsx（v3.0 新增的产出，由专门的测试断言）。

Copyright (c) 2026 AbeLin · MIT License
"""
import datetime
import os
import zipfile

import openpyxl

SKIP_FILES = {"运行日志.txt", "请核对_无归属.xlsx"}


def _value(v):
    """单元格的值 → 可存进 JSON 的形式。整数形式的小数记成整数（3.0 与 3 在表格里显示一样）。"""
    if v is None or isinstance(v, (bool, int, str)):
        return v
    if isinstance(v, float):
        return int(v) if v.is_integer() else round(v, 9)
    if isinstance(v, datetime.datetime):
        return "dt:" + v.isoformat()
    if isinstance(v, datetime.date):
        return "d:" + v.isoformat()
    if isinstance(v, datetime.time):
        return "t:" + v.isoformat()
    if isinstance(v, datetime.timedelta):
        return f"td:{v.total_seconds()}"
    return f"{type(v).__name__}:{v}"


def _color(c):
    return c.rgb if c is not None and getattr(c, "type", None) == "rgb" else None


def _style(cell):
    if not cell.has_style:
        return None
    f, fill = cell.font, cell.fill
    return [cell.number_format, bool(f.b), bool(f.i), _color(f.color), fill.fill_type,
            _color(fill.fgColor) if fill.fill_type else None, cell.alignment.horizontal]


def book(path):
    wb = openpyxl.load_workbook(path)
    try:
        sheets = []
        for ws in wb.worksheets:
            cells = []
            for row in ws.iter_rows():
                for cell in row:
                    st = _style(cell)
                    if cell.value is None and st is None:
                        continue
                    cells.append([cell.coordinate, _value(cell.value), cell.data_type, st])
            widths = {k: round(d.width, 4) for k, d in sorted(ws.column_dimensions.items())
                      if d.customWidth and d.width is not None}
            sheets.append({"title": ws.title, "cells": cells,
                           "merged": sorted(str(m) for m in ws.merged_cells.ranges), "widths": widths})
        return sheets
    finally:
        wb.close()


def take(output_dir):
    """输出目录 → {相对路径: 内容}。"""
    out = {}
    for root, _dirs, files in os.walk(output_dir):
        for fn in sorted(files):
            if fn in SKIP_FILES:
                continue
            path = os.path.join(root, fn)
            rel = os.path.relpath(path, output_dir).replace(os.sep, "/")
            low = fn.lower()
            if low.endswith(".xlsx"):
                out[rel] = {"book": book(path)}
            elif low.endswith(".zip"):
                with zipfile.ZipFile(path) as z:
                    out[rel] = {"zip": sorted(n for n in z.namelist()
                                              if not n.endswith("/") and n.split("/")[-1] not in SKIP_FILES)}
            else:
                out[rel] = {"file": True}
    return out


def run_scenario(name, root):
    """在 root 下生成场景输入、跑一次 run_split，返回输出目录的快照。"""
    from core.splitter import run_split

    from golden.samples import make_config
    cfg = make_config(name, root)
    output = run_split(cfg, log_fn=lambda m: None)
    return take(output)


def diff(expected, actual, path="", out=None, limit=40):
    """逐层比较，返回前 limit 处差异的说明。"""
    out = [] if out is None else out
    if len(out) >= limit:
        return out
    if isinstance(expected, dict) and isinstance(actual, dict):
        for k in sorted(set(expected) | set(actual), key=str):
            if k not in actual:
                out.append(f"{path}/{k}: 缺少（期望有）")
            elif k not in expected:
                out.append(f"{path}/{k}: 多出来（期望没有）")
            else:
                diff(expected[k], actual[k], f"{path}/{k}", out, limit)
    elif isinstance(expected, list) and isinstance(actual, list):
        if len(expected) != len(actual):
            out.append(f"{path}: 长度 {len(expected)} → {len(actual)}")
        for i, (e, a) in enumerate(zip(expected, actual)):
            diff(e, a, f"{path}[{i}]", out, limit)
    elif expected != actual:
        out.append(f"{path}: {expected!r} → {actual!r}")
    return out
