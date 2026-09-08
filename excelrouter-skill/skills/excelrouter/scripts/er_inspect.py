#!/usr/bin/env python
"""
excelrouter · 看表：拆分前先摸清一个文件/一批文件的表头行、字段名、某字段的取值。

Copyright (c) 2026 AbeLin
MIT License

这是 er_split.py 之前必跑的一步——不要凭表名或猜测直接决定拆分字段，
应该先跑这个脚本把 --column 的真实取值列出来，跟用户确认后再拆分。

多 sheet 文件（v2.6.1 起）：会列出「每一个」sheet 的表头行与列名；
用 --column 枚举取值时，会跨所有 sheet 搜索，并如实报告该字段出现在哪些 sheet。
某字段不在第一个 sheet、而在第二个 sheet 时，不再误报「不在列名里」。

用法：
  python er_inspect.py --input 表.xlsx
  python er_inspect.py --input 一批表/ --column 部门
  python er_inspect.py --input 表.xlsx --header-mode row --header-row 2

输出：stdout 一行 JSON，字段见下方 emit() 调用处。
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import bootstrap, build_config, check_deps, emit, emit_error  # noqa: E402

bootstrap()

import pandas as pd  # noqa: E402
from core.splitter import (  # noqa: E402
    normalize_to_xlsx, resolve_header_row, _find_col_index, _header_values,
)
from core.utils import apply_alias, is_skip_value, soft_clean  # noqa: E402


def _is_candidate_excel(name):
    low = name.lower()
    return low.endswith((".xlsx", ".xls")) and not name.startswith("~$") and not name.endswith("__tmp__.xlsx")


def _iter_input_files(input_path, output_root=None):
    """枚举待处理的 Excel 文件；目录场景镜像 core.splitter.run_split 的跳过规则
    （output_root 子树整体跳过、临时文件/锁文件跳过），保证数字不虚高。"""
    if os.path.isfile(input_path):
        return [input_path]
    files = []
    out_abs = os.path.abspath(output_root) if output_root else None
    for root, _dirs, names in os.walk(input_path):
        if out_abs and os.path.abspath(root).startswith(out_abs):
            continue
        for name in names:
            if _is_candidate_excel(name):
                files.append(os.path.join(root, name))
    return files


def _detect_all_sheets(file_path, config):
    """返回 [(sheet_name, header_row_1based, [columns]), ...]，覆盖文件里每一个 sheet。"""
    work_path, is_tmp = normalize_to_xlsx(file_path)
    try:
        df_dict = pd.read_excel(work_path, sheet_name=None, header=None, engine="openpyxl", nrows=20)
        out = []
        for sheet_name, df in df_dict.items():
            rows = df.values.tolist()
            h = resolve_header_row(rows, config)
            cols = [c for c in _header_values(df, h) if c] if (h != -1 and h <= len(df)) else []
            out.append((sheet_name, h, cols))
        return out
    finally:
        if is_tmp and os.path.exists(work_path):
            os.remove(work_path)


def _enum_values_all_sheets(file_path, column, config):
    """跨所有 sheet 枚举某列取值。返回 (values_sorted, sheets_with_col)。"""
    alias = config.get("value_alias_map", {})
    skip = config.get("skip_values", [])
    work_path, is_tmp = normalize_to_xlsx(file_path)
    try:
        df_dict = pd.read_excel(work_path, sheet_name=None, header=None, engine="openpyxl")
        values = set()
        sheets_with = []
        for sn, df in df_dict.items():
            head_rows = df.values[:20].tolist()
            h = resolve_header_row(head_rows, config)
            if h == -1 or h > len(df):
                continue
            ci = _find_col_index(df, h, column)
            if ci is None:
                continue
            sheets_with.append(sn)
            for v in df.iloc[h:, ci]:
                nv = apply_alias(v, alias)
                if not is_skip_value(nv, skip):
                    values.add(nv)
        return sorted(values), sheets_with
    finally:
        if is_tmp and os.path.exists(work_path):
            os.remove(work_path)


def main():
    ap = argparse.ArgumentParser(description="看一个 Excel 文件/目录：表头行、字段名、某字段取值")
    ap.add_argument("--input", required=True, help="Excel 文件或目录路径")
    ap.add_argument("--column", default=None, help="要枚举取值的字段名（不填则只看列名）")
    ap.add_argument("--output", default=None, help="计划用的输出目录（仅用于目录扫描时跳过其子树，可不填）")
    ap.add_argument("--header-mode", choices=["auto", "row", "keyword"], default="auto")
    ap.add_argument("--header-row", type=int, default=1, help="header-mode=row 时的 1 基行号")
    ap.add_argument("--grid-keys", default="", help="header-mode=keyword 时的第一类关键词，逗号分隔")
    ap.add_argument("--id-keys", default="", help="header-mode=keyword 时的第二类关键词，逗号分隔")
    args = ap.parse_args()

    check_deps()

    if not os.path.exists(args.input):
        emit_error(f"找不到输入路径：{args.input}")

    config = build_config(
        header_mode=args.header_mode,
        header_row=args.header_row,
        grid_keys=[s for s in args.grid_keys.split(",") if s.strip()],
        id_keys=[s for s in args.id_keys.split(",") if s.strip()],
    )

    is_dir = os.path.isdir(args.input)
    files = _iter_input_files(args.input, args.output)

    result = {
        "ok": True,
        "input": os.path.abspath(args.input),
        "is_dir": is_dir,
        "excel_count": len(files),
    }

    if not files:
        result["warning"] = "没找到任何 .xlsx / .xls 文件"
        emit(result)
        return

    template = files[0]
    result["template"] = os.path.basename(template)

    try:
        sheets = _detect_all_sheets(template, config)
    except Exception as e:
        emit_error(f"读取「{os.path.basename(template)}」失败：{e}", excel_count=len(files))
        return

    # 向后兼容：保留首个可识别 sheet 的 sheet/header_row/columns
    first_ok = next((s for s in sheets if s[1] != -1), None)
    if first_ok is None:
        result["warning"] = (
            "自动没能识别出任何 sheet 的表头行。若表头不在常见位置，改用 "
            "--header-mode row --header-row N 指定，或用 --header-mode keyword "
            "--grid-keys ... --id-keys ... 按关键词定位。"
        )
        result["sheets"] = [{"name": s[0], "header_row": s[1], "columns": s[2]} for s in sheets]
        emit(result)
        return

    result["sheet"] = first_ok[0]
    result["header_row"] = first_ok[1]
    result["columns"] = first_ok[2]
    # v2.6.1：完整 sheet 清单（多 sheet 文件的关键信息）
    result["sheets"] = [
        {"name": s[0], "header_row": s[1], "columns": s[2]} for s in sheets
    ]

    if args.column:
        # 跨所有 sheet 搜索该列，不再只盯第一个 sheet
        try:
            values, sheets_with = _enum_values_all_sheets(template, args.column, config)
        except Exception as e:
            emit_error(f"枚举「{args.column}」取值失败：{e}", excel_count=len(files))
            return
        if not sheets_with:
            all_cols = {s[0]: s[2] for s in sheets if s[1] != -1}
            result["warning"] = (
                f"字段「{args.column}」在任何一个 sheet 里都没找到。各 sheet 字段为：{all_cols}"
            )
            emit(result)
            return
        result["values"] = values
        result["column_sheets"] = sheets_with
        if len(sheets_with) < sum(1 for s in sheets if s[1] != -1):
            result["note"] = f"字段「{args.column}」仅出现在以下 sheet：{sheets_with}"

    emit(result)


if __name__ == "__main__":
    main()
