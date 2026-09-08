#!/usr/bin/env python
"""
excelrouter · 按字段拆分 Excel。

Copyright (c) 2026 AbeLin
MIT License

跑之前建议先用 er_inspect.py 看一眼字段名和取值，不要凭猜测直接填 --by。

用法举例：
  # 单文件，按「部门」拆
  python er_split.py --input 明细.xlsx --output 拆分结果 --by 部门

  # 一整个目录，只拆「东区/西区」两个取值，并附加按「姓名」拆到人
  python er_split.py --input 一批表/ --output 拆分结果 --by 区域 \
      --values 东区,西区 --to-person 姓名

  # 单个文件也要按人拆 → 照样可以，包装层会在一级拆分后补做二级拆分
  python er_split.py --input 明细.xlsx --output 拆分结果 --by 网格 --to-person 姓名

  # 多个 sheet 且各 sheet 人字段不同（sheet名:人字段，逗号分隔；不写 sheet 名用 * 兜底）
  python er_split.py --input 一批表/ --output 拆分结果 --by 网格 \
      --to-person "结对子维表:1看管人,渠道维表:看管 渠道管理员"

  # 懒得查人字段叫什么，让工具每个 sheet 自己挑（按 接收人/看管人/管理员… 关键词）
  python er_split.py --input 一批表/ --output 拆分结果 --by 网格 --to-person auto

  # 跨文件合并同一取值到一个文件（默认不合并，按原表各自拆分）
  python er_split.py --input 一批表/ --output 拆分结果 --by 区域 --merge

  # 输出直接落在 --output，不再套一层 {MMDDHHMM}结果 时间戳子目录
  python er_split.py --input 明细.xlsx --output 拆分结果 --by 部门 --no-timestamp

  # 只看会怎么拆、不真正写盘
  python er_split.py --input 明细.xlsx --output 拆分结果 --by 部门 --dry-run

输出：stdout 一行 JSON，成功时含 output_path（真实产出目录）。默认它会比 --output
多一层带时间戳的子目录，务必以 JSON 里的 output_path 为准、不要自己拼路径；
加 --no-timestamp 则 output_path 就等于你传的 --output。
"""

import argparse
import json
import os
import signal
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import bootstrap, build_config, check_deps, emit, emit_error, eprint  # noqa: E402

bootstrap()

from core.splitter import run_split  # noqa: E402
from _person_split import (  # noqa: E402  网格→人 二级拆分（scripts 包装层补强）
    flatten_output,
    parse_person_spec,
    person_split,
)


def _split_list(s):
    return [x.strip() for x in s.split(",") if x.strip()] if s else []


def main():
    ap = argparse.ArgumentParser(
        description="按字段拆分 Excel（保留原格式，支持跨文件合并、到人二级拆分）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    ap.add_argument("--input", required=True, help="源 Excel 文件或目录")
    ap.add_argument("--output", required=True, help="输出根目录（实际产出会在其下多一层时间戳子目录）")
    ap.add_argument("--by", dest="split_column", default=None, help="拆分字段（列名）")
    ap.add_argument("--values", default="", help="只拆这些取值，逗号分隔；留空 = 自动枚举全部取值")
    ap.add_argument("--to-person", dest="person_column", default=None, metavar="人字段",
                     help="附加产出「到人」结果。三种写法：① 列名（所有 sheet 共用）；"
                          "② auto（每个 sheet 自动挑人字段）；"
                          "③ \"sheet名:人字段,sheet名:人字段\"（各 sheet 人字段不同）。"
                          "单文件输入也支持，由 scripts 包装层补做二级拆分")
    ap.add_argument("--person-filter", default="", help="到人只处理文件名命中这些关键词的表，逗号分隔；留空=全部")
    ap.add_argument("--no-zip", action="store_true", help="批量拆分后不打包 ZIP")
    ap.add_argument("--no-timestamp", action="store_true",
                     help="输出直接写到 --output，不再套一层 {MMDDHHMM}结果 时间戳子目录"
                          "（路径可预测，方便脚本接力/直接告诉用户文件在哪）")
    ap.add_argument("--merge", action="store_true",
                     help="同一取值跨源文件合并到一个输出文件（默认关闭：按原表各自拆分，只在同一取值下打包在一起）")
    ap.add_argument("--fast", action="store_true", help="快速模式：数据行只写值不保留格式，速度换取不保留样式")
    ap.add_argument("--keep-formulas", action="store_true",
                     help="尽量保留同行公式（跨行/汇总/跨表公式仍会落成当前缓存数值，不会发送算错的公式）")
    ap.add_argument("--fuzzy-match", action="store_true", help="--values 按包含匹配而不是精确匹配")
    ap.add_argument("--header-mode", choices=["auto", "row", "keyword"], default="auto")
    ap.add_argument("--header-row", type=int, default=1, help="header-mode=row 时的 1 基行号")
    ap.add_argument("--grid-keys", default="", help="header-mode=keyword 时第一类关键词，逗号分隔")
    ap.add_argument("--id-keys", default="", help="header-mode=keyword 时第二类关键词，逗号分隔")
    ap.add_argument("--skip-values", default=None,
                     help="拆分列中要忽略的取值，逗号分隔；不填则用默认（合计/小计/总计/平均/空）")
    ap.add_argument("--alias-json", default="", help='取值归并，如 \'{"东区":["东片区","东部"]}\'')
    ap.add_argument("--config", default=None, help="直接传一份完整 config JSON（专家逃生口，字段同桌面版 user_config.json）")
    ap.add_argument("--dry-run", action="store_true", help="只解析配置并打印，不实际拆分、不写盘")
    args = ap.parse_args()

    check_deps()

    person_map = None  # --config 专家模式下保持 None：包装层不介入，完全按用户配置走

    if args.config:
        try:
            with open(args.config, encoding="utf-8") as f:
                user_cfg = json.load(f)
        except (OSError, json.JSONDecodeError) as e:
            emit_error(f"读取 --config 失败：{e}")
            return
        config = build_config(**user_cfg)
        config["input_path"] = os.path.abspath(args.input)
        config["output_path"] = os.path.abspath(args.output)
    else:
        if not args.split_column:
            emit_error("缺少 --by（拆分字段）。先用 er_inspect.py 看看有哪些字段可选，"
                       "或者用 --config 传一份完整配置。")
            return
        alias_map = {}
        if args.alias_json:
            try:
                alias_map = json.loads(args.alias_json)
            except json.JSONDecodeError as e:
                emit_error(f"--alias-json 不是合法 JSON：{e}")
                return
        # --to-person 三种写法：列名 / auto / "sheet:列,sheet:列"
        person_map = parse_person_spec(args.person_column) if args.person_column else None
        if args.person_column and person_map is None:
            emit_error("非法的 --to-person 语法；应为 列名 / auto / 或 sheet:列,sheet:列（逗号分隔）")
            return
        # 决定「到人」由谁做：内核只认「单一全局人字段 + 目录输入」，
        # per-sheet / auto / 单文件 这三种内核做不到（单文件还会静默只拆一级），
        # 交给 scripts 包装层补做；其余场景仍交给内核，内核的实现更完整（含按人打包 ZIP）。
        _wrapper_only = bool(person_map) and (
            any(k != "*" for k in person_map) or person_map.get("*") == "auto")
        vendor_person_col = "" if _wrapper_only else (
            (person_map or {}).get("*", "") if person_map else "")
        config = build_config(
            input_path=os.path.abspath(args.input),
            output_path=os.path.abspath(args.output),
            split_column=args.split_column,
            selected_values=_split_list(args.values),
            person_column=vendor_person_col,
            to_person=bool(person_map),
            person_file_filter=_split_list(args.person_filter),
            make_zip=not args.no_zip,
            merge_across_files=args.merge,
            preserve_format=not args.fast,
            keep_formulas=args.keep_formulas,
            no_timestamp=args.no_timestamp,
            exact_match=(False if args.fuzzy_match else None),
            header_mode=args.header_mode,
            header_row=args.header_row,
            grid_keys=_split_list(args.grid_keys),
            id_keys=_split_list(args.id_keys),
            skip_values=[s.strip() for s in args.skip_values.split(",")] if args.skip_values is not None else None,
            value_alias_map=alias_map or None,
        )

    if not os.path.exists(config["input_path"]):
        emit_error(f"找不到输入路径：{config['input_path']}")
        return
    if not config.get("split_column"):
        emit_error("拆分字段（split_column）为空，无法拆分")
        return

    if args.dry_run:
        eprint("dry-run：不写盘，仅打印解析后的配置")
        emit({"ok": True, "dry_run": True, "config": config})
        return

    state = {"stopped": False}

    def _on_sigint(_signum, _frame):
        state["stopped"] = True
        eprint("⛔ 收到中断信号，准备在当前文件处理完后停止…")

    signal.signal(signal.SIGINT, _on_sigint)

    last_pct = {"v": -1}

    def progress_fn(v):
        pct = int(v * 100)
        if pct != last_pct["v"]:
            last_pct["v"] = pct
            eprint(f"[进度 {pct}%]")

    try:
        output_path = run_split(config, log_fn=eprint, progress_fn=progress_fn,
                                 stop_flag=lambda: state["stopped"])
    except Exception as e:
        emit_error(f"拆分失败：{e}")
        return

    # 包装层补做「到人」：必须作用在内核刚产出的真实目录（默认带时间戳）上，
    # 若先扁平化会找不到刚拆好的网格文件——所以这一步要排在 --no-timestamp 之前。
    single_file = os.path.isfile(config["input_path"])
    need_wrapper_person = bool(person_map) and (
        single_file or any(k != "*" for k in person_map) or person_map.get("*") == "auto")
    person_count = 0
    if need_wrapper_person:
        person_count = person_split(output_path, person_map, log_fn=eprint)

    # --no-timestamp：把内容提升到 --output 根、删掉时间戳子目录。
    # 上游内核无条件建时间戳目录、且没有这个开关，所以扁平化只能在包装层做。
    if args.no_timestamp:
        output_path = flatten_output(output_path, config["output_path"], log_fn=eprint)

    emit({"ok": True, "output_path": output_path,
          "person_split_by": ("wrapper" if need_wrapper_person
                              else ("vendor" if person_map else None)),
          "person_files": person_count if need_wrapper_person else None,
          "stopped": state["stopped"]})


if __name__ == "__main__":
    main()
