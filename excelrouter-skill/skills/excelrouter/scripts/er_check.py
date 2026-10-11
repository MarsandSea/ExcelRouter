#!/usr/bin/env python
"""
excelrouter · 查收：表发下去让大家填，收回来几秒查出谁没交、谁交了空表、谁哪一列漏了哪几行，
写好一段可以直接发群的通报和逐人催办的话。不合并、不改原表。

Copyright (c) 2026 AbeLin
MIT License

和桌面版「Excel 汇总 → 先查收，再汇总」同一套内核（core/collect.py）：
  · 谁该交：--expect 名单 → 下发记录.json（用 ExcelRouter 拆分下发时写的）→ 文件夹里的「应交名单」表
  · 文件是谁的：表里「填报说明」写的填报对象 → 文件名 / 所在文件夹名（长名字优先，「城区二」不会被认成「城区」）
  · 哪些列该填：--required → 下发时的填写列 → 大多数表都填满了的列（备注、说明这类除外）
  · 同一个人交了几次（「副本」「(2)」）只查最新那份；有下发记录时还能查出删了行、改了不该改的格子
  · 每查一次记一轮，下次通报里写「比上次新交」了谁

用法举例：
  python er_check.py --input 九月收表/ --expect "城区、东城、西城、南城、北城"
  python er_check.py --input 九月收表/ --required "走访日期、走访结果" --due "10月15日 18:00"
  # 一张大家各填各行的总表（如共享表格导出的），按「网格」列分人查
  python er_check.py --input 走访登记总表.xlsx --by 网格 --expect "网格一、网格二"

查完要汇总交了的表：python er_merge.py --input 同一个文件夹 --checked（--expect 等参数照抄）。

输出：stdout 一行 JSON。notice 是通报全文（可直接复制发群），message 是给用户的完整说明。
"""

import argparse
import json
import os
import sys
import time

T_START = time.time()

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import (  # noqa: E402
    add_recipe_args, apply_recipe_args, bootstrap, emit, emit_error, ensure_deps, eprint, open_folder, save_recipe_from,
)

bootstrap()
ensure_deps(lite=True)  # 查收只要 openpyxl + xlrd；必须排在 core import 之前

from core import collect as em  # noqa: E402
from _gather import gather  # noqa: E402

BUDGET = 100   # 宿主 120 秒掐断命令，留出余量


def main():
    ap = argparse.ArgumentParser(
        description="查收：逐人查出未交 / 空表 / 漏填（到列到行）/ 和下发时不一样，写好可直接发群的通报",
        formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    ap.add_argument("--input", required=True, help="收回来的表所在的文件夹；或一张各填各行的总表（要配 --by）")
    ap.add_argument("--expect", default="", help="应交名单，顿号 / 逗号 / 空格分隔。不填就用下发记录或文件夹里的「应交名单」表")
    ap.add_argument("--required", default="", help="必须填的列，顿号 / 逗号分隔。不填自动判断")
    ap.add_argument("--by", default="", help="输入是一张总表时，按哪一列分人（如 网格、姓名、单位）")
    ap.add_argument("--title", default="", help="通报标题，不填用文件夹名")
    ap.add_argument("--due", default="", help="截止时间，写进通报和催办，如「10月15日 18:00」")
    ap.add_argument("--record", default="", help="下发记录.json 的路径（不填自动在文件夹和上一层找）")
    ap.add_argument("--header-row", type=int, default=0, help="表头在第几行（自动识别不对时才填）")
    ap.add_argument("--output", default="", help="通报和进度表放哪个文件夹。不填 = 输入旁边的「汇总结果」")
    ap.add_argument("--open", action="store_true", help="查完打开结果文件夹")
    ap.add_argument("--budget", type=float, default=BUDGET, help=argparse.SUPPRESS)   # 测试用
    add_recipe_args(ap)
    args = ap.parse_args()
    recipe, recipe_notes, _extra = apply_recipe_args(ap, args, "collect", {
        "collect_expect": ("expect", lambda v, s: v),
        "collect_required": ("required", lambda v, s: v),
        "collect_by": ("by", lambda v, s: v),
        "collect_title": ("title", lambda v, s: v),
        "header_row": ("header_row", lambda v, s: v if s.get("header_mode") == "row" else 0),
    })

    src = os.path.abspath(args.input)
    if not os.path.exists(src):
        emit_error(f"找不到「{src}」，请确认文件夹或文件的完整路径。")
    is_dir = os.path.isdir(src)
    if not is_dir and not args.by:
        emit_error("选的是一个文件。如果是大家各填各行的总表，请告诉我按哪一列分人（比如「网格」「姓名」），"
                   "带上 --by 再查；如果是想把它的几个 sheet 合成一张，用 er_merge.py。", need="by")
    record = os.path.abspath(args.record) if args.record else ""
    if record and not os.path.isfile(record):
        emit_error(f"找不到下发记录「{record}」。它是拆分下发时写在结果文件夹里的「{em.RECORD_NAME}」；"
                   f"没有就去掉 --record。")
    out_dir = os.path.abspath(args.output) if args.output else os.path.join(
        src if is_dir else os.path.dirname(src), em.RESULT_DIR)

    g = gather(src, out_dir, em._base_stem)
    if not g["files"]:
        emit_error(("文件夹里还没有交回来的表（.xlsx / .xls）。" if not g["problems"] else "这些文件都读不了：")
                   + "".join(f"\n· {p}" for p in g["problems"]), problems=g["problems"])

    cfg = {"header_mode": "row", "header_row": args.header_row} if args.header_row else {
        "header_mode": "auto", "header_row": 1}
    cfg.update(input_path=src, output_path=out_dir, collect_expect=args.expect, collect_required=args.required,
               collect_by=args.by, collect_title=args.title, collect_due=args.due, collect_record=record,
               grid_keys=[], id_keys=[])
    got = {}

    def log(msg):
        s = str(msg)
        if s.startswith("[SUMMARY] "):
            try:
                got.update(json.loads(s[len("[SUMMARY] "):]))
            except ValueError:
                pass
        elif s:
            eprint(s)

    try:
        em.check_collect(cfg, log_fn=log, stop_flag=lambda: time.time() > T_START + min(args.budget, BUDGET))
    except Exception as e:
        emit_error(f"查收出错：{e}")
    if got.get("stopped") or not got.get("notice"):
        if got.get("stopped"):
            emit_error(f"表格太多，{int(min(args.budget, BUDGET))} 秒内没查完。请把表格分到几个文件夹里分批查收。")
        emit_error(f"通报没写成（结果文件夹 {out_dir} 可能没有写入权限）。")
    with open(got["notice"], encoding="utf-8-sig") as f:
        notice = f.read().rstrip("\n")
    opened = open_folder(out_dir) if args.open else False

    tail = []
    # 新内核会把读不了的文件记在交表人名下、原因写进通报；通报里已经点名的不再重复
    problems = [p for (_path, rel), p in zip(g["bad"], g["problems"]) if f"「{rel}」" not in notice]
    if problems:
        tail.append(f"这些文件读不了、没查（{len(problems)}）：")
        tail += [f"  {p}" for p in problems]
    if got.get("expected") is None:
        tail.append("没有应交名单，查不出谁没交：告诉我应交的名字（或放一张「应交名单」表进文件夹）再查一次。")
    tail.append(f"通报已存为「{em.NOTICE_NAME}」，逐人状态在「{em.PROGRESS_NAME}」（未交标红、有问题标橙、合格标绿），"
                f"都在：{out_dir}" + ("（已为你打开）" if opened else ""))
    files = [p for p in (got.get("progress"), got.get("notice")) if p]
    saved = save_recipe_from(args, dict(cfg, collect_check=True), "collect", input_name=src)
    if saved:
        tail.append(f"这次的设置存成了方案：{saved}（下次加 --recipe 这个文件就行）")
    people = []
    if got.get("result"):
        try:
            with open(got["result"], encoding="utf-8") as f:
                people = json.load(f).get("people", [])
        except (OSError, ValueError):
            pass
    emit({"ok": True, "notice": notice, "message": notice + "\n\n" + "\n".join(tail),
          "counts": {k: got.get(k) for k in ("expected", "submitted", "ok", "missing", "empty", "partial",
                                             "invalid", "changed", "bad", "extra")},
          "people": people, "result_file": got.get("result") or "",
          "recipe": recipe["name"] if recipe else "", "recipe_notes": recipe_notes, "recipe_saved": saved,
          "files": files, "output_dir": out_dir, "warnings": g["warnings"], "opened": opened,
          "next": "把 message 原样告诉用户（通报可以直接复制发群，逐人催办可单独发给本人）。"
                  "用户追问某个人还差什么、只要有某类问题的人时，直接按 people（每人的 status 和 issues 编码）回答，"
                  "不用重新运行。用户说补齐了就再运行一次本命令（会写出新交、改好了、还差）；要汇总交了的表，"
                  "运行 er_merge.py --input 同一个路径 --checked，并带上同样的 --expect / --required / --by / --record。"})


if __name__ == "__main__":
    main()
