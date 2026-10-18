#!/usr/bin/env python
"""
excelrouter · 数据脱敏：表格里的手机号、身份证号、银行卡号、姓名、地址等，按用途打码 / 换代号 / 泛化 / 清空，
处理完把结果重新扫一遍（复检），还有完整号码的表不写出来。原表不改，结果放在旁边的「脱敏结果」。

Copyright (c) 2026 AbeLin
MIT License

和桌面版「数据脱敏」同一套内核（core/desens.py）。用途（--scene）：
  · 发群通报（默认）：认出来的全部打码 —— 138****5678、张*、证件号留前 6 后 4、地址留到镇街
  · 内部下发：只遮证件号和银行卡号，手机号、姓名、地址照发（外呼要完整号码）
  · 交给AI分析：换成一致的代号（同一个人同一个代号，还能统计、关联），地址到区县、日期到年
  · 对外报送：个人信息整列清空，编码换成代号
某几列单独处理用 --set "列名:处理方式、…"（打码 / 换代号 / 泛化 / 清空 / 照发）；只按点名的列处理用 --scene 自定义。
备注这类文字里夹带的手机号、证件号、银行卡号也就地打码（--no-notes 关掉）。

用法举例：
  # 先看会怎么处理（只读，不写盘）
  python er_desens.py --input 客户清单.xlsx --scene 发群通报 --dry-run
  # 处理一个文件夹，要交给 AI 分析
  python er_desens.py --input 九月清单/ --scene 交给AI分析
  # 内部下发，备注清空、客户姓名照发
  python er_desens.py --input 明细.xlsx --scene 内部下发 --set "备注:清空、客户姓名:照发"

拆分时顺带打码：er_split.py ... --desens 内部下发；汇总完另存脱敏版：er_merge.py ... --desens 对外报送。

输出：stdout 一行 JSON。message 是给用户的完整说明；columns 只有列名、类型和处理方式，不含任何格子的值。
"""

import argparse
import json
import os
import sys
import time

T_START = time.time()

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import (  # noqa: E402
    DESENS_SCENES, METHOD_TEXT, add_recipe_args, apply_recipe_args, bootstrap, emit, emit_error, ensure_deps, eprint,
    open_folder, parse_desens_set, save_recipe_from,
)

bootstrap()
ensure_deps(lite=True)  # 脱敏只要 openpyxl + xlrd；必须排在 core import 之前

try:
    from core import desens as D  # noqa: E402
except ImportError:      # 技能目录里的内核还没同步到带数据脱敏的版本
    D = None

BUDGET = 100   # 宿主 120 秒掐断命令，留出余量
SCENES = DESENS_SCENES + ("自定义",)


def main():
    ap = argparse.ArgumentParser(
        description="数据脱敏：按用途给表格里的个人信息打码 / 换代号 / 泛化 / 清空，处理完复检",
        formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    ap.add_argument("--input", required=True, help="一个表格，或装表格的文件夹")
    ap.add_argument("--scene", default="", choices=("",) + SCENES, help="用途，默认 发群通报")
    ap.add_argument("--set", default="", help="逐列调整：「列名:处理方式、…」")
    ap.add_argument("--no-notes", action="store_true", help="文字里夹带的号码不处理")
    ap.add_argument("--keep-hidden", action="store_true", help="隐藏的 sheet 和列也照样留着（默认不带、清空）")
    ap.add_argument("--header-row", type=int, default=0, help="表头在第几行（自动识别不对时才填）")
    ap.add_argument("--dry-run", action="store_true", help="只识别、说明会怎么处理，不写任何文件")
    ap.add_argument("--open", action="store_true", help="做完打开结果文件夹")
    ap.add_argument("--budget", type=float, default=BUDGET, help=argparse.SUPPRESS)   # 测试用
    add_recipe_args(ap)
    args = ap.parse_args()
    recipe, recipe_notes, _extra = apply_recipe_args(ap, args, "desens", {
        "desens_scene": ("scene", lambda v, s: v),
        "desens_columns": ("set", lambda v, s: v),
        "desens_notes": ("no_notes", lambda v, s: not v),
        "keep_hidden": ("keep_hidden", lambda v, s: v),
        "header_row": ("header_row", lambda v, s: v if s.get("header_mode") == "row" else 0),
    })
    if D is None:
        emit_error("技能里的 ExcelRouter 内核还没更新到带数据脱敏的版本（v3.1 起才有），请更新本技能后再用。")

    src = os.path.abspath(args.input)
    if not os.path.exists(src):
        emit_error(f"找不到「{src}」，请确认文件或文件夹的完整路径。")
    if os.path.isfile(src) and not src.lower().endswith((".xlsx", ".xlsm", ".xls")):
        emit_error(f"「{os.path.basename(src)}」不是 Excel 表格（.xlsx / .xlsm / .xls）。PDF、Word 里的内容不做脱敏。")
    columns = parse_desens_set(args.set)
    scene = args.scene or "发群通报"
    if scene == "自定义" and not columns:
        emit_error("选了「自定义」，但没有说哪列怎么处理：用 --set \"列名:处理方式\" 指定，或换一个用途。")
    setting = {"scene": scene, "columns": columns, "notes": not args.no_notes}
    cfg = {"header_mode": "row", "header_row": args.header_row} if args.header_row else {
        "header_mode": "auto", "header_row": 1}
    cfg.update(input_path=src, desens=setting, keep_hidden=args.keep_hidden, grid_keys=[], id_keys=[])
    tasks = D._tasks(src)
    if not tasks:
        emit_error(f"「{src}」里没有找到表格（.xlsx / .xlsm / .xls）。")

    if args.dry_run:
        found = {}
        for fp, rel in tasks[:20]:
            try:
                info = D.inspect(fp, setting, cfg)
            except Exception as e:
                eprint(f"读不了「{rel}」：{e}")
                continue
            for s in info:
                if s["hidden"] and not args.keep_hidden:
                    continue
                for c in s["columns"]:
                    found.setdefault((c["column"], c["kind"], c["method"]), 0)
                    found[(c["column"], c["kind"], c["method"])] += 1
        cols = [{"column": c, "kind": k, "method": m, "sheets": n} for (c, k, m), n in found.items()]
        lines = [f"用途：{scene}——{D.SCENE_TEXT.get(scene, '')}。共 {len(tasks)} 个表"
                 + ("（只看了前 20 个）" if len(tasks) > 20 else "") + "。"]
        todo = [c for c in cols if c["method"] != "keep"]
        if todo:
            lines.append("会处理：" + "；".join(f"「{c['column']}」{c['kind']} → {METHOD_TEXT[c['method']]}" for c in todo))
        else:
            lines.append("按这个用途没有认出要处理的列。")
        kept = [c for c in cols if c["method"] == "keep"]
        if kept:
            lines.append("认出了但照发：" + "、".join(f"「{c['column']}」{c['kind']}" for c in kept))
        lines.append("识别靠列名和号码校验，不保证找全：没认出的列照发，漏了的用 --set 指定。")
        emit({"ok": True, "dry_run": True, "scene": scene, "columns": cols, "files": len(tasks),
              "message": "\n".join(lines),
              "next": "把 message 告诉用户，问按这样处理吗；用户同意就去掉 --dry-run 再运行一次（其他参数照抄）。"})
        return

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
        out = D.run_desens(cfg, log_fn=log, stop_flag=lambda: time.time() > T_START + min(args.budget, BUDGET))
    except Exception as e:
        emit_error(f"脱敏出错：{e}。原表没有被改动。")
    report = got.get("report") or ""
    written = sorted(os.path.join(r, n) for r, _d, ns in os.walk(out) for n in ns if n.endswith("_脱敏.xlsx"))
    opened = open_folder(out) if args.open else False
    lines = [f"脱敏好了（{scene}）：{got.get('files', 0)} 个表写进了结果，处理 {got.get('columns', 0)} 列、"
             f"{got.get('cells', 0)} 个格子" + (f"，文字里夹带的号码 {got['embedded']} 处" if got.get("embedded") else "") + "。",
             f"结果在：{out}" + ("（已为你打开）" if opened else ""),
             f"「{D.REPORT_NAME}」逐列写了识别为什么、怎么处理、处理了几格，示例是虚构的，不含表里的真实内容。"]
    warnings = []
    if got.get("stopped"):
        warnings.append(f"表格太多，{int(min(args.budget, BUDGET))} 秒内没做完：已处理的在结果里，剩下的请分到几个文件夹再处理。")
    if got.get("blocked"):
        warnings.append(f"{got['blocked']} 个表没通过复检（还有完整的号码），没有写出来；结果里的「{D.FAIL_NOTE}」写了是哪列，"
                        "用 --set \"列名:打码\" 指定后再运行一次。")
    if got.get("failed_files"):
        warnings.append(f"{got['failed_files']} 个表没处理成（加了打开密码或正被 Excel 占用？）。")
    lines += [f"请注意：{w}" for w in warnings]
    lines.append("复检：" + ("处理完重新扫了一遍结果，处理过的类型里没有发现完整的号码（通过）。" if got.get("ok") else "没通过的表已拦下，没有写出来。"))
    lines.append("批注、图片、图表、文档属性没有带进结果；打码不可逆，原表没有改动。")
    saved = save_recipe_from(args, {"desens_scene": scene, "desens_columns": args.set, "desens_notes": not args.no_notes,
                                    "keep_hidden": args.keep_hidden, **{k: cfg[k] for k in ("header_mode", "header_row")}},
                             "desens", input_name=src)
    if saved:
        lines.append(f"这次的设置存成了方案：{saved}（下次加 --recipe 这个文件就行）")
    emit({"ok": bool(got.get("ok")) and not got.get("failed_files") and not got.get("stopped"),
          "output_dir": out, "report": report, "files": written[:10], "files_total": len(written),
          "summary": {k: got.get(k) for k in ("scene", "files", "blocked", "failed_files", "columns", "cells",
                                             "embedded", "residual", "ok", "stopped")},
          "message": "\n".join(lines), "warnings": warnings, "opened": opened,
          "recipe": recipe["name"] if recipe else "", "recipe_notes": recipe_notes, "recipe_saved": saved,
          "next": "把 message 原样告诉用户（「请注意」各条都要说到），请用户打开结果文件夹查看。"})


if __name__ == "__main__":
    main()
