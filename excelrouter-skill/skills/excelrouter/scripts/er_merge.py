#!/usr/bin/env python
"""
excelrouter · Excel 汇总：把一个文件夹里同一套模板的表合成一张（或把一个文件的多个 sheet 合成一张）。

Copyright (c) 2026 AbeLin
MIT License

和桌面版「Excel 汇总」同一套内核（core/collect.py）：
  · 表头以第一个表为准（值、格式、合并单元格、列宽原样），其余表**按列名对齐**——
    列顺序被调过、多一列少一列都能合；能对上的列不到一半的表不硬合，单独点名
  · 每行最后加一列「来源」写来自哪个文件（一个文件合 sheet 时写 sheet 名）
  · 各表自己的合计 / 小计行不合进来（合进来再求和会重复计算）；内容完全相同的文件只合一份；
    隐藏的 sheet、「填报说明」页不合
  · 最后一个 sheet「汇总说明」逐个列出合了几行、缺哪些列、多哪些列、没合进去的原因
  · 结果里没有公式（按算好的值保存），打开不会执行任何东西

用法举例：
  # 文件夹里的表合成一张（结果默认在文件夹里的「汇总结果」）
  python er_merge.py --input 九月报表/ --open

  # 一个文件的 12 个月份 sheet 合成一张
  python er_merge.py --input 全年明细.xlsx --open

  # 先查收过（er_check.py），只合交了的表，「来源」写交表人
  python er_merge.py --input 九月收表/ --checked --expect "城区、东城、西城" --open

表多时会分几次做完：宿主会掐断跑太久的命令（WorkBuddy 120 秒），每次约 100 秒就存好进度返回
done:false，**原样再运行一次同一条命令**就接着合，直到 done:true。

输出：stdout 一行 JSON。成功时 output 是汇总表路径，message 是可以直接转述的说明。
"""

import argparse
import hashlib
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
ensure_deps(lite=True)  # 汇总只要 openpyxl + xlrd；必须排在 core import 之前

from core import collect as em  # noqa: E402
from _gather import gather, order_key  # noqa: E402

BUDGET = 100          # 从进程启动算起，做到这么多秒就存好进度返回（宿主 120 秒掐断）
SCAN_SECONDS = 60     # 合并前读表头的时间上限
STATE_SUFFIX = "(未完成).进度.json"
PART_SUFFIX = "(未完成).xlsx"


def main():
    ap = argparse.ArgumentParser(
        description="Excel 汇总：同一套模板的多个表（或一个文件的多个 sheet）合成一张，按列名对齐、保留表头格式",
        formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    ap.add_argument("--input", required=True, help="要汇总的文件夹，或一个有多个 sheet 的文件")
    ap.add_argument("--output", default="", help="结果放哪个文件夹。不填 = 输入旁边的「汇总结果」（与桌面版一致）")
    ap.add_argument("--title", default="", help="结果文件名里的标题，不填用文件夹名 / 文件名")
    ap.add_argument("--header-row", type=int, default=0, help="表头在第几行（自动识别不对时才填）")
    ap.add_argument("--open", action="store_true", help="做完在文件管理器里打开结果文件夹。AI 在用户本机运行时应总是加上")
    ap.add_argument("--checked", action="store_true",
                    help="按查收结果只合交了的表（没交、空表不合），「来源」写交表人。配合下面几个参数，和 er_check.py 用同样的值")
    ap.add_argument("--expect", default="", help="--checked 时：应交名单，顿号 / 逗号分隔")
    ap.add_argument("--required", default="", help="--checked 时：必须填的列，空着的格子在结果里标浅红")
    ap.add_argument("--record", default="", help="--checked 时：下发记录.json 的路径（不填自动在文件夹和上一层找）")
    ap.add_argument("--by", default="", help="--checked 且输入是一张各填各行的总表时：按哪一列分人")
    ap.add_argument("--include-old-results", action="store_true",
                    help="文件夹里混着以前的拆分结果目录时也照样合（默认拦下，免得行数翻倍）")
    ap.add_argument("--restart", action="store_true", help="丢掉上次没做完的进度，从头合")
    ap.add_argument("--budget", type=float, default=BUDGET, help=argparse.SUPPRESS)   # 测试用
    add_recipe_args(ap)
    args = ap.parse_args()
    recipe, recipe_notes, _extra = apply_recipe_args(ap, args, "collect", {
        "collect_check": ("checked", lambda v, s: v),
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
    out_dir = os.path.abspath(args.output) if args.output else os.path.join(
        src if is_dir else os.path.dirname(src), em.RESULT_DIR)
    cfg = _config(args, src, out_dir)

    g = gather(src, out_dir, em._base_stem)
    if g["stale"] and not args.include_old_results:
        emit_error(f"文件夹里有 {len(g['stale'])} 个以前的拆分结果目录（如 {'、'.join(g['stale'][:3])}），"
                   f"和原表一起合会让同样的行出现两次。先问用户：把它们移走，"
                   f"还是确实要一起合（那就加 --include-old-results）。", stale_dirs=g["stale"])
    # 普通汇总不认名单（与桌面版一致：名字像名单的表也照样合）；表头和别的表对不上的名单表在 _plan 里剔掉
    files = g["files"] if args.checked else sorted(
        g["files"] + [(p, os.path.relpath(p, src)) for p in g["lists"]], key=lambda f: order_key(f[1]))
    if not is_dir and files and not args.checked:
        files = [files[0]]
    if not files:
        emit_error(("没找到能汇总的表格（.xlsx / .xlsm / .xls）。" if not g["problems"] else "这些文件都合不了：")
                   + "".join(f"\n· {p}" for p in g["problems"]), problems=g["problems"])

    key = _key(args, src, out_dir)
    fp = _fingerprint(files)
    state, state_path, notes = _resume(out_dir, key, fp, args.restart)
    if state is None:
        state, state_path = _plan(args, cfg, src, is_dir, out_dir, files, g, key, fp)
        state["pre_warnings"] += notes     # 「表格有改动、从头重新合」要一直带到最后那次的结果里
        notes = []

    def save():
        with open(state_path, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False)

    def progress(frac, stage=""):
        eprint(f"[进度 {int(frac * 100)}%] {stage}")

    os.makedirs(out_dir, exist_ok=True)
    save()
    st = state["st"]
    try:
        done = em.merge_step(st, [tuple(f) for f in state["files"]], state["cfg"], state["mode"],
                             required=state["required"], until=T_START + min(args.budget, BUDGET), log=eprint,
                             progress=progress, commit=save, labels=state["labels"], people=state["people"])
    except Exception as e:
        emit_error(f"汇总出错：{e}。原表没有被改动，可以加 --restart 从头再试。")

    if not done:
        total = len(state["files"])
        emit({"ok": True, "done": False, "progress": f"{st['next']}/{total}",
              "message": f"表比较多，已合 {st['next']}/{total} 个表、进度已存好。"
                         f"原样再运行一次同一条命令就接着合（每次约 {int(BUDGET)} 秒）。",
              "next": "把进度用一句话告诉用户，然后原样再运行一次同一条命令，直到返回 done:true。"})
        return

    final = state["final"]
    if os.path.exists(st["output"]):
        if os.path.exists(final):
            final = em._result_path(out_dir, state["title"])
        os.replace(st["output"], final)
    try:
        os.remove(state_path)
    except OSError:
        pass
    output = final if os.path.exists(final) else ""
    opened = open_folder(out_dir) if args.open and output else False
    merged = [s for s in st["sources"] if s["status"] == "合并"]
    rows = sum(s["rows"] for s in merged)
    written = sum(s["next_row"] - s["h"] - 1 for s in st["sheets"].values())
    warnings = state["pre_warnings"] + notes
    msg = _message(state, st, output, rows, written, opened)
    if not output:
        emit_error("这些表里没有能合并的数据。" + ("\n" + msg if msg else ""),
                   problems=state["pre_problems"])
    saved = save_recipe_from(args, dict(cfg, collect_check=args.checked), "collect", input_name=src)
    emit({"ok": True, "done": True, "output": output, "output_dir": out_dir, "files": [output],
          "recipe": recipe["name"] if recipe else "", "recipe_notes": recipe_notes, "recipe_saved": saved,
          "summary": {"mode": state["mode"], "sources": len({s["file"] for s in merged}), "sheets_merged": len(merged),
                      "result_sheets": len(st["sheets"]), "rows": rows, "rows_written": written,
                      "not_merged": len([s for s in st["sources"] if s["status"] != "合并"]),
                      "duplicates": len(st["dupes"]), "failed_files": len(st["errors"]),
                      "skipped_files": len(state["pre_problems"])},
          "message": msg, "warnings": warnings, "opened": opened})


def _config(args, src, out_dir):
    cfg = {"header_mode": "row", "header_row": args.header_row} if args.header_row else {
        "header_mode": "auto", "header_row": 1}
    cfg.update(input_path=src, output_path=out_dir, collect_expect=args.expect, collect_required=args.required,
               collect_by=args.by, collect_title=args.title, collect_record=os.path.abspath(args.record) if args.record else "",
               grid_keys=[], id_keys=[])
    return cfg


def _key(args, src, out_dir):
    """同一个汇总任务的标识：输入和选项都一样，再运行时才接着上次的进度做。"""
    parts = [os.path.normcase(src), os.path.normcase(out_dir), args.title, str(args.header_row), str(args.checked),
             args.expect, args.required, args.record, args.by]
    return hashlib.sha1("\x1f".join(parts).encode("utf-8")).hexdigest()[:16]


def _fingerprint(files):
    """表格的路径、大小、修改时间：上次没做完、这期间表格有改动，就从头合，免得合出新旧混杂的结果。"""
    h = hashlib.sha256()
    for path, _rel in files:
        try:
            s = os.stat(path)
            h.update(f"{os.path.normcase(path)}|{s.st_size}|{int(s.st_mtime)}\n".encode())
        except OSError:
            h.update(f"{path}|missing\n".encode())
    return h.hexdigest()[:16]


def _resume(out_dir, key, fp, restart):
    """找上次没做完的同一个任务。返回 (state, state 路径, 提醒)；没有可接着做的返回 (None, None, [])。"""
    if not os.path.isdir(out_dir):
        return None, None, []
    notes = []
    for n in sorted(os.listdir(out_dir)):
        if not n.endswith(STATE_SUFFIX):
            continue
        path = os.path.join(out_dir, n)
        try:
            with open(path, encoding="utf-8") as f:
                state = json.load(f)
        except (OSError, ValueError):
            continue
        if state.get("key") != key:
            continue
        if restart or state.get("fingerprint") != fp:
            _discard(path, state)
            if not restart:
                notes.append("上次没合完以后表格有改动，已从头重新合")
            continue
        eprint(f"接着上次的进度合（已合 {state['st']['next']}/{len(state['files'])} 个表）")
        return state, path, notes
    return None, None, notes


def _discard(state_path, state):
    for p in (state_path, state.get("st", {}).get("output", "")):
        try:
            if p and os.path.exists(p):
                os.remove(p)
        except OSError:
            pass


def _plan(args, cfg, src, is_dir, out_dir, files, g, key, fp):
    """第一次运行：读表头、定合并方式，建好可续做的状态。"""
    labels = people = None
    lines = []
    if args.checked:
        files, labels, people, lines, title, required, mode = _checked(cfg, files, g["bad"])
    else:
        title = args.title or (os.path.basename(src.rstrip(os.sep)) if is_dir else os.path.splitext(os.path.basename(src))[0])
        res = em.scan(files, cfg, deadline=time.time() + SCAN_SECONDS, log=eprint)
        _need_complete(res, files)
        files, res = _drop_name_lists(files, res, g)
        if not res["sheets"]:
            _no_data(src, res)
        mode, required = em.pick_mode(len(files), res["sheets"]), {}
    final = em._result_path(out_dir, title)
    partial = final[:-len(".xlsx")] + PART_SUFFIX
    state_path = final[:-len(".xlsx")] + STATE_SUFFIX
    eprint(f"开始汇总 {len(files)} 个表…" if len(files) > 1 else "开始把各个 sheet 合成一张…")
    return {"key": key, "fingerprint": fp, "files": files, "cfg": cfg, "mode": mode, "required": required,
            "labels": labels, "people": people, "title": title, "final": final, "started": time.time(),
            "checked_lines": lines, "pre_problems": g["problems"], "pre_warnings": g["warnings"],
            "st": em.merge_start(partial)}, state_path


def _drop_name_lists(files, res, g):
    """应交名单 / 发送清单这类表和数据表对不上时不合，也不当成「表头对不上」报错——
    它本来就不是数据（多半是查收时放进来的）。和数据表是同一个模板的照合。"""
    lists = {os.path.normcase(p) for p in g["lists"]}
    idx = [i for i, (p, _r) in enumerate(files) if os.path.normcase(p) in lists]
    if not idx or len(idx) == len(files):
        return files, res
    groups = em.compare(res["sheets"], em.pick_mode(len(files), res["sheets"]))
    bad = {it["file"] for gr in groups.values() for it in gr["items"] if it["mismatch"]}
    good = {it["file"] for gr in groups.values() for it in gr["items"] if not it["mismatch"]}
    drop = {i for i in idx if i in bad and i not in good}
    if not drop:
        return files, res
    for i in sorted(drop):
        g["warnings"].append(f"「{files[i][1]}」是名单表，没有合进汇总表")
    keep = [i for i in range(len(files)) if i not in drop]
    new = {old: n for n, old in enumerate(keep)}
    res = dict(res, sheets=[dict(s, file=new[s["file"]]) for s in res["sheets"] if s["file"] in new])
    return [files[i] for i in keep], res


def _need_complete(res, files):
    if not res["complete"]:
        emit_error(f"表格太多，{SCAN_SECONDS} 秒内只读了 {res['files_scanned']}/{len(files)} 个表的表头。"
                   f"请把表格分到几个文件夹里分批汇总。")


def _no_data(src, res):
    why = f"读不了：{res['errors'][0]['file']}：{res['errors'][0]['error']}" if res["errors"] else \
        "没有认出带表头和数据的 sheet"
    emit_error(f"「{src}」里{why}。如果表头不在前 20 行，用 --header-row 指定表头在第几行。")


def _checked(cfg, files, bad):
    """按查收结果只合交了的表（与桌面版 run_collect 同一规则）：没交、空表不合，名单外的照合，「来源」写交表人。

    读不了的文件（.et、网页格式的 .xls、加了口令的）也一起交给查收认人——它们的主人是「交了但读不了」，
    不能和真没交的混在一起说；只是不合进汇总表。
    """
    ctx = em._prepare(cfg, eprint)
    bad_rel = {rel for _p, rel in bad}
    files = sorted(list(files) + list(bad), key=lambda f: order_key(f[1]))
    res = em.scan(files, cfg, deadline=time.time() + SCAN_SECONDS, log=eprint,
                  by=ctx["claim_by"] or None, fp_keys=ctx["fp_keys"])
    _need_complete(res, files)
    report = em.analyze(files, res, ctx["expect"], ctx["wanted"], ctx["record"], ctx["by"], ctx["rows_mode"])
    rel_idx = {rel: i for i, (_p, rel) in enumerate(files)}
    labels, take = {}, set()
    for p in report["people"]:
        if p["status"] in (em.ST_MISSING, em.ST_EMPTY):
            continue
        for rel in p["files"]:
            labels[rel] = p["name"]
            take.add(rel_idx[rel])
    if report["mode"] == "files":
        take |= {rel_idx[r] for r in report["extra"]}
    else:
        take = set(range(len(files)))
    take = {i for i in take if files[i][1] not in bad_rel}
    sel = [files[i] for i in sorted(take)]
    if not sel:
        emit_error("还没有交了、填了内容的表，没有可以汇总的。先用 er_check.py 看看谁没交。")
    sel_idx = {files[i][1]: n for n, i in enumerate(sorted(take))}
    sel_sheets = [dict(s, file=sel_idx[s["rel"]]) for s in res["sheets"] if s["rel"] in sel_idx]
    mode = em.pick_mode(len(sel) if report["mode"] == "rows" else max(2, len(sel)), sel_sheets)
    groups = em.compare(sel_sheets, mode)
    em.completeness(groups, ctx["wanted"])
    c = em.counts(report)
    lines = []
    if c["expected"] is not None:
        lines.append(f"按查收结果只合交了的表：应交 {c['expected']}，已交 {c['submitted']}"
                     + (f"，其中 {c['empty']} 份是空表没合" if c["empty"] else ""))
    for what, names in (("还没交", [p["name"] for p in report["people"] if p["status"] == em.ST_MISSING]),
                        ("交了空表", [p["name"] for p in report["people"] if p["status"] == em.ST_EMPTY]),
                        ("交了但文件读不了", [p["name"] for p in report["people"]
                                         if p["status"] not in (em.ST_MISSING, em.ST_EMPTY)
                                         and p["files"] and all(f in bad_rel for f in p["files"])])):
        if names:
            lines.append(f"{what}、没有汇总的（{len(names)}）：" + "、".join(names[:30]) + (" 等" if len(names) > 30 else ""))
    if report["mode"] == "files" and report["extra"]:
        lines.append(f"名单外的文件也合进来了（{len(report['extra'])}），「来源」写文件名：" + "、".join(report["extra"][:5]))
    files_mode = report["mode"] == "files"
    people = [{k: p[k] for k in ("name", "status", "flags", "files", "notes")} for p in report["people"]] \
        if files_mode else None
    return (sel, labels if files_mode else None, people, lines, ctx["title"],
            {k: g["required"] for k, g in groups.items()}, mode)


def _message(state, st, output, rows, written, opened):
    """一段可以直接转述给用户的话：合了什么、哪些没合为什么、行数核对、结果在哪。"""
    lines = [em.merge_text(st, state["title"]).rstrip("\n")]
    lines += state["checked_lines"]
    if state["pre_problems"]:
        lines.append(f"这些文件没收进来（{len(state['pre_problems'])}）：")
        lines += [f"  {p}" for p in state["pre_problems"]]
    if state["required"] and any(state["required"].values()):
        holes = [s for s in st["sources"] if s.get("blanks")]
        if holes:
            n = sum(v for s in holes for v in s["blanks"].values())
            lines.append(f"有 {len(holes)} 个表没填完整，共空着 {n} 格，在汇总表里标成了浅红色")
    totals = sum(s.get("totals", 0) for s in st["sources"] if s["status"] == "合并")
    check = f"核对：合进来 {rows} 行，汇总表里 {written} 行" + ("，一致" if rows == written else "，不一致，请核对")
    if totals:
        check += f"；各表自己的合计 / 小计 {totals} 行没合进来，需要合计请在汇总表里重新求和"
    lines.append(check)
    if output:
        lines.append(f"汇总表：{output}" + ("（已为你打开所在文件夹）" if opened else ""))
    return "\n".join(lines)


if __name__ == "__main__":
    main()
