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
import re
import shutil
import signal
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import (  # noqa: E402
    _join, add_recipe_args, apply_recipe_args, bootstrap, build_config, default_output_for, emit, emit_error,
    ensure_deps, eprint, list_outputs, match_column, open_folder, save_recipe_from, sheet_headers,
)

bootstrap()
ensure_deps()  # 必须排在下面这些 import 之前：缺包时要给人话提示 / 自动装，而不是 ImportError 堆栈

from core.splitter import UNASSIGNED, run_split  # noqa: E402
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
    ap.add_argument("--output", default="",
                     help="输出根目录（实际产出会在其下多一层时间戳子目录）。不填 = 源表旁边的「拆分结果」，"
                          "与桌面版默认一致")
    ap.add_argument("--open", action="store_true",
                     help="拆完在文件管理器里打开结果目录（桌面版跑完自动弹出结果的等价物）。"
                          "AI 在用户本机运行时应总是加上")
    ap.add_argument("--by", dest="split_column", default=None, help="拆分字段（列名）")
    ap.add_argument("--values", default="", help="只拆这些取值，逗号分隔；留空 = 自动枚举全部取值")
    ap.add_argument("--to-person", dest="person_column", default=None, metavar="人字段",
                     help="附加产出「到人」结果。三种写法：① 列名（所有 sheet 共用）；"
                          "② auto（每个 sheet 自动挑人字段）；"
                          "③ \"sheet名:人字段,sheet名:人字段\"（各 sheet 人字段不同）。"
                          "单文件输入也支持，由 scripts 包装层补做二级拆分")
    ap.add_argument("--person-filter", default="", help="到人只处理文件名命中这些关键词的表，逗号分隔；留空=全部")
    ap.add_argument("--no-zip", action="store_true",
                     help="不打包 ZIP（默认：每组一个 ZIP + 全部结果一个总包，方便直接转发）")
    ap.add_argument("--no-timestamp", action="store_true",
                     help="输出直接写到 --output，不再套一层 {MMDDHHMM}结果 时间戳子目录"
                          "（路径可预测，方便脚本接力/直接告诉用户文件在哪）")
    ap.add_argument("--merge", action="store_true",
                     help="同一取值跨源文件合并到一个输出文件（默认关闭：按原表各自拆分，只在同一取值下打包在一起）")
    ap.add_argument("--fast", action="store_true", help="快速模式：数据行只写值不保留格式，速度换取不保留样式")
    ap.add_argument("--keep-formulas", action="store_true",
                     help="尽量保留同行公式（跨行/汇总/跨表公式仍会落成当前缓存数值，不会发送算错的公式）")
    ap.add_argument("--keep-hidden", action="store_true",
                    help="隐藏的 sheet 和列也照样发出（默认：隐藏的 sheet 不发，隐藏列的内容清空、仍隐藏）")
    ap.add_argument("--fuzzy-match", action="store_true", help="--values 按包含匹配而不是精确匹配")
    ap.add_argument("--header-mode", choices=["auto", "row", "keyword"], default="auto")
    ap.add_argument("--header-row", type=int, default=1, help="header-mode=row 时的 1 基行号")
    ap.add_argument("--grid-keys", default="", help="header-mode=keyword 时第一类关键词，逗号分隔")
    ap.add_argument("--id-keys", default="", help="header-mode=keyword 时第二类关键词，逗号分隔")
    ap.add_argument("--skip-values", default=None,
                     help="拆分列中要忽略的取值，逗号分隔；不填则用默认（合计/小计/总计/平均/空）")
    ap.add_argument("--alias-json", default="", help='取值归并，如 \'{"东区":["东片区","东部"]}\'')
    ap.add_argument("--config", default=None, help="直接传一份完整 config JSON（专家逃生口，字段同桌面版 user_config.json）")
    ap.add_argument("--include-old-results", action="store_true",
                     help="输入文件夹里有以前的拆分结果目录时也照样拆（默认拦下来，免得旧结果被重复拆）")
    ap.add_argument("--dry-run", action="store_true", help="只解析配置并打印，不实际拆分、不写盘")
    add_recipe_args(ap)
    args = ap.parse_args()
    if args.recipe and args.config:
        emit_error("--recipe 和 --config 二选一")
    recipe, recipe_notes, recipe_extra = apply_recipe_args(ap, args, "excel", {
        "split_column": ("split_column", lambda v, s: v or None),
        "selected_values": ("values", _join),
        "person_column": ("person_column", lambda v, s: (v or None) if s.get("to_person") else None),
        "person_file_filter": ("person_filter", _join),
        "make_zip": ("no_zip", lambda v, s: not v),
        "merge_across_files": ("merge", lambda v, s: v),
        "preserve_format": ("fast", lambda v, s: not v),
        "keep_formulas": ("keep_formulas", lambda v, s: v),
        "keep_hidden": ("keep_hidden", lambda v, s: v),
        "exact_match": ("fuzzy_match", lambda v, s: not v),
        "header_mode": ("header_mode", lambda v, s: v),
        "header_row": ("header_row", lambda v, s: v),
        "grid_keys": ("grid_keys", _join),
        "id_keys": ("id_keys", _join),
        "skip_values": ("skip_values", _join),
        "value_alias_map": ("alias_json", lambda v, s: json.dumps(v, ensure_ascii=False) if v else ""),
    })
    recipe_extra.pop("to_person", None)          # 跟着 person_column 走

    if not args.output:
        args.output = default_output_for(args.input)

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
            keep_hidden=args.keep_hidden or None,
            no_timestamp=args.no_timestamp,
            exact_match=(False if args.fuzzy_match else None),
            header_mode=args.header_mode,
            header_row=args.header_row,
            grid_keys=_split_list(args.grid_keys),
            id_keys=_split_list(args.id_keys),
            skip_values=[s.strip() for s in args.skip_values.split(",")] if args.skip_values is not None else None,
            value_alias_map=alias_map or None,
        )

    config.update(recipe_extra)                     # 方案里没有对应参数的设置（如填写列）
    if recipe:
        eprint(f"📋 套用方案「{recipe['name']}」")
    if not os.path.exists(config["input_path"]):
        emit_error(f"找不到输入路径：{config['input_path']}")
        return
    if not config.get("split_column"):
        emit_error("拆分字段（split_column）为空，无法拆分")
        return

    # ---- 开跑前先确认拆分字段真的存在 ----
    # 不查的话内核会把每个 sheet 都当「找不到该列」跳过，最后「成功」产出 0 个文件——
    # AI 拿到 ok:true 就会告诉用户「拆完了」。这是用户最难自己发现的一种失败。
    resolved, also = _preflight_column(config)
    if resolved is None:
        return
    warnings = []
    if resolved != config["split_column"]:
        eprint(f"「{config['split_column']}」按列名包含匹配到「{resolved}」")
    if also:
        warnings.append(f"「{config['split_column']}」同时能匹配 {also}，实际按「{resolved}」拆；"
                        f"不对的话把 --by 写成完整列名")
    if os.path.isdir(config["input_path"]) and not args.include_old_results:
        stale = _find_stale_results(config["input_path"], config["output_path"])
        if stale:
            emit_error(
                f"输入文件夹里有 {len(stale)} 个以前的拆分结果目录（如 {'、'.join(stale[:3])}），"
                f"直接拆会把它们当成数据再拆一遍、行数翻倍。先问用户：删掉/移走它们，"
                f"还是确实要一起拆（那就加 --include-old-results）。",
                stale_dirs=stale)
            return
    lock_warn = _lock_file_warning(config["input_path"])
    if lock_warn:
        warnings.append(lock_warn)
    for w in warnings:
        eprint(f"⚠ {w}")

    if args.dry_run:
        eprint("dry-run：不写盘，仅打印解析后的配置")
        emit({"ok": True, "dry_run": True, "split_column_resolved": resolved,
              "warnings": warnings, "config": config})
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

    summary = {}

    def log_fn(msg):
        # 内核收尾的 [SUMMARY] 行截下来放进结果 JSON：很多 Agent 只看 stdout，
        # 让 AI 去 stderr 里翻这一行，等于大概率没人看
        s = str(msg)
        if s.startswith("[SUMMARY] "):
            try:
                summary.update(json.loads(s[len("[SUMMARY] "):]))
            except ValueError:
                pass
        eprint(s)

    out_root = config["output_path"]
    existed = set(os.listdir(out_root)) if os.path.isdir(out_root) else set()

    try:
        output_path = run_split(config, log_fn=log_fn, progress_fn=progress_fn,
                                 stop_flag=lambda: state["stopped"])
    except Exception as e:
        emit_error(f"拆分失败：{e}")
        return

    if os.path.basename(output_path) in existed:
        # 旧内核时间戳只到分钟，一分钟内重跑会写进上一次的目录（新内核已改为加序号）
        warnings.append(f"本次结果和一分钟内上一次的结果写进了同一个文件夹 {output_path}，"
                        f"两次的文件混在一起；建议删掉这个文件夹再拆一次")
        eprint(f"⚠ {warnings[-1]}")

    if summary.get("groups") == 0 and not state["stopped"]:
        hint = (f"「--values {','.join(config['selected_values'])}」里的取值在表中一个都没对上，"
                f"先用 er_inspect.py --column {resolved} 看真实取值"
                if config.get("selected_values") else
                f"字段「{resolved}」下没有可拆的取值（可能全是空或合计行）")
        if summary.get("unassigned"):
            hint += f"；有 {summary['unassigned']} 行没写「{resolved}」，放在 {UNASSIGNED} 里"
        emit_error(f"没有拆出任何文件：{hint}。", output_path=output_path, summary=summary)
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

    bundle = None
    if not args.no_zip:
        _zip_group_folders(output_path, log_fn=eprint)
        bundle = _make_bundle(output_path, resolved, log_fn=eprint)

    after = _result_warnings(summary, resolved)
    warnings.extend(after)
    for w in after:
        eprint(f"⚠ {w}")

    saved = save_recipe_from(args, config, "excel", input_name=config["input_path"])
    files = list_outputs(output_path)
    if bundle:  # 「拿去发」的那一个文件排第一，AI 先 present 它
        files = [bundle] + [f for f in files if f != bundle]
    opened = open_folder(output_path) if args.open else False
    emit({"ok": True, "output_path": output_path,
          "split_column_resolved": resolved,
          "summary": summary,
          "bundle_zip": bundle,
          "files_total": len(files),
          "files": files[:10],
          "message": _human_message(summary, resolved, files, output_path, opened,
                                    state["stopped"], bundle),
          "warnings": warnings,
          "opened": opened,
          "recipe": recipe["name"] if recipe else "", "recipe_notes": recipe_notes, "recipe_saved": saved,
          "person_split_by": ("wrapper" if need_wrapper_person
                              else ("vendor" if person_map else None)),
          "person_files": person_count if need_wrapper_person else None,
          "stopped": state["stopped"]})


def _preflight_column(config):
    """确认 split_column 在输入里存在。找到返回 (实际列名, 其余也命中的列)；
    找不到直接 emit_error 并返回 (None, [])。目录输入先看样本表，样本里没有再逐个找
    （只读表头，很快），一个都没有才报错——各表结构不一致时不能只凭样本就拒绝。"""
    target = config["split_column"]
    path = config["input_path"]
    if os.path.isfile(path):
        files = [path]
    else:
        out_abs = os.path.abspath(config["output_path"])
        files = sorted(
            os.path.join(r, n) for r, _d, ns in os.walk(path)
            if not os.path.abspath(r).startswith(out_abs)
            for n in ns
            if n.lower().endswith((".xlsx", ".xls")) and not n.startswith("~$")
            and not n.lower().endswith("__tmp__.xlsx"))
        if not files:
            emit_error(f"目录里没找到任何 .xlsx / .xls 文件：{path}")
            return None, []
    sample_cols = None
    for f in files:
        try:
            sheets = sheet_headers(f, config)
        except Exception:
            continue  # 坏文件交给内核去跳过并记日志
        cols = [c for _s, h, cs in sheets if h != -1 for c in cs]
        if sample_cols is None:
            sample_cols = {s: cs for s, h, cs in sheets if h != -1}
        hit, also = match_column(list(dict.fromkeys(cols)), target)
        if hit:
            return hit, also
    emit_error(
        f"表里没有「{target}」这个字段，没有开始拆分。"
        f"{'样本表' if len(files) > 1 else '这张表'}可选的字段：{sample_cols or '（没识别出表头）'}",
        available_columns=sample_cols or {})
    return None, []


def _find_stale_results(input_dir, output_root):
    """输入文件夹里不被本次输出目录覆盖的历史拆分结果目录（与桌面版
    _find_stale_results 同一规则）。内核只跳过本次的 output_root 子树，
    别处的旧结果会被当成数据重复拆。"""
    out = os.path.normcase(os.path.abspath(output_root))
    hits = []
    for root, dirs, _ in os.walk(input_dir):
        for d in list(dirs):
            full = os.path.normcase(os.path.abspath(os.path.join(root, d)))
            if full == out or full.startswith(out + os.sep):
                dirs.remove(d)
                continue
            if d == "拆分结果" or d.endswith("_拆分结果") or re.fullmatch(r"\d{8}结果(_\d+)?", d):
                hits.append(os.path.join(root, d))
                dirs.remove(d)
    return hits


def _lock_file_warning(path):
    """源表正被 Excel/WPS 打开时，同目录会有 ~$ 锁文件；读到的是上次保存的内容。"""
    if os.path.isfile(path):
        d, n = os.path.split(path)
        # Excel 用 ~$ + 全名；Word 系的习惯是吃掉前两个字，两种都认
        lock_names = ("~$" + n, "~$" + n[2:])
        locked = [n] if any(os.path.exists(os.path.join(d, ln)) for ln in lock_names) else []
    else:
        locked = [n[2:] for _r, _d, ns in os.walk(path) for n in ns if n.startswith("~$")]
    if locked:
        return (f"{locked[:3]} 正被 Excel/WPS 打开着，拆的是它上次保存的内容；"
                f"有没保存的修改请先保存再拆")
    return ""


def _zip_group_folders(output_path, log_fn=None):
    """给还没有 ZIP 的分组文件夹补打包。内核只在**目录输入**时按组打 ZIP；单文件 + 到人
    （包装层产出 {组}/到人/...）原来就只有散文件，要发给每组负责人还得用户自己压。"""
    for name in sorted(os.listdir(output_path)):
        folder = os.path.join(output_path, name)
        if not os.path.isdir(folder) or os.path.exists(folder + ".zip"):
            continue
        if any(os.scandir(folder)):
            try:
                shutil.make_archive(folder, "zip", folder)
                if log_fn:
                    log_fn(f"  📦 {name}.zip")
            except Exception as e:
                if log_fn:
                    log_fn(f"  ⚠ {name}.zip 打包失败：{e}")


def _make_bundle(output_path, column, log_fn=None):
    """把这次的全部结果打成**一个** ZIP，放在结果目录里、排在最前——用户最常见的下一步
    就是「整包转给谁」。分组 ZIP 不重复放进去（内容和分组文件夹一样，只会让体积翻倍）；
    文件名带上拆分字段，脱离对话单独转发时也看得懂是什么。失败返回 None，不影响拆分结果。"""
    from core.utils import safe_filename
    members = []
    for root, _dirs, names in os.walk(output_path):
        for n in names:
            low = n.lower()
            if low.endswith(".zip") or low.startswith("~$") or n in ("运行日志.txt", UNASSIGNED):
                continue   # 「请核对_无归属」是给用户自己核对的，不跟着整包发出去
            members.append(os.path.join(root, n))
    if not members:
        return None
    stamp = os.path.basename(output_path)[:8]
    path = os.path.join(output_path, f"全部结果_按{safe_filename(column)}拆分_{stamp}.zip")
    try:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in sorted(members):
                zf.write(f, os.path.relpath(f, output_path))
    except Exception as e:
        if log_fn:
            log_fn(f"⚠ 打总包失败：{e}（分组文件都在，不影响使用）")
        return None
    if log_fn:
        log_fn(f"📦 已打成一个总包：{os.path.basename(path)}（{len(members)} 个文件）")
    return path


def _result_warnings(summary, column):
    """拆完之后必须告诉用户的几件事（v3.0）：没写分组的行、被拦下的隐藏内容、只拆出 1 组。"""
    out = []
    if summary.get("unassigned"):
        out.append(f"有 {summary['unassigned']} 行没写「{column}」，没分进任何一组，"
                   f"放在结果目录的 {UNASSIGNED} 里（一个源 sheet 一页），请用户核对后补上再拆或手动发")
    hidden = []
    if summary.get("hidden_sheets"):
        hidden.append(f"{summary['hidden_sheets']} 个隐藏的 sheet 没发出")
    if summary.get("hidden_cols"):
        hidden.append(f"{summary['hidden_cols']} 个隐藏列的内容已清空（仍隐藏）")
    if hidden:
        out.append("源表里有隐藏内容：" + "，".join(hidden) + "；用户确认要一起发时加 --keep-hidden 重拆")
    if summary.get("groups") == 1 and not summary.get("stopped"):
        out.append(f"只拆出了 1 组：每一行的「{column}」都一样，可能选错了拆分字段，跟用户确认一下")
    return out


def _human_message(summary, column, files, output_path, opened, stopped, bundle=None):
    """一句可以直接念给用户的话，AI 不用自己拼数字。"""
    if not summary:
        return f"已完成，结果在：{output_path}"
    parts = [f"按「{column}」拆成 {summary.get('groups', 0)} 组，"
             f"共 {summary.get('files', 0)} 个文件、{summary.get('rows', 0)} 行数据"]
    if summary.get("zips"):
        parts.append(f"每组已打包成 ZIP（{summary['zips']} 个），可以直接发")
    person = sum(1 for f in files if os.path.basename(os.path.dirname(f)) == "到人")
    if person:
        parts.append(f"其中按人拆出 {person} 份")
    msg = "；".join(parts) + "。"
    if stopped:
        msg = "⏹ 已按要求中途停止，下面是已完成的部分：" + msg
    problems = []
    if summary.get("skipped_sheets"):
        problems.append(f"{summary['skipped_sheets']} 个 sheet 里没有这个字段，已跳过")
    if summary.get("failed_files"):
        problems.append(f"{summary['failed_files']} 个文件读取失败")
    if summary.get("failed_saves"):
        problems.append(f"{summary['failed_saves']} 个文件保存失败（可能被 Excel 占用）")
    if summary.get("unassigned"):
        problems.append(f"{summary['unassigned']} 行没写「{column}」，放在 {UNASSIGNED} 里请核对")
    if summary.get("hidden_sheets") or summary.get("hidden_cols"):
        problems.append("源表里隐藏的 sheet / 列没有发出去")
    if problems:
        msg += "⚠ " + "，".join(problems) + "，详情见日志。"
    if bundle:
        size = os.path.getsize(bundle) / 1024 / 1024
        msg += (f"\n全部结果已打成一个压缩包，可以直接转发：{os.path.basename(bundle)}"
                f"（{size:.1f} MB）")
    msg += f"\n结果在：{output_path}" + ("（已为你打开这个文件夹）" if opened else "")
    return msg


if __name__ == "__main__":
    main()
