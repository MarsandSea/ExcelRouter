#!/usr/bin/env python
"""
excelrouter · 文件按网格加密分发：同一批 PDF / Word / Excel / PPT，按「网格→密码」映射清单，
为每个网格生成一份专属打开密码的副本（PDF 还可加专属水印，可溯源），并输出含明文密码的分发清单。

Copyright (c) 2026 AbeLin
MIT License

用法：
  # 起步：手上还没有映射清单时，先生成一份模板给用户填（网格/密码/接收人 + 填写说明）
  python er_pdf_dist.py --mapping 映射.xlsx --make-template 映射.xlsx

  # 第一步：不确定映射清单的列名时，先看一眼表头
  python er_pdf_dist.py --mapping 映射.xlsx --list-columns

  # 可选：懒得想密码时，让工具给空白网格生成随机密码（写入新文件，不动原清单）
  python er_pdf_dist.py --mapping 映射.xlsx --grid-col 网格 --random-password

  # 第二步：确认列名后正式分发（PDF、Word、Excel、PPT 可以混着给）
  python er_pdf_dist.py --files 通报.pdf 明细.xlsx 说明.docx --mapping 映射.xlsx \
      --grid-col 网格 --password-col 密码 --output 分发结果

  # 只给 PDF 加水印、不设密码（清单只要网格列）
  python er_pdf_dist.py --files 通报.pdf --mapping 映射.xlsx --grid-col 网格 --no-encrypt

输出：stdout 一行 JSON。分发清单「分发清单.xlsx」含明文密码，务必提醒用户
不要把这份清单和加密后的文件一起发给同一批人。

--pdf 是 --files 的旧名字，照样能用。
Word / Excel / PPT 只加打开密码、不加水印（要水印请先另存为 PDF）；只支持 .docx / .xlsx / .pptx
这类新格式，老格式 .doc / .xls 请先另存。

--random-password 会用「原文件名_含密码.xlsx」另存一份新清单，正式分发时改用
那份新清单（JSON 里的 mapping_with_passwords 就是它的绝对路径）。
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import (  # noqa: E402
    add_recipe_args, apply_recipe_args, bootstrap, build_config, emit, emit_error, ensure_deps, eprint, list_outputs,
    open_folder, save_recipe_from,
)

bootstrap()
# 整个脚本都是分发流程（连 --make-template 也是它的第一步），一次把 PDF 依赖备齐；
# 必须排在 core.pdf_dist 的 import 之前，否则缺包时是 ImportError 堆栈而不是人话。
# Office 加密要的 msoffcrypto-tool 等解析完参数、确实有 Word/Excel/PPT 时再检查。
ensure_deps(need_pdf=True)

import core.pdf_dist as pdf_dist  # noqa: E402
from core.pdf_dist import (  # noqa: E402
    MANIFEST_NAME,
    fill_random_passwords,
    list_mapping_columns,
    run_pdf_dist,
    write_mapping_template,
)

OFFICE_EXTS = (".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm")
OLD_EXTS = (".doc", ".xls", ".ppt", ".et", ".wps", ".dps")


def main():
    ap = argparse.ArgumentParser(
        description="文件按网格加密分发（PDF / Word / Excel / PPT 专属密码 + PDF 水印 + 分发清单）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    ap.add_argument("--files", "--pdf", dest="files", nargs="+", default=[],
                    help="待分发的文件，可多个：PDF、.docx、.xlsx、.pptx（--pdf 是旧名字）")
    # 不设 required=True：--make-template 要生成一个「还不存在」的清单，此时 --mapping 本就该缺席。
    # 缺 --mapping 的场景在下面按需校验，报错信息也能带上一句「先用 --make-template 生成」。
    ap.add_argument("--mapping", default="", help="「网格→密码」映射清单 xlsx 路径")
    ap.add_argument("--list-columns", action="store_true",
                     help="只列出映射清单第 1 行的表头，不做分发（先跑这个确认列名）")
    ap.add_argument("--make-template", metavar="PATH", default="",
                     help="生成一份「网格/密码/接收人」映射清单模板到 PATH（含两行示例和一张填写说明 sheet），"
                          "不做分发。用户手上还没有清单时用这个起步")
    ap.add_argument("--random-password", action="store_true",
                     help="给映射清单里密码为空的网格自动生成 8 位随机密码，另存为"
                          "「原名_含密码.xlsx」（不改动原清单），不做分发；"
                          "正式分发时改用新清单。需要 --grid-col")
    ap.add_argument("--grid-col", default="", help="映射清单里的网格列名")
    ap.add_argument("--password-col", default="", help="映射清单里的密码列名（--no-encrypt 时不用）")
    ap.add_argument("--receiver-col", default="",
                    help="映射清单里的接收人列名（选填；不填时自动认「接收人/联系人/负责人/姓名」这类列）")
    ap.add_argument("--output", default="", help="输出目录；不填则用「第一个文件所在目录/分发结果」")
    ap.add_argument("--no-encrypt", action="store_true",
                    help="不设打开密码，只给 PDF 加水印（清单不需要密码列；Word/Excel/PPT 不能这样用）")
    ap.add_argument("--no-watermark", action="store_true", help="不加水印（只设密码）")
    ap.add_argument("--watermark-text", default="{grid} {date}",
                    help="水印模板，支持 {grid} 网格名、{receiver} 接收人、{date} 日期")
    ap.add_argument("--opacity", type=float, default=0.15, help="水印透明度 0~1")
    ap.add_argument("--angle", type=float, default=45, help="水印旋转角度")
    ap.add_argument("--open", action="store_true",
                     help="分发完在文件管理器里打开结果目录。AI 在用户本机运行时应总是加上")
    ap.add_argument("--font", default="", help="手动指定水印字体 .ttf/.otf/.ttc 路径（覆盖自动探测）")
    add_recipe_args(ap)
    args = ap.parse_args()
    recipe, recipe_notes, _extra = apply_recipe_args(ap, args, "pdf", {
        "pdf_grid_column": ("grid_col", lambda v, s: v),
        "pdf_password_column": ("password_col", lambda v, s: v),
        "pdf_receiver_column": ("receiver_col", lambda v, s: v),
        "pdf_encrypt": ("no_encrypt", lambda v, s: not v),
        "pdf_watermark": ("no_watermark", lambda v, s: not v),
        "pdf_watermark_text": ("watermark_text", lambda v, s: v),
        "pdf_watermark_opacity": ("opacity", lambda v, s: v),
        "pdf_watermark_angle": ("angle", lambda v, s: v),
    })

    # ---- 起步：生成映射清单模板（唯一不需要已有文件的分支，所以排在所有存在性检查之前）----
    if args.make_template:
        target = os.path.abspath(args.make_template)
        parent = os.path.dirname(target)
        if parent and not os.path.isdir(parent):
            emit_error(f"模板要写到的目录不存在：{parent}")
            return
        try:
            written = write_mapping_template(target)
        except Exception as e:
            emit_error(f"生成映射清单模板失败：{e}")
            return
        eprint(f"✅ 已生成映射清单模板：{written}")
        eprint("   第 2、3 行是示例行，填好真实网格后请删除；第二个 sheet「填写说明」用户可直接看。")
        emit({"ok": True, "template": written,
              "next": f"让用户填好后，用 --mapping \"{written}\" 先跑 --list-columns 确认列名"})
        return

    if not args.mapping:
        emit_error("缺少 --mapping（映射清单 xlsx）。手上还没有清单的话，"
                   "先跑 --make-template 路径.xlsx 生成一份模板。")
        return
    if not os.path.exists(args.mapping):
        emit_error(f"找不到映射清单：{args.mapping}")
        return

    # ---- 可选：给空白网格补随机密码（写新文件，原清单不动）----
    if args.random_password:
        if not args.grid_col:
            emit_error("--random-password 需要 --grid-col 指出哪一列是网格"
                       "（不确定列名就先跑 --list-columns 看表头）")
            return
        # 这里曾有一段包装层规避：上游 fill_random_passwords 在 password_col 传空串时
        # 会在表尾追加重复的「密码」列。上游 v2.7.2 已修（留空时先复用既有「密码」列，
        # 只填空白格），vendor 现在带的就是修好的版本，规避已删除 —— 留着会让人以为
        # 上游还坏着，也多一次不必要的清单读取。
        try:
            out_path, count = fill_random_passwords(
                os.path.abspath(args.mapping), args.grid_col, args.password_col)
        except Exception as e:
            emit_error(f"生成随机密码失败：{e}")
            return
        # 密码列可能原本不存在而被新建成「密码」，以产物实际表头为准，别猜
        out_cols = list_mapping_columns(out_path)
        pwd_col = args.password_col if args.password_col in out_cols else (
            "密码" if "密码" in out_cols else (out_cols[-1] if out_cols else "密码"))
        eprint(f"✅ 已为 {count} 个网格生成随机密码，新清单：{out_path}（原清单未改动）")
        eprint("⚠ 新清单含明文密码，只留给分发人自己用，不要随文件一起发出去。")
        emit({"ok": True, "mapping_with_passwords": out_path, "generated": count,
              "password_col": pwd_col,
              "next": f'正式分发改用这份新清单：--mapping "{out_path}" '
                      f'--grid-col {args.grid_col} --password-col {pwd_col}'})
        return

    if args.list_columns:
        try:
            cols = list_mapping_columns(args.mapping)
        except Exception as e:
            emit_error(f"读取映射清单表头失败：{e}")
            return
        emit({"ok": True, "columns": cols})
        return

    # ---- 正式分发：先把文件和选项的问题一次说清，别跑一半才发现 ----
    if not args.files:
        emit_error("缺少 --files（至少一个待分发的文件：PDF / .docx / .xlsx / .pptx），"
                   "或改用 --list-columns 只看映射清单表头")
        return
    for p in args.files:
        if not os.path.exists(p):
            emit_error(f"找不到文件：{p}")
            return
    exts = [os.path.splitext(p)[1].lower() for p in args.files]
    old = [os.path.basename(p) for p, e in zip(args.files, exts) if e in OLD_EXTS]
    if old:
        emit_error(f"{'、'.join(old)} 是老格式，加不了密码。请用 Word / Excel / WPS 另存为 "
                   ".docx / .xlsx / .pptx 后再分发", old_format=old)
        return
    bad = [os.path.basename(p) for p, e in zip(args.files, exts)
           if e != ".pdf" and e not in OFFICE_EXTS]
    if bad:
        emit_error(f"{'、'.join(bad)} 不是 PDF / Word / Excel / PPT，这个功能处理不了", unsupported=bad)
        return
    office = [os.path.basename(p) for p, e in zip(args.files, exts) if e in OFFICE_EXTS]
    encrypt = not args.no_encrypt
    watermark = not args.no_watermark
    if not encrypt and not watermark:
        emit_error("--no-encrypt 和 --no-watermark 不能同时用：加密和水印至少要一个")
        return
    if office and not encrypt:
        emit_error(f"{'、'.join(office)} 是 Word / Excel / PPT，只能加密、不能加水印。"
                   "要水印请先让用户另存为 PDF；要加密就去掉 --no-encrypt", office_files=office)
        return
    if not args.grid_col or (encrypt and not args.password_col):
        need = "--grid-col / --password-col" if encrypt else "--grid-col"
        emit_error(f"缺少 {need}。不确定列名时先跑 --list-columns 看表头。")
        return
    if office:
        ensure_deps(need_office=True)

    if args.font:
        if not os.path.exists(args.font):
            emit_error(f"找不到字体文件：{args.font}")
            return
        # v2.8.0 起上游 _find_cjk_font() 已经会扫 Linux 字体目录 + fc-match，
        # 所以 --font 从「非 Windows 必填」降级成纯覆盖手段。仍用 monkeypatch 覆盖，
        # 不改 vendor 代码——下次同步不会把这个补丁冲掉。
        pdf_dist._find_cjk_font = lambda: args.font

    config = build_config(
        pdf_input_paths=[os.path.abspath(p) for p in args.files],
        pdf_mapping_path=os.path.abspath(args.mapping),
        output_path=os.path.abspath(args.output) if args.output else "",
        pdf_grid_column=args.grid_col,
        pdf_password_column=args.password_col if encrypt else "",
        pdf_receiver_column=args.receiver_col,
        pdf_encrypt=encrypt,
        pdf_watermark=watermark,
        pdf_watermark_text=args.watermark_text,
        pdf_watermark_opacity=args.opacity,
        pdf_watermark_angle=args.angle,
    )

    try:
        output_root = run_pdf_dist(config, log_fn=eprint, progress_fn=lambda v: eprint(f"[进度 {int(v * 100)}%]"))
    except Exception as e:
        emit_error(f"分发失败：{e}")
        return

    if not output_root:
        emit_error("分发失败：没有任何网格成功生成（详见上方日志）")
        return

    manifest = os.path.join(output_root, MANIFEST_NAME)
    files = [f for f in list_outputs(output_root, exts=("pdf",) + tuple(e[1:] for e in OFFICE_EXTS))
             if os.path.normcase(f) != os.path.normcase(manifest)]
    grids = len({os.path.dirname(f) for f in files})
    opened = open_folder(output_root) if args.open else False
    has_pdf = ".pdf" in exts
    how = "、".join(x for x, on in (("专属打开密码", encrypt),
                                    ("PDF 专属水印" if office else "专属水印", watermark and has_pdf)) if on)
    msg = (f"已为 {grids} 个网格生成 {len(files)} 份文件（每个网格{how}）。\n"
           f"结果在：{output_root}" + ("（已为你打开这个文件夹）" if opened else ""))
    if office and watermark and has_pdf:
        msg += "\nWord / Excel / PPT 只加了密码，没有水印（要水印请先另存为 PDF）。"
    if encrypt:
        eprint(f"⚠ {MANIFEST_NAME} 含明文密码，请只发给需要核对密码的管理者，不要和加密文件一起群发出去")
        msg += f"\n⚠ 「{MANIFEST_NAME}」里是明文密码，只给核对密码的管理者，不要和文件一起发出去。"
    saved = save_recipe_from(args, config, "pdf", input_name=args.mapping)
    emit({"ok": True, "output_path": output_root, "manifest": manifest,
          "recipe": recipe["name"] if recipe else "", "recipe_notes": recipe_notes, "recipe_saved": saved,
          "encrypted": encrypt, "watermarked": watermark and has_pdf,
          "files_total": len(files), "files": files[:10], "opened": opened, "message": msg})


if __name__ == "__main__":
    main()
