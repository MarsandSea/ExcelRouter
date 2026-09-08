#!/usr/bin/env python
"""
excelrouter · PDF 按网格加密分发：同一批 PDF，按「网格→密码」映射清单，
为每个网格生成专属打开密码 + 专属水印（可溯源）的副本，并输出含明文密码的分发清单。

Copyright (c) 2026 AbeLin
MIT License

用法：
  # 起步：手上还没有映射清单时，先生成一份模板给用户填（网格/密码/接收人 + 填写说明）
  python er_pdf_dist.py --mapping 映射.xlsx --make-template 映射.xlsx

  # 第一步：不确定映射清单的列名时，先看一眼表头
  python er_pdf_dist.py --mapping 映射.xlsx --list-columns

  # 可选：懒得想密码时，让工具给空白网格生成随机密码（写入新文件，不动原清单）
  python er_pdf_dist.py --mapping 映射.xlsx --grid-col 网格 --random-password

  # 第二步：确认列名后正式分发
  python er_pdf_dist.py --pdf a.pdf b.pdf --mapping 映射.xlsx \
      --grid-col 网格 --password-col 密码 --receiver-col 接收人 --output 分发结果

输出：stdout 一行 JSON。分发清单「分发清单.xlsx」含明文密码，务必提醒用户
不要把这份清单和加密后的 PDF 一起发给同一批人。

--random-password 会用「原文件名_含密码.xlsx」另存一份新清单，正式分发时改用
那份新清单（JSON 里的 mapping_with_passwords 就是它的绝对路径）。
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import bootstrap, build_config, check_deps, emit, emit_error, eprint  # noqa: E402

bootstrap()

import core.pdf_dist as pdf_dist  # noqa: E402
from core.pdf_dist import (  # noqa: E402
    MANIFEST_NAME,
    fill_random_passwords,
    list_mapping_columns,
    run_pdf_dist,
    write_mapping_template,
)


def main():
    ap = argparse.ArgumentParser(
        description="PDF 按网格加密分发（专属密码 + 专属水印 + 分发清单）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    ap.add_argument("--pdf", nargs="+", default=[], help="待分发的 PDF 文件，可多个")
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
    ap.add_argument("--password-col", default="", help="映射清单里的密码列名")
    ap.add_argument("--receiver-col", default="", help="映射清单里的接收人列名（选填）")
    ap.add_argument("--output", default="", help="输出目录；不填则用「第一个PDF所在目录/分发结果」")
    ap.add_argument("--no-watermark", action="store_true", help="不加水印")
    ap.add_argument("--watermark-text", default="{grid} {date}", help="水印模板，支持 {grid} {date} 占位")
    ap.add_argument("--opacity", type=float, default=0.15, help="水印透明度 0~1")
    ap.add_argument("--angle", type=float, default=45, help="水印旋转角度")
    ap.add_argument("--font", default="", help="中文字体 .ttf 路径（非 Windows 环境找不到系统字体时用）")
    args = ap.parse_args()

    # ---- 起步：生成映射清单模板（唯一不需要已有文件的分支，所以排在所有存在性检查之前）----
    if args.make_template:
        check_deps()
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
        check_deps()
        if not args.grid_col:
            emit_error("--random-password 需要 --grid-col 指出哪一列是网格"
                       "（不确定列名就先跑 --list-columns 看表头）")
            return
        # 【上游 bug 的包装层规避，不改 vendor】core.pdf_dist.fill_random_passwords 在
        # password_col 传空串时，即便清单里本就有「密码」列，也会走到 `p_idx = len(header)`
        # 分支，在表尾再追加一个同名「密码」列，产物出现两个密码列。上游 gui/app.py 正是
        # 这样无参调用的（桌面版同样会中招）。这里在调用前先把默认值解析成真实列名，
        # 让上游函数走「复用已有列」的分支。上游修好后本段可安全删除。
        effective_pwd_col = args.password_col
        if not effective_pwd_col:
            if "密码" in list_mapping_columns(os.path.abspath(args.mapping)):
                effective_pwd_col = "密码"
        try:
            out_path, count = fill_random_passwords(
                os.path.abspath(args.mapping), args.grid_col, effective_pwd_col)
        except Exception as e:
            emit_error(f"生成随机密码失败：{e}")
            return
        # 密码列可能原本不存在而被新建成「密码」，以产物实际表头为准，别猜
        out_cols = list_mapping_columns(out_path)
        pwd_col = args.password_col if args.password_col in out_cols else (
            "密码" if "密码" in out_cols else (out_cols[-1] if out_cols else "密码"))
        eprint(f"✅ 已为 {count} 个网格生成随机密码，新清单：{out_path}（原清单未改动）")
        eprint("⚠ 新清单含明文密码，只留给分发人自己用，不要随 PDF 一起发出去。")
        emit({"ok": True, "mapping_with_passwords": out_path, "generated": count,
              "password_col": pwd_col,
              "next": f'正式分发改用这份新清单：--mapping "{out_path}" '
                      f'--grid-col {args.grid_col} --password-col {pwd_col}'})
        return

    if args.list_columns:
        check_deps()
        try:
            cols = list_mapping_columns(args.mapping)
        except Exception as e:
            emit_error(f"读取映射清单表头失败：{e}")
            return
        emit({"ok": True, "columns": cols})
        return

    check_deps(need_pdf=True)

    if not args.pdf:
        emit_error("缺少 --pdf（至少一个待分发的 PDF 文件），或改用 --list-columns 只看映射清单表头")
        return
    for p in args.pdf:
        if not os.path.exists(p):
            emit_error(f"找不到 PDF：{p}")
            return
    if not args.grid_col or not args.password_col:
        emit_error("缺少 --grid-col / --password-col。不确定列名时先跑 --list-columns 看表头。")
        return

    if args.font:
        if not os.path.exists(args.font):
            emit_error(f"找不到字体文件：{args.font}")
            return
        # 上游 _find_cjk_font() 只扫 Windows 系统字体目录，非 Windows 环境（或想指定字体时）
        # 用 monkeypatch 覆盖，不改 vendor 代码——下次同步不会把这个补丁冲掉。
        pdf_dist._find_cjk_font = lambda: args.font

    config = build_config(
        pdf_input_paths=[os.path.abspath(p) for p in args.pdf],
        pdf_mapping_path=os.path.abspath(args.mapping),
        output_path=os.path.abspath(args.output) if args.output else "",
        pdf_grid_column=args.grid_col,
        pdf_password_column=args.password_col,
        pdf_receiver_column=args.receiver_col,
        pdf_watermark=not args.no_watermark,
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

    eprint(f"⚠ {MANIFEST_NAME} 含明文密码，请只发给需要核对密码的管理者，不要和加密 PDF 一起群发出去")
    emit({"ok": True, "output_path": output_root, "manifest": os.path.join(output_root, MANIFEST_NAME)})


if __name__ == "__main__":
    main()
