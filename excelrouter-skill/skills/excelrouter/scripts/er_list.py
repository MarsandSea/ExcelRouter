#!/usr/bin/env python
"""
excelrouter · 列出拆分/分发产出，方便把"真正的文件"交给用户。

为什么需要它：拆分产出的真实目录在 JSON 的 output_path 里，但它是一个**目录**。
如果把目录直接丢给 present_files / 附件，很多界面会把它显示成「0 B」的空卡片，
用户以为什么都没产出——其实里面几百个 xlsx 都好好的。这个脚本把目录里的
产出文件扫出来，按 size 排好序，输出成 JSON，让调用方（人或上层 Agent）能
精确地把**具体文件**呈现出来，而不是那个看起来是 0 B 的目录。

用法：
  python scripts/er_list.py --output 拆分结果
  python scripts/er_list.py --output 拆分结果 --limit 10      # 只回前 10 个（用于呈现样例）
  python scripts/er_list.py --output 分发结果 --ext pdf        # PDF 分发场景

输出（stdout 一行 JSON）：
  {"ok": true, "dir": "...", "total": 128, "total_bytes": ...,
   "by_ext": {"xlsx": 128}, "files": ["绝对路径", ...]}   # files 最多 limit 个（不传=全部）
失败（目录不存在等）也是一行 JSON：{"ok": false, "error": "..."}
"""

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import bootstrap, emit, emit_error  # noqa: E402

bootstrap()


def _walk_files(root, ext):
    ext = ext.lower().lstrip(".")
    out = []
    for dirpath, _dirnames, filenames in os.walk(root):
        for fn in filenames:
            low = fn.lower()
            if low.startswith("~$"):          # Excel 临时锁文件，跳过
                continue
            if low.endswith("__tmp__.xlsx"):
                continue
            if ext and not low.endswith("." + ext):
                continue
            full = os.path.join(dirpath, fn)
            try:
                size = os.path.getsize(full)
            except OSError:
                size = 0
            out.append((full, size))
    # 大文件排前面，让人一眼看到"主产出"而不是一堆小附件
    out.sort(key=lambda t: (t[1], t[0]), reverse=True)
    return out


def main():
    ap = argparse.ArgumentParser(description="列出 ExcelRouter 产出目录里的真实文件")
    ap.add_argument("--output", required=True, help="产出目录（er_split/er_pdf_dist 返回的 output_path）")
    ap.add_argument("--ext", default="xlsx", help="只列这种扩展名，留空列全部（默认 xlsx）")
    ap.add_argument("--limit", type=int, default=0, help="files 最多返回几个（0=全部，用于呈现样例）")
    args = ap.parse_args()

    root = os.path.abspath(args.output)
    if not os.path.isdir(root):
        emit_error(f"产出目录不存在或不是目录：{root}")

    files = _walk_files(root, args.ext)
    if args.limit and args.limit > 0:
        shown = files[: args.limit]
    else:
        shown = files

    by_ext = {}
    total_bytes = 0
    for _f, sz in files:
        total_bytes += sz
        _e = os.path.splitext(_f)[1].lower().lstrip(".") or "(无扩展名)"
        by_ext[_e] = by_ext.get(_e, 0) + 1

    emit({
        "ok": True,
        "dir": root,
        "total": len(files),
        "total_bytes": total_bytes,
        "by_ext": by_ext,
        "files": [f for f, _ in shown],
    })


if __name__ == "__main__":
    main()
