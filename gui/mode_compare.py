"""
ExcelRouter - 图形界面：「两表核对」模式。

①原来的表 / 新的表 ②关键列 + 结果卡片（新增 / 没了 / 改过的行）。核对本身在 core/compare.py，
界面文案在 gui/compare_text.py（不 import tkinter，可无头测）。

本模块只是 App 的一块（mixin）：控件、状态都挂在同一个 App 实例上，
方法按名字互相调用，和写在 gui/app.py 里时完全一样。

Copyright (c) 2026 AbeLin
MIT License
"""
# App 的属性分散在几个 mixin 里，pyright 在本文件里看不到别处定义的 self.xxx
# pyright: reportAttributeAccessIssue=false

import os
import threading
from typing import Any

import customtkinter as ctk
from tkinter import filedialog, messagebox

from gui import compare_text as CT
from gui.platform_compat import open_folder
from gui.common import (
    CMP_IN_TIP, CMP_KEY_PLACEHOLDER, CMP_MAX_CELLS, CMP_COLORS, C_OK, C_WARN, C_ERR, C_MUTED,
    PRIMARY, PRIMARY_HOVER, OK_BANNER_BG, CARD_BORDER, _ghost_button, _flat_button, save_config,
)


class CompareModeMixin:
    # ① 选原来的表和新的表 → ② 确认按哪一列对上每一行 → ③ 开始核对；结果固定在新表旁的「核对结果」
    def _build_compare_steps(self, parent):
        card, _ = self._step_card(parent, 0, 1, "选两张要核对的表")
        self._cmp_a_var = ctk.StringVar(value=self.cfg.get("compare_a", ""))
        self._cmp_b_var = ctk.StringVar(value=self.cfg.get("compare_b", ""))
        for i, (label, var, side) in enumerate((("原来的表（A）", self._cmp_a_var, "a"),
                                                 ("新的表（B）", self._cmp_b_var, "b"))):
            row = ctk.CTkFrame(card, fg_color="transparent")
            row.grid(row=1 + i, column=0, padx=12, pady=(6 if i == 0 else 4, 0), sticky="ew")
            row.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(row, text=label, width=96, anchor="w").grid(row=0, column=0, sticky="w")
            entry = ctk.CTkEntry(row, textvariable=var, placeholder_text="选一个 Excel 文件，或拖到这里")
            entry.grid(row=0, column=1, padx=8, sticky="ew")
            entry.bind("<Return>", lambda e: self._scan_compare())
            entry.bind("<FocusOut>", lambda e: self._scan_compare())
            self._enable_drop(entry, lambda e, s=side: self._on_drop_compare(e, s))
            _ghost_button(row, text="📄 选择", width=70,
                          command=lambda s=side: self._pick_compare(s)).grid(row=0, column=2)
        self._enable_drop(card, lambda e: self._on_drop_compare(e, None))
        self._cmp_in_status = ctk.CTkLabel(card, text=CMP_IN_TIP, font=ctk.CTkFont(size=11), text_color=C_MUTED,
                                           anchor="w", justify="left", wraplength=640)
        self._cmp_in_status.grid(row=3, column=0, padx=12, pady=(4, 10), sticky="w")

        card2, _ = self._step_card(parent, 1, 2, "按哪一列对上每一行")
        krow = ctk.CTkFrame(card2, fg_color="transparent")
        krow.grid(row=1, column=0, padx=12, pady=(6, 0), sticky="w")
        ctk.CTkLabel(krow, text="关键列").pack(side="left")
        self._cmp_key_var = ctk.StringVar(value=CMP_KEY_PLACEHOLDER)
        self._cmp_key_menu = ctk.CTkOptionMenu(krow, variable=self._cmp_key_var, values=[CMP_KEY_PLACEHOLDER],
                                               width=320, dynamic_resizing=False)
        self._cmp_key_menu.pack(side="left", padx=(8, 0))
        _flat_button(krow, text="🔄 重新识别", width=90,
                     command=lambda: self._scan_compare(force=True)).pack(side="left", padx=(8, 0))
        ctk.CTkLabel(card2, text="关键列是每行都不一样、两张表里都有的列，比如工号、编号、身份证号。"
                                 "数字按数值比（100 和 100.00 一样），文字不分全角半角、不看首尾空格；"
                                 "关键列重复或空着的行不配对，单独列出来，免得配错。",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED, anchor="w", justify="left",
                     wraplength=640).grid(row=2, column=0, padx=12, pady=(4, 10), sticky="w")

        res = self._cmp_result_card = ctk.CTkFrame(parent, corner_radius=12, border_width=1,
                                                   border_color=CARD_BORDER)
        res.grid(row=2, column=0, padx=8, pady=(0, 10), sticky="ew")
        res.grid_columnconfigure(0, weight=1)
        self._cmp_result_title = ctk.CTkLabel(res, text="", font=ctk.CTkFont(size=14, weight="bold"))
        self._cmp_result_title.grid(row=0, column=0, padx=14, pady=(10, 2), sticky="w")
        self._cmp_lines_box = ctk.CTkFrame(res, fg_color="transparent")
        self._cmp_lines_box.grid(row=1, column=0, padx=14, pady=(2, 6), sticky="ew")
        rb = ctk.CTkFrame(res, fg_color="transparent")
        rb.grid(row=2, column=0, padx=12, pady=(0, 12), sticky="w")
        ctk.CTkButton(rb, text="📄 打开核对结果", width=130, fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
                      command=self._open_compare_result).pack(side="left")
        ctk.CTkLabel(rb, text="「有变化」逐格列出原值、新值和差额，「变化行」把改过的格子标红",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED).pack(side="left", padx=(10, 0))
        res.grid_remove()

    def _pick_compare(self, side):
        p = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls")])
        if p:
            (self._cmp_a_var if side == "a" else self._cmp_b_var).set(os.path.normpath(p))
            self._scan_compare()

    def _on_drop_compare(self, event, side):
        """拖进来两个表：先拖的当原来的表；拖一个到某一行就填那一行，拖到卡片上就填还空着的那一行。"""
        paths = [p for p in self._parse_drop_paths(getattr(event, "data", ""))
                 if p.lower().endswith((".xlsx", ".xlsm", ".xls"))]
        if not paths:
            self._cmp_in_status.configure(text="⚠ 请拖入 Excel 文件（.xlsx / .xls）", text_color=C_WARN)
            return
        if len(paths) >= 2:
            self._cmp_a_var.set(os.path.normpath(paths[0]))
            self._cmp_b_var.set(os.path.normpath(paths[1]))
        else:
            if side is None:
                side = "a" if not self._cmp_a_var.get().strip() else "b"
            (self._cmp_a_var if side == "a" else self._cmp_b_var).set(os.path.normpath(paths[0]))
        self._scan_compare()

    def _cmp_header_row(self):
        h = self._header_cfg()
        return h["header_row"] if h["header_mode"] == "row" else 0

    def _scan_compare(self, force=False):
        """两张表都选好后在子线程读一遍：两表都有的列、推荐关键列（只回列名和计数）。"""
        a, b = self._cmp_a_var.get().strip(), self._cmp_b_var.get().strip()
        if not a or not b:
            if a or b:
                self._cmp_in_status.configure(text="再选另一张表", text_color=C_MUTED)
            return
        if (a, b) == self._cmp_scanned and not force:
            return
        self._cmp_scanned = (a, b)
        self._cmp_result_card.grid_remove()
        missing = [p for p in (a, b) if not os.path.isfile(p)]
        if missing:
            self._cmp_info = {}
            self._cmp_in_status.configure(text=f"⚠ 找不到文件：{os.path.basename(missing[0])}", text_color=C_WARN)
            return
        if os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b)):
            self._cmp_info = {}
            self._cmp_in_status.configure(text="⚠ 两次选的是同一个文件，请选两张不同的表", text_color=C_WARN)
            return
        self._cmp_in_status.configure(text="正在读两张表、找能对上的列…", text_color=C_MUTED)
        hr = self._cmp_header_row()

        def work():
            info: dict[str, Any] = {"a": a, "b": b}
            try:
                from core import compare
                info.update(compare.scan_pair(a, b, header_row=hr, max_cells=CMP_MAX_CELLS))
            except Exception as e:
                info["error"] = str(e)
            self._ui_q.put(("compare_scan", info))

        threading.Thread(target=work, daemon=True).start()

    def _on_compare_scan(self, info):
        if (info["a"], info["b"]) != (self._cmp_a_var.get().strip(), self._cmp_b_var.get().strip()):
            return    # 结果已过期
        self._cmp_info = info
        if info.get("error"):
            self._cmp_set_keys([], {}, CMP_KEY_PLACEHOLDER)
            self._cmp_in_status.configure(text=f"⚠ 读不了：{info['error']}", text_color=C_WARN)
            return
        saved = self.cfg.get("compare_key") or []
        labels, mapping, default = CT.key_options(info, saved if isinstance(saved, list) else [])
        self._cmp_set_keys(labels, mapping, default or CMP_KEY_PLACEHOLDER)
        text = (f"✓ 原来的表 {info['rows_a']} 行（{info['sheet_a']}），新的表 {info['rows_b']} 行（{info['sheet_b']}），"
                f"两表都有的列 {len(info['columns'])} 个")
        color = C_OK
        if not info.get("keys"):
            text += "；⚠ 没找到每行都不一样、两边能对上的列，请在 ② 里自己选关键列"
            color = C_WARN
        self._cmp_in_status.configure(text=text, text_color=color)

    def _cmp_set_keys(self, labels, mapping, current):
        self._cmp_keymap = mapping
        self._cmp_key_menu.configure(values=labels or [CMP_KEY_PLACEHOLDER])
        self._cmp_key_var.set(current)

    def _cmp_key_cols(self):
        """关键列下拉当前选的列名列表；还没识别时沿用上次存的。"""
        cols = self._cmp_keymap.get(self._cmp_key_var.get())
        if cols:
            return list(cols)
        saved = self.cfg.get("compare_key") or []
        return list(saved) if isinstance(saved, list) else []

    def _start_compare(self):
        a, b = self._cmp_a_var.get().strip(), self._cmp_b_var.get().strip()
        if not a or not b or not os.path.isfile(a) or not os.path.isfile(b):
            messagebox.showwarning("先选两张表", "请在第 ① 步选好「原来的表」和「新的表」。")
            return
        if os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b)):
            messagebox.showwarning("选了同一个文件", "两次选的是同一个文件，请选两张不同的表。")
            return
        key = self._cmp_keymap.get(self._cmp_key_var.get())
        if not key:
            messagebox.showwarning("先选关键列", "请在第 ② 步选「关键列」——按哪一列把两张表的行一一对上"
                                               "（工号、编号这类每行都不一样的列）。")
            return
        locks = self._find_open_locks(a) + self._find_open_locks(b)
        if locks and not messagebox.askyesno(
                "文件正被打开",
                f"检测到 {'、'.join(locks)} 正被 Excel/WPS 打开：读到的是「打开前保存的内容」，刚改的不会算进去。\n\n"
                "建议先关闭。仍要继续吗？"):
            return
        cfg = self._collect_config()
        run_cfg = {"compare_a": a, "compare_b": b, "compare_key": list(key), "max_cells": CMP_MAX_CELLS,
                   **self._header_cfg()}
        self._enter_running("⏳ 正在核对…", "已开始，正在读取两张表…")
        self._cmp_result_card.grid_remove()

        def run():
            try:
                from core.compare import run_compare
                output_path = run_compare(run_cfg,
                                          log_fn=lambda m: self._ui_q.put(("log", m)),
                                          progress_fn=lambda v: self._ui_q.put(("progress", v)),
                                          stop_flag=lambda: self._stop_flag)
            except Exception as e:
                self._ui_q.put(("log", f"\n❌ 运行出错：{e}"))
                output_path = None
            self._ui_q.put(("done", (cfg, output_path)))

        threading.Thread(target=run, daemon=True).start()

    def _on_compare_done(self, cfg, output_path):
        s = self._last_summary or {}
        if not output_path or s.get("mode") != "compare":
            if self._stop_flag:
                self._run_status.configure(text="⏹ 已停止，没有写核对结果", text_color=C_WARN)
                return
            self._run_status.configure(
                text="❌ 没有完成：原因见下方「处理详情」；可点「复制日志」连同反馈一起发给作者", text_color=C_ERR)
            self._copylog_btn.grid()
            self._toggle_log(show=True)
            return
        self._last_output = output_path
        self._cmp_result = s.get("result", "")
        self._open_btn.grid()
        try:
            save_config(cfg)
        except Exception:
            pass
        for w in self._cmp_lines_box.winfo_children():
            w.destroy()
        for i, (text, kind) in enumerate(CT.result_lines(s)):
            ctk.CTkLabel(self._cmp_lines_box, text=text, anchor="w", justify="left", wraplength=620, height=22,
                         font=ctk.CTkFont(size=12, weight="bold" if not text.startswith(" ") else "normal"),
                         text_color=CMP_COLORS.get(kind, C_MUTED)).grid(row=i, column=0, sticky="w")
        same = not (s.get("only_a") or s.get("only_b") or s.get("changed_rows"))
        self._cmp_result_title.configure(text="✅ 两张表一样" if same else "🔍 核对完成：这些地方不一样")
        self._cmp_result_card.grid()
        text, warn = CT.banner_lines(s, output_path)
        self._summary_lbl.configure(text=text, text_color=C_WARN if warn else C_OK)
        self._summary_banner.configure(fg_color=("gray90", "gray20") if warn else OK_BANNER_BG)
        self._show_summary_banner()
        self._run_status.configure(text="✅ 核对完成！点「打开核对结果」看逐格对比", text_color=C_OK)
        self._log("💬 用得顺手或踩了坑？点「反馈建议」匿名告诉作者（1 分钟）。")

    def _open_compare_result(self):
        p = self._cmp_result
        if not p or not os.path.exists(p):
            self._run_status.configure(text="⚠ 核对结果不在了，请重新核对一次", text_color=C_WARN)
            return
        if not open_folder(p):
            self._run_status.configure(text=f"⚠ 打不开核对结果，文件在：{p}", text_color=C_WARN)
