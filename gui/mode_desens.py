"""
ExcelRouter - 图形界面：「数据脱敏」模式。

①表格或文件夹 ②用途 + 逐列处理方式 + 结果卡片，以及汇总结果「另存脱敏版」。

本模块只是 App 的一块（mixin）：控件、状态都挂在同一个 App 实例上，
方法按名字互相调用，和写在 gui/app.py 里时完全一样。

Copyright (c) 2026 AbeLin
MIT License
"""
# App 的属性分散在几个 mixin 里，pyright 在本文件里看不到别处定义的 self.xxx
# pyright: reportAttributeAccessIssue=false

import os
import re
import threading
from typing import Any

import customtkinter as ctk
from tkinter import filedialog, messagebox

from gui.platform_compat import open_folder
from gui.common import (
    DS_IN_TIP, DS_SCENES, DS_SCENE_TEXT, DS_METHODS, DS_METHOD_WORDS, C_OK, C_WARN, C_ERR,
    C_MUTED, PRIMARY, PRIMARY_HOVER, OK_BANNER_BG, CARD_BORDER, _ghost_button, save_config,
)


class DesensModeMixin:
    # ── 数据脱敏（v3.1）─────────────────────────────────────
    # ① 选表格 → ② 选用途、看识别结果、逐列调整 → ③ 开始脱敏；结果固定在输入旁的「脱敏结果」
    def _build_desens_steps(self, parent):
        card, _ = self._step_card(parent, 0, 1, "选择要脱敏的表格")
        btns = ctk.CTkFrame(card, fg_color="transparent")
        btns.grid(row=1, column=0, padx=12, pady=(6, 2), sticky="w")
        ctk.CTkButton(btns, text="📄 选一个 Excel 文件", width=170,
                      command=self._pick_desens_file).pack(side="left")
        _ghost_button(btns, text="📁 选整个文件夹", width=130,
                      command=self._pick_desens_folder).pack(side="left", padx=(10, 0))
        self._ds_in_var = ctk.StringVar(value=self.cfg.get("desens_input_path", ""))
        entry = ctk.CTkEntry(card, textvariable=self._ds_in_var,
                             placeholder_text="也可以把表格或文件夹直接拖到这里")
        entry.grid(row=2, column=0, padx=12, pady=(6, 2), sticky="ew")
        entry.bind("<Return>",   lambda e: self._scan_desens())
        entry.bind("<FocusOut>", lambda e: self._scan_desens())
        self._enable_drop(card, self._on_drop_desens)
        self._enable_drop(entry, self._on_drop_desens)
        self._ds_in_status = ctk.CTkLabel(
            card, text=DS_IN_TIP, font=ctk.CTkFont(size=11), text_color=C_MUTED,
            anchor="w", justify="left", wraplength=640)
        self._ds_in_status.grid(row=3, column=0, padx=12, pady=(0, 10), sticky="w")

        card2, _ = self._step_card(parent, 1, 2, "处理完拿去做什么")
        scene = self.cfg.get("desens_scene") or DS_SCENES[0]
        self._ds_scene_seg = ctk.CTkSegmentedButton(card2, values=list(DS_SCENES), command=self._on_desens_scene,
                                                    selected_color=PRIMARY, selected_hover_color=PRIMARY_HOVER)
        self._ds_scene_seg.set(scene if scene in DS_SCENES else DS_SCENES[0])
        self._ds_scene_seg.grid(row=1, column=0, padx=12, pady=(6, 0), sticky="w")
        self._ds_scene_tip = ctk.CTkLabel(card2, text=DS_SCENE_TEXT[self._ds_scene_seg.get()],
                                          font=ctk.CTkFont(size=11), text_color=C_MUTED, anchor="w",
                                          justify="left", wraplength=640)
        self._ds_scene_tip.grid(row=2, column=0, padx=12, pady=(2, 4), sticky="w")
        self._ds_notes_var = ctk.BooleanVar(value=self.cfg.get("desens_notes", True) is not False)
        ctk.CTkCheckBox(card2, text="备注等文字里夹带的手机号、证件号、银行卡号也打码",
                        variable=self._ds_notes_var).grid(row=3, column=0, padx=12, pady=(2, 6), sticky="w")
        self._ds_cols_title = ctk.CTkLabel(card2, text="识别结果：选好表格后这里列出认出的列，处理方式可以逐列改",
                                           font=ctk.CTkFont(size=12, weight="bold"), anchor="w")
        self._ds_cols_title.grid(row=4, column=0, padx=12, pady=(4, 0), sticky="w")
        self._ds_cols_box = ctk.CTkFrame(card2, fg_color="transparent")
        self._ds_cols_box.grid(row=5, column=0, padx=12, pady=(2, 2), sticky="ew")
        xrow = ctk.CTkFrame(card2, fg_color="transparent")
        xrow.grid(row=6, column=0, padx=12, pady=(4, 0), sticky="ew")
        ctk.CTkLabel(xrow, text="其他列（选填）").pack(side="left")
        self._ds_extra_var = ctk.StringVar(value=str(self.cfg.get("desens_columns", "") or ""))
        ctk.CTkEntry(xrow, textvariable=self._ds_extra_var, width=360,
                     placeholder_text="如 备注:清空、宽带账号:换代号").pack(side="left", padx=(8, 0))
        ctk.CTkLabel(card2, text="识别靠列名和号码校验，不保证找全：没认出的列照发，漏了的在这里写「列名:处理方式」"
                                 "（打码 / 换代号 / 泛化 / 清空 / 照发）。处理完会把结果重新扫一遍，还有完整号码的表不写出来。",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED, anchor="w", justify="left",
                     wraplength=640).grid(row=7, column=0, padx=12, pady=(2, 10), sticky="w")

        res = self._ds_result = ctk.CTkFrame(parent, corner_radius=12, border_width=1, border_color=CARD_BORDER)
        res.grid(row=2, column=0, padx=8, pady=(0, 10), sticky="ew")
        res.grid_columnconfigure(0, weight=1)
        self._ds_result_title = ctk.CTkLabel(res, text="", font=ctk.CTkFont(size=14, weight="bold"))
        self._ds_result_title.grid(row=0, column=0, padx=14, pady=(10, 2), sticky="w")
        self._ds_result_lbl = ctk.CTkLabel(res, text="", font=ctk.CTkFont(size=12), anchor="w", justify="left",
                                           wraplength=640)
        self._ds_result_lbl.grid(row=1, column=0, padx=14, pady=(2, 6), sticky="w")
        rb = ctk.CTkFrame(res, fg_color="transparent")
        rb.grid(row=2, column=0, padx=12, pady=(0, 12), sticky="w")
        ctk.CTkButton(rb, text="📄 打开脱敏报告", width=130, fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
                      command=self._open_desens_report).pack(side="left")
        res.grid_remove()
        p = self._ds_in_var.get().strip()
        if p and os.path.exists(p):
            self._scan_desens()

    def _pick_desens_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls")])
        if p:
            self._ds_in_var.set(os.path.normpath(p))
            self._scan_desens()

    def _pick_desens_folder(self):
        p = filedialog.askdirectory()
        if p:
            self._ds_in_var.set(os.path.normpath(p))
            self._scan_desens()

    def _on_drop_desens(self, event):
        paths = self._parse_drop_paths(getattr(event, "data", ""))
        if self._drop_recipe(paths):
            return
        target = next((p for p in paths
                       if os.path.isdir(p) or p.lower().endswith((".xlsx", ".xlsm", ".xls"))), None)
        if target is None:
            self._ds_in_status.configure(text="⚠ 请拖入一个 Excel 文件，或装表格的文件夹", text_color=C_WARN)
            return
        self._ds_in_var.set(os.path.normpath(target))
        self._scan_desens()

    def _scan_desens(self, force=False):
        """子线程识别：数一数有几个表，拿第一个表当样本认出哪些列是个人信息（只有列名和类型，没有值）。"""
        p = self._ds_in_var.get().strip()
        if not p or (p == self._ds_scanned and not force):
            return
        self._ds_scanned = p
        if not os.path.exists(p):
            self._ds_info = {}
            self._ds_in_status.configure(text="⚠ 找不到这个文件或文件夹，请检查路径", text_color=C_WARN)
            return
        self._ds_in_status.configure(text="正在识别…", text_color=C_MUTED)
        setting = self._desens_setting(with_overrides=False)
        cfg = self._header_cfg()

        def work():
            info: dict[str, Any] = {"path": p}
            try:
                from core import desens
                tasks = desens._tasks(p)
                info["n"] = len(tasks)
                if tasks:
                    sample = tasks[0][0]
                    info["sample"] = os.path.basename(sample)
                    info["sheets"] = desens.inspect(sample, setting, cfg)
            except Exception as e:
                info["error"] = str(e)
            self._ui_q.put(("desens_scan", info))

        threading.Thread(target=work, daemon=True).start()

    def _header_cfg(self):
        mode_map = {"自动": "auto", "指定行号": "row", "关键词": "keyword"}
        try:
            row = int(self._hrow_var.get())
        except (TypeError, ValueError):
            row = 1
        return {"header_mode": mode_map.get(self._header_seg.get(), "auto"), "header_row": row,
                "grid_keys": self._split_list(self._grid_keys_var.get()),
                "id_keys": self._split_list(self._id_keys_var.get())}

    def _on_desens_scan(self, info):
        if info["path"] != self._ds_in_var.get().strip():
            return    # 结果已过期
        self._ds_info = info
        if info.get("error"):
            self._ds_in_status.configure(text=f"⚠ 读不了这个位置：{info['error']}", text_color=C_WARN)
            self._render_desens_cols()
            return
        n = info.get("n", 0)
        if not n:
            self._ds_in_status.configure(text="⚠ 这里没有表格（.xlsx / .xls）", text_color=C_WARN)
        elif os.path.isfile(info["path"]):
            self._ds_in_status.configure(text=f"✓ 已选择：{os.path.basename(info['path'])}；结果放在旁边的「脱敏结果」，原表不改",
                                         text_color=C_OK)
        else:
            self._ds_in_status.configure(text=f"✓ 找到 {n} 个表，全部一起处理；结果放在这个文件夹里的「脱敏结果」，原表不改",
                                         text_color=C_OK)
        self._render_desens_cols()

    def _ds_detected(self):
        """样本表里认出的列：[(列名, 类型)]，按出现顺序去重；隐藏的 sheet 不算（结果里本来就不带）。"""
        seen, out = set(), []
        for s in self._ds_info.get("sheets") or []:
            if s.get("hidden") and not self._keep_hidden_var.get():
                continue
            for c in s["columns"]:
                if c["column"] not in seen:
                    seen.add(c["column"])
                    out.append((c["column"], c["kind"]))
        return out

    def _render_desens_cols(self):
        """识别结果一列一行：列名 · 识别为 · 处理方式下拉（默认值随用途变；改过的记进 _ds_overrides）。"""
        for w in self._ds_cols_box.winfo_children():
            w.destroy()
        cols = self._ds_detected()
        info = self._ds_info
        if not cols:
            self._ds_cols_title.configure(
                text="识别结果：没认出个人信息列（可以在下面「其他列」里自己写）" if info.get("sheets") is not None
                else "识别结果：选好表格后这里列出认出的列，处理方式可以逐列改")
            return
        from core.desens import _scene_method
        scene = self._ds_scene_seg.get()
        more = f"（样本：{info.get('sample', '')}，其他表处理时各自重新识别）" if info.get("n", 0) > 1 else ""
        self._ds_cols_title.configure(text=f"识别出 {len(cols)} 列{more}")
        for i, (col, kind) in enumerate(cols[:30]):
            ctk.CTkLabel(self._ds_cols_box, text=col, anchor="w", width=150).grid(row=i, column=0, sticky="w")
            ctk.CTkLabel(self._ds_cols_box, text=kind, anchor="w", width=80, text_color=C_MUTED).grid(
                row=i, column=1, sticky="w", padx=(6, 0))
            method = self._ds_overrides.get(col) or _scene_method(scene, kind)
            var = ctk.StringVar(value=DS_METHODS[method])
            ctk.CTkOptionMenu(self._ds_cols_box, variable=var, values=list(DS_METHODS.values()), width=110,
                              command=lambda label, c=col, k=kind: self._on_desens_method(c, k, label)).grid(
                row=i, column=2, sticky="w", padx=(6, 0), pady=1)
        if len(cols) > 30:
            ctk.CTkLabel(self._ds_cols_box, text=f"……还有 {len(cols) - 30} 列按用途默认处理，报告里逐列写明",
                         text_color=C_MUTED, font=ctk.CTkFont(size=11)).grid(row=31, column=0, columnspan=3, sticky="w")

    def _on_desens_method(self, col, kind, label):
        from core.desens import _scene_method
        method = next(k for k, v in DS_METHODS.items() if v == label)
        if method == _scene_method(self._ds_scene_seg.get(), kind):
            self._ds_overrides.pop(col, None)      # 改回和用途一样，就不算单独调整
        else:
            self._ds_overrides[col] = method

    def _on_desens_scene(self, label):
        self._ds_scene_tip.configure(text=DS_SCENE_TEXT.get(label, ""))
        self._ds_overrides = {}                    # 换了用途：逐列调整按新用途重来，免得留着旧用途的选择
        self._render_desens_cols()
        self._ds_result.grid_remove()

    def _parse_desens_extra(self):
        """「备注:清空、宽带账号:换代号」→ ({列名: 处理方式}, 看不懂的几段)。"""
        out, bad = {}, []
        for part in [x.strip() for x in re.split(r"[、，,；;\n]+", self._ds_extra_var.get()) if x.strip()]:
            m = re.match(r"^(.+?)\s*[:：=]\s*(.+)$", part)
            how = DS_METHOD_WORDS.get(m.group(2).strip()) if m else None
            if m is None or how is None:
                bad.append(part)
            else:
                out[m.group(1).strip()] = how
        return out, bad

    def _desens_setting(self, with_overrides=True):
        cols = {}
        if with_overrides:
            extra, _bad = self._parse_desens_extra()
            cols = dict(extra, **self._ds_overrides)
        return {"scene": self._ds_scene_seg.get(), "columns": cols, "notes": self._ds_notes_var.get()}

    def _start_desens(self):
        p = self._ds_in_var.get().strip()
        if not p or not os.path.exists(p):
            messagebox.showwarning("先选择表格", "请先在第 ① 步选择要脱敏的 Excel 文件或文件夹。")
            return
        _extra, bad = self._parse_desens_extra()
        if bad:
            messagebox.showwarning("「其他列」看不懂",
                                   f"这几段看不懂：{'、'.join(bad[:3])}\n\n写法是「列名:处理方式」，处理方式只能是"
                                   "打码、换代号、泛化、清空、照发，多个用顿号隔开。")
            return
        locks = self._find_open_locks(p)
        if locks:
            shown = "、".join(locks[:3]) + (" 等" if len(locks) > 3 else "")
            if not messagebox.askyesno(
                    "文件正被打开",
                    f"检测到以下文件正被 Excel/WPS 打开：{shown}\n"
                    "读到的是「打开前保存的内容」，刚改的不会算进去。\n\n建议先关闭这些文件。仍要继续吗？"):
                return
        cfg = self._collect_config()
        run_cfg = dict(self._header_cfg(), input_path=p, desens=self._desens_setting(),
                       keep_hidden=self._keep_hidden_var.get())
        self._enter_running("⏳ 正在脱敏…", "已开始，正在读取表格…")
        self._ds_result.grid_remove()

        def run():
            try:
                from core.desens import run_desens
                output_path = run_desens(run_cfg,
                                         log_fn=lambda m: self._ui_q.put(("log", m)),
                                         progress_fn=lambda v: self._ui_q.put(("progress", v)),
                                         stop_flag=lambda: self._stop_flag)
            except Exception as e:
                self._ui_q.put(("log", f"\n❌ 运行出错：{e}"))
                output_path = None
            self._ui_q.put(("done", (cfg, output_path)))

        threading.Thread(target=run, daemon=True).start()

    def _on_desens_done(self, cfg, output_path):
        s = self._last_summary or {}
        if not output_path or s.get("mode") != "desens":
            self._run_status.configure(
                text="❌ 没有完成：原因见下方「处理详情」；可点「复制日志」连同反馈一起发给作者", text_color=C_ERR)
            self._copylog_btn.grid()
            self._toggle_log(show=True)
            return
        self._last_output = output_path
        self._ds_report = s.get("report", "")
        self._open_btn.grid()
        try:
            save_config(cfg)
        except Exception:
            pass
        ok, blocked = s.get("ok", False), s.get("blocked", 0)
        lines = [f"🛡 {s.get('files', 0)} 个表写进了结果 · 处理 {s.get('columns', 0)} 列、{s.get('cells', 0)} 个格子"
                 + (f" · 文字里夹带的号码 {s['embedded']} 处" if s.get("embedded") else "")]
        if s.get("stopped"):
            lines.append("⏹ 中途停了：处理完的表已写进结果，没处理的不在里面")
        if ok:
            lines.append("✅ 复检通过：处理过的类型里没有发现完整的号码")
        else:
            lines.append(f"❌ {blocked} 个表没通过复检（还有完整的号码），没有写出来；"
                         "结果文件夹里的「未通过复检_请勿发出.txt」写了是哪列，在「其他列」里指定后再跑一次")
        if s.get("failed_files"):
            lines.append(f"⚠ {s['failed_files']} 个表没处理成（点「处理详情」看原因）")
        lines.append(f"📁 {output_path}")
        warn = not ok or bool(s.get("failed_files")) or bool(s.get("stopped"))
        self._ds_result_title.configure(text="✅ 脱敏完成" if not warn else "⚠ 脱敏完成，但有表没写出来")
        self._ds_result_lbl.configure(
            text="「脱敏报告.xlsx」逐列写了识别为什么、怎么处理、处理了几格（示例是虚构的）；"
                 "批注、图片、图表、文档属性没有带进结果。打码不可逆，原表请自己留好。")
        self._ds_result.grid()
        self._summary_lbl.configure(text="\n".join(lines), text_color=C_WARN if warn else C_OK)
        self._summary_banner.configure(fg_color=("gray90", "gray20") if warn else OK_BANNER_BG)
        self._show_summary_banner()
        self._run_status.configure(text="✅ 脱敏完成！点「打开输出文件夹」查看" if not warn else
                                   "⚠ 有表没通过复检，没有写出来，详见下方", text_color=C_WARN if warn else C_OK)
        if not warn:
            self._log("💬 用得顺手或踩了坑？点「反馈建议」匿名告诉作者（1 分钟）。")
            self._offer_recipe(cfg)

    def _open_desens_report(self):
        p = self._ds_report
        if not p or not os.path.exists(p):
            self._run_status.configure(text="⚠ 脱敏报告不在了，请重新跑一次", text_color=C_WARN)
            return
        if not open_folder(p):
            self._run_status.configure(text=f"⚠ 打不开脱敏报告，文件在：{p}", text_color=C_WARN)

    def _desens_merged(self):
        """汇总完另存脱敏版（默认「对外报送」：个人信息整列清空）。原汇总表不动。"""
        p = self._col_merged
        if not p or not os.path.exists(p):
            self._run_status.configure(text="⚠ 汇总表不在了，请重新汇总一次", text_color=C_WARN)
            return
        self._col_desens_btn.configure(state="disabled")
        self._run_status.configure(text="⏳ 正在生成脱敏版…", text_color=C_MUTED)

        def work():
            try:
                from core.desens import desens_copy
                out, job = desens_copy(p, "对外报送", lambda m: self._ui_q.put(("log", m)))
                res = (out, job.summary())
            except Exception as e:
                self._ui_q.put(("log", f"❌ 脱敏版没生成：{e}"))
                res = ("", None)
            self._ui_q.put(("desens_copy", res))

        threading.Thread(target=work, daemon=True).start()

    def _on_desens_copy(self, res):
        out, s = res
        self._col_desens_btn.configure(state="normal")
        if out:
            self._run_status.configure(
                text=f"✅ 脱敏版已另存：{os.path.basename(out)}（个人信息整列清空，复检通过），旁边有脱敏报告",
                text_color=C_OK)
        elif s:
            self._run_status.configure(text="⚠ 脱敏版没通过复检，没有写出来；汇总结果里的「未通过复检_请勿发出.txt」写了是哪列",
                                       text_color=C_WARN)
        else:
            self._run_status.configure(text="❌ 脱敏版没生成，原因见「处理详情」", text_color=C_ERR)
            self._toggle_log(show=True)
