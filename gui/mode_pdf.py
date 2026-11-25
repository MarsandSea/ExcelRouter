"""
ExcelRouter - 图形界面：「文件加密分发」模式。

①选文件 ②映射清单 + 保护方式，以及开始分发 _start_pdf。

本模块只是 App 的一块（mixin）：控件、状态都挂在同一个 App 实例上，
方法按名字互相调用，和写在 gui/app.py 里时完全一样。

Copyright (c) 2026 AbeLin
MIT License
"""
# App 的属性分散在几个 mixin 里，pyright 在本文件里看不到别处定义的 self.xxx
# pyright: reportAttributeAccessIssue=false

import os
import threading

import customtkinter as ctk
from tkinter import filedialog, messagebox

from gui.common import PDF_COL_PLACEHOLDER, DIST_EXTS, C_OK, C_WARN, C_MUTED, _ghost_button


class PdfModeMixin:
    def _on_drop_pdfs(self, event):
        """文件加密分发模式拖放：接受一批 PDF / Word / Excel / PPT。"""
        paths = [p for p in self._parse_drop_paths(getattr(event, "data", ""))
                 if p.lower().endswith(DIST_EXTS)]
        if not paths:
            self._pdf_in_status.configure(
                text="⚠ 请拖入 PDF / Word / Excel / PPT 文件（.doc / .xls 老格式请先另存为新格式）",
                text_color=C_WARN)
            return
        self._pdf_paths = [os.path.normpath(p) for p in paths]
        self._update_pdf_in_status()
        self._suggest_pdf_output()
        self._update_pdf_preview()

    def _on_drop_mapping(self, event):
        """映射清单拖放：接受单个 xlsx。"""
        paths = self._parse_drop_paths(getattr(event, "data", ""))
        target = next((p for p in paths if p.lower().endswith(".xlsx")), None)
        if target is None:
            self._pdf_map_status.configure(text="⚠ 映射清单请拖入 Excel 文件（.xlsx）",
                                           text_color=C_WARN)
            return
        self._pdf_map_var.set(os.path.normpath(target))
        self._scan_mapping()

    # ── 文件加密分发模式的步骤卡（内部名仍是 pdf，v2.10 起也收 Word/Excel/PPT）──
    def _build_pdf_step_input(self, parent):
        card, _ = self._step_card(parent, 0, 1, "选择要分发的文件")

        btns = ctk.CTkFrame(card, fg_color="transparent")
        btns.grid(row=1, column=0, padx=12, pady=(6, 2), sticky="w")
        ctk.CTkButton(btns, text="📄 选文件（PDF / Word / Excel / PPT，可多选）", width=300,
                      command=self._pick_pdfs).pack(side="left")
        _ghost_button(btns, text="清空", width=60,
                      command=self._clear_pdfs).pack(side="left", padx=(10, 0))
        self._enable_drop(card, self._on_drop_pdfs)

        self._pdf_in_status = ctk.CTkLabel(card, text="", font=ctk.CTkFont(size=11),
                                           text_color=C_MUTED, anchor="w", justify="left")
        self._pdf_in_status.grid(row=2, column=0, padx=12, pady=(0, 10), sticky="w")
        self._update_pdf_in_status()

    def _build_pdf_step_map(self, parent):
        card, _ = self._step_card(parent, 1, 2, "选择映射清单（Excel：网格 → 密码）")

        prow = ctk.CTkFrame(card, fg_color="transparent")
        prow.grid(row=1, column=0, padx=12, pady=(6, 2), sticky="ew")
        prow.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(prow, text="📋 选映射清单", width=120,
                      command=self._pick_mapping).grid(row=0, column=0)
        self._pdf_map_var = ctk.StringVar(value=self.cfg.get("pdf_mapping_path", ""))
        entry = ctk.CTkEntry(prow, textvariable=self._pdf_map_var)
        entry.grid(row=0, column=1, padx=8, sticky="ew")
        entry.bind("<Return>",   lambda e: self._on_map_edited())
        entry.bind("<FocusOut>", lambda e: self._on_map_edited())
        self._enable_drop(entry, self._on_drop_mapping)
        _ghost_button(prow, text="生成模板", width=90,
                      command=self._make_pdf_template).grid(row=0, column=2)

        ctk.CTkLabel(card, text="清单每行一个网格：网格名和专属密码（只加水印时不用密码）；"
                                "有「接收人」列会自动抄进分发清单；没有密码列时开始前可自动生成",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED).grid(
            row=2, column=0, padx=12, pady=(2, 0), sticky="w")

        crow = ctk.CTkFrame(card, fg_color="transparent")
        crow.grid(row=3, column=0, padx=12, pady=(6, 2), sticky="w")
        ctk.CTkLabel(crow, text="网格列").pack(side="left")
        self._pdfgrid_var = ctk.StringVar(value=self.cfg.get("pdf_grid_column") or PDF_COL_PLACEHOLDER)
        self._pdfgrid_menu = ctk.CTkOptionMenu(crow, variable=self._pdfgrid_var,
                                               values=[self._pdfgrid_var.get()], width=140)
        self._pdfgrid_menu.pack(side="left", padx=(6, 12))
        self._pdfpwd_lbl = ctk.CTkLabel(crow, text="密码列")
        self._pdfpwd_lbl.pack(side="left")
        self._pdfpwd_lbl_color = self._pdfpwd_lbl.cget("text_color")
        self._pdfpwd_var = ctk.StringVar(value=self.cfg.get("pdf_password_column") or PDF_COL_PLACEHOLDER)
        self._pdfpwd_menu = ctk.CTkOptionMenu(crow, variable=self._pdfpwd_var,
                                              values=[self._pdfpwd_var.get()], width=140)
        self._pdfpwd_menu.pack(side="left", padx=(6, 0))

        self._pdf_map_status = ctk.CTkLabel(card, text="", font=ctk.CTkFont(size=11),
                                            text_color=C_MUTED, anchor="w")
        self._pdf_map_status.grid(row=4, column=0, padx=12, pady=(0, 6), sticky="w")

        # 保护方式：加密、水印各自可选（v2.10，作者：有些人只要其中一个），至少留一个
        wm = ctk.CTkFrame(card)
        wm.grid(row=5, column=0, padx=12, pady=(2, 12), sticky="ew")
        opts = ctk.CTkFrame(wm, fg_color="transparent")
        opts.grid(row=0, column=0, padx=12, pady=(8, 4), sticky="w")
        self._pdfenc_var = ctk.BooleanVar(value=self.cfg.get("pdf_encrypt", True))
        ctk.CTkCheckBox(opts, text="加打开密码（每个网格一个）", variable=self._pdfenc_var,
                        command=lambda: self._on_protect_toggle(self._pdfenc_var, self._pdfwm_var)
                        ).pack(side="left")
        self._pdfwm_var = ctk.BooleanVar(value=self.cfg.get("pdf_watermark", True))
        ctk.CTkCheckBox(opts, text="加水印（仅 PDF，泄露可溯源）", variable=self._pdfwm_var,
                        command=lambda: self._on_protect_toggle(self._pdfwm_var, self._pdfenc_var)
                        ).pack(side="left", padx=(18, 0))
        wrow = ctk.CTkFrame(wm, fg_color="transparent")
        wrow.grid(row=1, column=0, padx=(38, 12), pady=(0, 8), sticky="w")
        ctk.CTkLabel(wrow, text="水印文字").pack(side="left")
        self._pdfwmtext_var = ctk.StringVar(value=self.cfg.get("pdf_watermark_text", "{grid} {date}"))
        self._pdfwmtext_entry = ctk.CTkEntry(wrow, textvariable=self._pdfwmtext_var, width=220)
        self._pdfwmtext_entry.pack(side="left", padx=(6, 8))
        ctk.CTkLabel(wrow, text="{grid}＝网格名  {receiver}＝接收人  {date}＝日期",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED).pack(side="left")
        if not (self._pdfenc_var.get() or self._pdfwm_var.get()):
            self._pdfenc_var.set(True)    # 旧配置两项全关时兜底
        self._update_wm_state()

    def _on_protect_toggle(self, changed, other):
        """加密 / 水印至少留一个：把最后一个勾掉时自动勾回，并说明原因。"""
        if not changed.get() and not other.get():
            changed.set(True)
            self._pdf_map_status.configure(text="⚠ 加密和水印至少要选一个", text_color=C_WARN)
        self._update_wm_state()

    def _update_wm_state(self):
        self._pdfwmtext_entry.configure(state="normal" if self._pdfwm_var.get() else "disabled")
        enc = "normal" if self._pdfenc_var.get() else "disabled"
        self._pdfpwd_menu.configure(state=enc)
        self._pdfpwd_lbl.configure(text_color=C_MUTED if enc == "disabled" else self._pdfpwd_lbl_color)
        self._update_pdf_in_status()

    def _update_pdf_in_status(self):
        n = len(self._pdf_paths)
        if not n:
            self._pdf_in_status.configure(
                text="还没有选择文件；每个网格都会拿到全部所选文件。PDF 可加密和加水印，Word / Excel / PPT 只加密",
                text_color=C_MUTED)
            return
        names = [os.path.basename(p) for p in self._pdf_paths]
        shown = "、".join(names[:3]) + ("…" if n > 3 else "")
        n_off = sum(not p.lower().endswith(".pdf") for p in self._pdf_paths)
        enc = getattr(self, "_pdfenc_var", None)     # ①卡片先于②卡片搭，启动时还没有
        if n_off and enc is not None and not enc.get():
            self._pdf_in_status.configure(
                text=f"✓ 已选 {n} 个：{shown}　⚠ 其中 {n_off} 个 Word / Excel / PPT 只能加密，"
                     "只加水印时会跳过（要水印请先转成 PDF）", text_color=C_WARN)
        else:
            self._pdf_in_status.configure(text=f"✓ 已选 {n} 个：{shown}", text_color=C_OK)
        self._update_pdf_preview()   # 文件个数变化会改变「N 网格 × M 文件」预览

    def _pick_pdfs(self):
        pats = " ".join("*" + e for e in DIST_EXTS)
        paths = filedialog.askopenfilenames(filetypes=[("PDF / Word / Excel / PPT", pats),
                                                       ("PDF 文件", "*.pdf")])
        if paths:
            self._pdf_paths = [os.path.normpath(p) for p in paths]
            self._update_pdf_in_status()
            self._suggest_pdf_output()

    def _clear_pdfs(self):
        self._pdf_paths = []
        self._update_pdf_in_status()

    def _suggest_pdf_output(self):
        """文件加密分发的默认输出：第一个文件所在目录下「分发结果」。规则同 _suggest_output。"""
        cur = self._output_var.get().strip()
        if cur and cur != self._out_auto:
            return
        if self._pdf_paths:
            auto = os.path.join(os.path.dirname(self._pdf_paths[0]), "分发结果")
            self._output_var.set(auto)
            self._out_auto = auto

    def _pick_mapping(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self._pdf_map_var.set(os.path.normpath(path))
            self._scan_mapping()

    def _on_map_edited(self):
        p = self._pdf_map_var.get().strip()
        if p == self._scanned_map_path:
            return
        if p and os.path.exists(p):
            self._scan_mapping()
        elif p:
            self._scanned_map_path = p
            self._pdf_map_status.configure(text="⚠ 找不到这个文件，请检查路径", text_color=C_WARN)

    def _scan_mapping(self):
        """子线程读映射清单表头，经 UI 泵回填三个列下拉（机制同 _scan_input）。"""
        p = self._pdf_map_var.get().strip()
        self._scanned_map_path = p
        if not p or not os.path.exists(p):
            return
        self._pdf_map_status.configure(text="正在读取清单…", text_color=C_MUTED)

        def work():
            cols, n_rows = [], 0
            try:
                from core.pdf_dist import list_mapping_columns
                cols = list_mapping_columns(p)
                # 顺带数网格行数（分发预览「N 个网格 × M 个文件」用）；清单是小表，直接全读。
                # 逐行数有内容的行：只读模式的 max_row 来自文件里写的尺寸，程序生成的表常写错
                from openpyxl import load_workbook
                wb = load_workbook(p, read_only=True)
                try:
                    ws = wb.worksheets[0]
                    ws.reset_dimensions()
                    n_rows = sum(1 for row in ws.iter_rows(min_row=2, values_only=True)
                                 if any(v is not None and str(v).strip() for v in row))
                finally:
                    wb.close()
            except Exception:
                cols, n_rows = [], 0
            self._ui_q.put(("pdf_scan", (cols, p, n_rows)))

        threading.Thread(target=work, daemon=True).start()

    def _on_pdf_scan(self, payload):
        cols, p, n_rows = payload
        if p != self._pdf_map_var.get().strip():
            return    # 结果已过期
        self._pdf_n_cols = len(cols)
        self._pdf_grid_rows = n_rows
        if not cols:
            self._pdf_map_status.configure(
                text="⚠ 没读到表头：请确认清单第 1 行是列名（如 网格 / 密码 / 接收人）", text_color=C_WARN)
            return
        self._pdfgrid_menu.configure(values=cols)
        if self._pdfgrid_var.get() not in cols:
            self._pdfgrid_var.set(self._recommend_col(cols, ("网格", "部门", "区域", "单位", "分组", "组")) or cols[0])
        self._pdfpwd_menu.configure(values=cols)
        if self._pdfpwd_var.get() not in cols:
            self._pdfpwd_var.set(self._recommend_col(cols, ("密码", "口令", "pwd", "密")) or cols[-1])
        # 接收人列不再让用户选（v2.10）：它只抄进分发清单 / 水印 {receiver}，core 按列名自动认
        self._update_pdf_preview()

    def _update_pdf_preview(self):
        """分发预览（C3）：选完清单和文件后，直接告诉用户将产出多少份副本。"""
        g, m, cols_n = self._pdf_grid_rows, len(self._pdf_paths), self._pdf_n_cols
        if not cols_n or not g:
            return
        extra = f"；{g} 个网格 × {m} 个文件 = {g * m} 份副本" if m else ""
        self._pdf_map_status.configure(
            text=f"✓ 识别到 {cols_n} 列 / {g} 个网格{extra}，确认下拉选的列对不对",
            text_color=C_OK)

    def _make_pdf_template(self):
        """一键生成「网格 → 密码」映射清单模板并自动载入（C1：解决没有清单的第一步卡点）。"""
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile="映射清单模板.xlsx",
            filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        try:
            from core.pdf_dist import write_mapping_template
            write_mapping_template(path)
        except Exception as e:
            messagebox.showerror("生成失败", f"生成模板失败：{e}")
            return
        self._pdf_map_var.set(os.path.normpath(path))
        self._scan_mapping()

    def _start_pdf(self):
        cfg = self._collect_config()
        if not cfg["pdf_input_paths"]:
            messagebox.showwarning("先选择文件", "请先在第 ① 步选择要分发的文件。")
            return
        missing = [p for p in cfg["pdf_input_paths"] if not os.path.exists(p)]
        if missing:
            messagebox.showwarning("文件不存在", "以下文件找不到了，请重新选择：\n" + "\n".join(missing[:3]))
            return
        if not cfg["pdf_encrypt"] and not any(p.lower().endswith(".pdf") for p in cfg["pdf_input_paths"]):
            messagebox.showwarning("没有能加水印的文件",
                                   "只加水印时只能处理 PDF；Word / Excel / PPT 只能加密。\n"
                                   "请勾上「加打开密码」，或先把文件转成 PDF。")
            return
        if not cfg["pdf_mapping_path"] or not os.path.exists(cfg["pdf_mapping_path"]):
            messagebox.showwarning("先选择映射清单", "请在第 ② 步选择列出各网格的 Excel 映射清单。")
            return
        if not cfg["pdf_grid_column"]:
            messagebox.showwarning("先选择网格列", "请在第 ② 步选好「网格列」。\n如果下拉框没有内容，重新选一次映射清单。")
            return
        if cfg["pdf_encrypt"] and not cfg["pdf_password_column"]:
            # 没选密码列 → 提供自动生成随机密码（C2），另存新清单不改动原文件
            if not messagebox.askyesno("自动生成密码？",
                                       "没有选择密码列。要我为每个网格自动生成随机密码，\n"
                                       "另存一份「含密码」清单并继续分发吗？\n（8 位随机密码只存在清单里，丢失无法找回）"):
                return
            try:
                from core.pdf_dist import fill_random_passwords
                out_path, npwd = fill_random_passwords(cfg["pdf_mapping_path"],
                                                       cfg["pdf_grid_column"])
            except Exception as e:
                messagebox.showerror("生成失败", f"自动生成密码失败：{e}")
                return
            cfg["pdf_mapping_path"] = out_path
            cfg["pdf_password_column"] = "密码"
            self._pdf_map_var.set(out_path)
            self._pdfpwd_var.set("密码")
            self._pdf_map_status.configure(
                text=f"🔑 已为 {npwd} 个网格生成随机密码 → {os.path.basename(out_path)}（含明文，保管好）",
                text_color=C_OK)
        if not cfg["output_path"]:
            self._suggest_pdf_output()
            cfg["output_path"] = self._output_var.get().strip()
        if not cfg["output_path"]:
            messagebox.showwarning("先选择保存位置", "请在第 ③ 步选择结果保存到哪个文件夹。")
            return
        try:
            os.makedirs(cfg["output_path"], exist_ok=True)
        except OSError as e:
            messagebox.showerror("无法创建保存位置",
                                 f"没法在这里创建结果文件夹：\n{cfg['output_path']}\n\n{e}\n\n"
                                 "请点「浏览」换一个能写入的位置（比如桌面或文档）。")
            return

        self._enter_running("⏳ 正在分发…", "已开始，正在读取映射清单…")

        def run():
            try:
                from core.pdf_dist import run_pdf_dist
                output_path = run_pdf_dist(
                    cfg,
                    log_fn=lambda m: self._ui_q.put(("log", m)),
                    progress_fn=lambda v: self._ui_q.put(("progress", v)),
                    stop_flag=lambda: self._stop_flag,
                )
            except Exception as e:
                self._ui_q.put(("log", f"\n❌ 运行出错：{e}"))
                output_path = None
            self._ui_q.put(("done", (cfg, output_path)))

        threading.Thread(target=run, daemon=True).start()
