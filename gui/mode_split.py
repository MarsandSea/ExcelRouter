"""
ExcelRouter - 图形界面：「Excel 拆分」模式。

①选表格 ②选字段（取值预览、批量选项、发出去时打码、通报表）▸高级设置，以及开始拆分 _start（也负责按模式分派）。

本模块只是 App 的一块（mixin）：控件、状态都挂在同一个 App 实例上，
方法按名字互相调用，和写在 gui/app.py 里时完全一样。

Copyright (c) 2026 AbeLin
MIT License
"""
# App 的属性分散在几个 mixin 里，pyright 在本文件里看不到别处定义的 self.xxx
# pyright: reportAttributeAccessIssue=false

import os
import re
import json
import threading

import customtkinter as ctk
from tkinter import filedialog, messagebox

from gui.platform_compat import lock_owner_name
from gui.common import (
    COL_PLACEHOLDER, SPEED_FAST, DS_SCENE_TEXT, C_OK, C_WARN, C_MUTED, _ghost_button,
    _flat_button,
)


class SplitModeMixin:
    def _build_step_input(self, parent):
        card, _ = self._step_card(parent, 0, 1, "选择要拆的表格")

        btns = ctk.CTkFrame(card, fg_color="transparent")
        btns.grid(row=1, column=0, padx=12, pady=(6, 2), sticky="w")
        ctk.CTkButton(btns, text="📄 选一个 Excel 文件", width=170,
                      command=self._pick_file).pack(side="left")
        ctk.CTkButton(btns, text="📁 选整个文件夹（批量拆）", width=190,
                      command=self._pick_folder).pack(side="left", padx=(10, 0))

        self._input_var = ctk.StringVar(value=self.cfg.get("input_path", ""))
        entry = ctk.CTkEntry(card, textvariable=self._input_var)
        entry.grid(row=2, column=0, padx=12, pady=(6, 2), sticky="ew")
        entry.bind("<Return>",   lambda e: self._on_path_edited())
        entry.bind("<FocusOut>", lambda e: self._on_path_edited())
        # 拖拽：整个卡片和路径框都接收拖入的文件/文件夹
        self._enable_drop(card, self._on_drop_input)
        self._enable_drop(entry, self._on_drop_input)

        self._in_status = ctk.CTkLabel(card, text=self._idle_input_text(),
                                       font=ctk.CTkFont(size=11), text_color=C_MUTED,
                                       anchor="w", justify="left")
        self._in_status.grid(row=3, column=0, padx=12, pady=(0, 10), sticky="w")

    def _build_step_field(self, parent):
        card, _ = self._step_card(parent, 1, 2, "按哪个字段拆分")
        self._field_card = card   # 取值勾选面板的挂载点

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.grid(row=1, column=0, padx=12, pady=(6, 2), sticky="w")
        ctk.CTkLabel(row, text="拆分字段").pack(side="left")
        self._split_var = ctk.StringVar(value=self.cfg.get("split_column") or COL_PLACEHOLDER)
        self._split_menu = ctk.CTkOptionMenu(row, variable=self._split_var,
                                             values=[self._split_var.get()], width=220,
                                             command=self._on_split_col_change)
        self._split_menu.pack(side="left", padx=(8, 8))
        _ghost_button(row, text="🔄 重新识别", width=104,
                      command=self._scan_input).pack(side="left")
        # 自动挑的样本是文件夹里排序最靠前的那个表，它未必有代表性（表头写法不同、
        # 取值只覆盖一部分）。用户比程序更清楚哪张表是标准模板，给个入口自己指。
        self._tpl_btn = _ghost_button(row, text="📄 换样本表", width=104,
                                      command=self._pick_template)
        self._tpl_btn.pack(side="left", padx=(6, 0))
        self._tpl_btn.pack_forget()      # 仅批量（文件夹、多于 1 个表）时才有意义

        ctk.CTkLabel(card, text="拆分字段的每个取值各生成一个文件：比如按「部门」拆 → 销售部.xlsx、财务部.xlsx…",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED).grid(
            row=2, column=0, padx=12, pady=(2, 0), sticky="w")
        self._status = ctk.CTkLabel(card, text="", font=ctk.CTkFont(size=11),
                                    text_color=C_MUTED, anchor="w")
        self._status.grid(row=3, column=0, padx=12, pady=(0, 6), sticky="w")

        # 取值预览（v2.7）：选定拆分字段后自动枚举该列取值——点开始之前就能看到
        # 「将拆成哪几组」，并可展开勾选只拆其中一部分（替代高级设置里手填文本的主路径）
        vrow = ctk.CTkFrame(card, fg_color="transparent")
        vrow.grid(row=4, column=0, padx=12, pady=(0, 4), sticky="ew")
        # 摘要文案与按钮**分两行**：抽样提示那句话很长，和按钮挤在同一行时
        # 「扫描全部文件的分组」会被推出卡片右边缘（截图目检发现）。
        self._values_status = ctk.CTkLabel(vrow, text="选定拆分字段后，这里会列出将拆出的分组",
                                           font=ctk.CTkFont(size=11), text_color=C_MUTED,
                                           anchor="w", justify="left", wraplength=600)
        self._values_status.pack(anchor="w", fill="x")
        vbtns = ctk.CTkFrame(vrow, fg_color="transparent")
        vbtns.pack(anchor="w", pady=(2, 0))
        self._scanall_btn = _flat_button(vbtns, text="🔍 扫描全部文件的分组", width=170,
                                         command=self._scan_all_values)
        self._values_btn = _flat_button(vbtns, text="▸ 查看/勾选分组", width=110,
                                        command=self._toggle_values)
        # 两个按钮都由 _refresh_values_ui 按状态 pack/forget（pack 顺序即左右顺序）
        self._scanall_btn.pack_forget()
        self._values_btn.pack_forget()   # 两者都由 _refresh_values_ui 按状态显隐
        # 展开中的勾选面板 grid 到 row=5（_toggle_values 控制）

        self._batch_frame = self._build_batch(card)
        self._batch_frame.grid(row=6, column=0, padx=12, pady=(2, 12), sticky="ew")
        self._batch_frame.grid_remove()   # 仅输入为文件夹时显示（_update_input_ui 控制）

        # 让对方填写的列（v2.10，选填）：拆出去的是「待填写的表」，收回来用「Excel 汇总」查收
        frow = ctk.CTkFrame(card, fg_color="transparent")
        frow.grid(row=7, column=0, padx=12, pady=(0, 12), sticky="ew")
        ctk.CTkLabel(frow, text="让对方填写的列（选填）").pack(side="left")
        self._fill_var = ctk.StringVar(value=str(self.cfg.get("fill_columns", "") or ""))
        ctk.CTkEntry(frow, textvariable=self._fill_var, width=300,
                     placeholder_text="如 走访日期（日期）、走访结果（已走访/未走访）").pack(side="left", padx=(8, 12))
        ctk.CTkLabel(frow, text="截止时间").pack(side="left")
        self._due_var = ctk.StringVar(value=str(self.cfg.get("due", "") or ""))
        ctk.CTkEntry(frow, textvariable=self._due_var, width=130,
                     placeholder_text="如 10月15日 18:00").pack(side="left", padx=(8, 0))
        ctk.CTkLabel(card, text="填了的话：这几列标浅黄可编辑、其余列锁定，并生成下发记录和发送清单；"
                                "收回来后切到「Excel 汇总」一键查谁没交、谁漏填。括号里可写要求："
                                "（日期）（数字）（整数）（手机号）（电话）（选填），或用 / 隔开的选项做成下拉框，"
                                "填错的收回来也查得出",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED, anchor="w", justify="left",
                     wraplength=640).grid(row=8, column=0, padx=12, pady=(0, 10), sticky="w")

        # 发出去时打码（v3.1）：样本表里认出个人信息列才出现；证件号、银行卡号默认打勾
        drow = self._split_ds_row = ctk.CTkFrame(card, fg_color="transparent")
        drow.grid(row=9, column=0, padx=12, pady=(0, 10), sticky="ew")
        self._split_ds_lbl = ctk.CTkLabel(drow, text="", anchor="w", justify="left", wraplength=640,
                                          text_color=C_WARN)
        self._split_ds_lbl.grid(row=0, column=0, columnspan=3, sticky="w")
        self._split_ds_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(drow, text="发出去时打码", variable=self._split_ds_var,
                        command=self._on_split_desens_toggle).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ctk.CTkLabel(drow, text="用途").grid(row=1, column=1, sticky="w", padx=(12, 6), pady=(4, 0))
        self._split_ds_scene = ctk.CTkOptionMenu(drow, values=["内部下发", "发群通报", "交给AI分析", "对外报送"],
                                                 width=130, command=lambda _v: self._on_split_desens_toggle())
        self._split_ds_scene.set(self.cfg.get("split_desens") or "内部下发")
        self._split_ds_scene.grid(row=1, column=2, sticky="w", pady=(4, 0))
        self._split_ds_tip = ctk.CTkLabel(drow, text="", font=ctk.CTkFont(size=11), text_color=C_MUTED,
                                          anchor="w", justify="left", wraplength=640)
        self._split_ds_tip.grid(row=2, column=0, columnspan=3, sticky="w")
        drow.grid_remove()

        # 通报表（v3.2）：样本表里认出「序号」列或合计行才出现；两个都默认不勾，不勾时和以前完全一样
        rrow = self._rpt_row = ctk.CTkFrame(card, fg_color="transparent")
        rrow.grid(row=10, column=0, padx=12, pady=(0, 10), sticky="ew")
        self._rpt_lbl = ctk.CTkLabel(rrow, text="", anchor="w", justify="left", wraplength=640)
        self._rpt_lbl.grid(row=0, column=0, columnspan=2, sticky="w")
        self._renumber_var = ctk.BooleanVar(value=bool(self.cfg.get("renumber_seq", False)))
        self._renumber_cb = ctk.CTkCheckBox(rrow, text="拆完序号每组从 1 重新编", variable=self._renumber_var)
        self._renumber_cb.grid(row=1, column=0, sticky="w", pady=(4, 0), padx=(0, 16))
        self._totals_var = ctk.BooleanVar(value=bool(self.cfg.get("group_totals", False)))
        self._totals_cb = ctk.CTkCheckBox(rrow, text="每组补一行自己的合计", variable=self._totals_var)
        self._totals_cb.grid(row=1, column=1, sticky="w", pady=(4, 0))
        ctk.CTkLabel(rrow, text="不勾：序号照抄原表、原表的合计行不发（和以前一样）。勾上合计：按每组自己的行重算"
                                "求和、平均、人数、完成率这类两列之比，算不出的格子留空并在处理详情里点名",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED, anchor="w", justify="left",
                     wraplength=640).grid(row=2, column=0, columnspan=2, sticky="w")
        rrow.grid_remove()

    def _build_batch(self, parent):
        f = ctk.CTkFrame(parent)
        f.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(f, text="批量选项（拆整个文件夹时）",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=("gray25", "gray75")).grid(row=0, column=0, padx=12, pady=(8, 0), sticky="w")

        self._zip_var = ctk.BooleanVar(value=self.cfg.get("make_zip", True))
        ctk.CTkCheckBox(f, text="每个分组打包成 ZIP（方便整体转发）",
                        variable=self._zip_var).grid(row=1, column=0, padx=12, pady=(8, 4), sticky="w")

        self._person_var = ctk.BooleanVar(value=self.cfg.get("to_person", False))
        ctk.CTkCheckBox(f, text="同时拆到人（每人再单独生成一个文件）",
                        variable=self._person_var,
                        command=self._update_person_state).grid(row=2, column=0, padx=12, pady=4, sticky="w")

        prow = ctk.CTkFrame(f, fg_color="transparent")
        prow.grid(row=3, column=0, padx=(38, 12), pady=(0, 2), sticky="w")
        ctk.CTkLabel(prow, text="按哪个字段区分人").pack(side="left")
        self._pcol_var = ctk.StringVar(value=self.cfg.get("person_column") or COL_PLACEHOLDER)
        self._pcol_menu = ctk.CTkOptionMenu(prow, variable=self._pcol_var,
                                            values=[self._pcol_var.get()], width=150)
        self._pcol_menu.pack(side="left", padx=(6, 10))
        ctk.CTkLabel(prow, text="只处理文件名含").pack(side="left")
        self._pfilter_var = ctk.StringVar(value=", ".join(self.cfg.get("person_file_filter", [])))
        self._pfilter_entry = ctk.CTkEntry(prow, textvariable=self._pfilter_var, width=150,
                                           placeholder_text="如 工资,费用")
        self._pfilter_entry.pack(side="left", padx=6)
        ctk.CTkLabel(prow, text="的表（留空＝全部）").pack(side="left")

        ctk.CTkLabel(f, text="「到人」只对上面指定的表生效，其余文件只进汇总。",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED).grid(
            row=4, column=0, padx=(38, 12), pady=(0, 8), sticky="w")
        self._update_person_state()
        return f

    def _update_person_state(self):
        """「同时拆到人」未勾选时置灰它的子选项，减少一次性摆在面前的决策数。"""
        state = "normal" if self._person_var.get() else "disabled"
        self._pcol_menu.configure(state=state)
        self._pfilter_entry.configure(state=state)

    def _idle_input_text(self):
        """①卡片未选输入时的提示语；支持拖拽时优先教拖拽（零学习成本的操作）。"""
        if self._dnd_ok:
            return "还没有选择文件（可点上方按钮，或直接把文件 / 文件夹拖进来）"
        return "还没有选择文件（也可以把路径粘贴到上面的输入框）"

    def _on_drop_input(self, event):
        """Excel 模式拖放：接受单个 Excel 文件或文件夹（多个文件时取第一个所在项）；拖进方案文件就打开方案。"""
        paths = self._parse_drop_paths(getattr(event, "data", ""))
        if self._drop_recipe(paths):
            return
        target = next((p for p in paths
                       if os.path.isdir(p) or p.lower().endswith((".xlsx", ".xls"))), None)
        if target is None:
            self._set_in_status("⚠ 请拖入 Excel 文件（.xlsx / .xls）或装表格的文件夹", C_WARN)
            return
        self._input_var.set(os.path.normpath(target))
        self._after_pick()

    # ── 取值预览（v2.7，A1）──────────────────────────────
    def _on_split_col_change(self, _choice):
        """拆分字段变更：防抖 300ms 后重新枚举该列取值（连续切换不重复读文件）。"""
        self._enum_token += 1
        if self._enum_after is not None:
            self.after_cancel(self._enum_after)
        self._enum_after = self.after(300, self._enum_values)

    def _enum_values(self):
        """子线程枚举当前拆分字段的全部取值（复用扫描时识别的模板文件）。"""
        self._enum_after = None
        col = self._split_var.get()
        # 换了字段，之前的「已全量扫描」结论随之作废（新字段只枚举了样本表）
        self._values_scope = "sample"
        self._scan_all_token += 1
        self._values_all, self._value_vars = [], {}
        self._refresh_values_ui()
        if (col in ("", COL_PLACEHOLDER) or not self._tpl_path
                or not os.path.exists(self._tpl_path)):
            return
        token = self._enum_token
        tpl, cfg = self._tpl_path, self._collect_config()
        self._values_status.configure(text="正在枚举该字段的取值…", text_color=C_MUTED)

        def work():
            try:
                from core.splitter import list_values
                vals = list_values(tpl, cfg, col)
            except Exception as e:
                vals = None
                # 不静默吞掉：失败原因进日志，界面只显示笼统提示（与 _scan_input 同理）
                self._ui_q.put(("log", f"⚠ 枚举「{col}」的取值失败：{e}"))
            self._ui_q.put(("values", (col, token, vals)))

        threading.Thread(target=work, daemon=True).start()

    def _on_values(self, payload):
        """枚举结果回填：过期（已换字段/又发起新枚举）结果直接丢弃。"""
        col, token, vals = payload
        if token != self._enum_token or col != self._split_var.get():
            return
        if vals is None:
            self._values_status.configure(
                text="⚠ 没能枚举取值（不影响拆分，可直接开始）", text_color=C_WARN)
            return
        self._set_values(vals)

    def _set_values(self, vals):
        """把枚举到的取值装进勾选面板（抽样枚举与全量扫描共用这一条路径）。"""
        self._values_all = vals
        # 默认全选；高级设置里手填过「只拆这些取值」且能对上号时，按它预选
        wanted = set(self._split_list(self._selvals_var.get()))
        use_filter = bool(wanted) and any(w in vals for w in wanted)
        if wanted and not use_filter:
            # 换了拆分字段（或换了输入）之后，上一轮的名单对新字段一个都对不上。
            # 留着它 = 这一跑按不存在的取值筛选 → 一行都拆不出来，界面却还挂着
            # 「🔒 只拆其中 N 组」。清掉，并在日志里点名说清楚，别让它静默生效。
            self._selvals_var.set("")
            self._log("⚠ 原先限定的分组（" + "、".join(sorted(wanted))
                      + "）在当前字段里一个都没有，已自动取消限定，本次将按全部分组拆")
        self._values_locked = use_filter
        self._value_vars = {}
        for v in vals:
            var = ctk.BooleanVar(value=(v in wanted) if use_filter else True)
            var.trace_add("write", lambda *_: self._sync_selvals())
            self._value_vars[v] = var
        self._refresh_values_ui()
        if self._values_panel is not None:      # 面板开着就重建内容
            self._close_values_panel()
            self._toggle_values()

    def _refresh_values_ui(self):
        """按 _values_all 更新取值摘要行：几组、示例、**来源**、异常多的告警。"""
        n = len(self._values_all)
        self._values_btn.pack_forget()
        self._scanall_btn.pack_forget()
        if not n:
            self._close_values_panel()
            if self._split_var.get() not in ("", COL_PLACEHOLDER) and self._tpl_path:
                self._values_status.configure(
                    text="选定拆分字段后，这里会列出将拆出的分组", text_color=C_MUTED)
            return
        shown = "、".join(self._values_all[:6]) + ("…" if n > 6 else "")
        sel = self._split_list(self._selvals_var.get())
        from_sample = self._tpl_is_sample and self._values_scope != "all"
        # 已显式限定名单时不再喊「其他表可能还有别的分组」——名单对全部文件生效，
        # 那句警告已经不成立了，和 🔒 角标并排显示只会自相矛盾。
        sampled = from_sample and not sel
        if sampled:
            # ★ 关键的诚实性修复（v2.9）：取值只枚举自一个样本表，拆分却对全部
            # 文件生效——真实分组数几乎总是更多。旧文案直说「将拆成 6 组」，用户
            # 按它理解、跑完看到 29 组，会认为是工具算错了而不是自己看的是抽样。
            # 数字本身没错，错的是把抽样结论说成了全量结论。**别改回去。**
            tpl = os.path.basename(self._tpl_path or "")
            self._values_status.configure(
                text=f"样本「{tpl}」里有 {n} 组：{shown}"
                     f"　⚠ 这只是 {self._scan_n} 个表中的 1 个，其他表可能还有别的分组",
                text_color=C_WARN)
        else:
            scope = (f"全部 {self._scan_n} 个表共 "
                     if self._values_scope == "all" and self._scan_n > 1 else "将拆成 ")
            if n > 50:
                self._values_status.configure(
                    text=f"⚠ {scope}{n} 组：{shown}——取值异常多，若这是工号 / 姓名类字段，建议换个分组字段",
                    text_color=C_WARN)
            elif n == 1:
                self._values_status.configure(
                    text=f"只会拆出 1 组「{self._values_all[0]}」——确认字段没选错？",
                    text_color=C_WARN)
            else:
                self._values_status.configure(
                    text=f"{scope}{n} 组：{shown}", text_color=C_OK)
        if sel:
            self._values_status.configure(
                text=self._values_status.cget("text") + f"　🔒 只拆其中 {len(sel)} 组",
                text_color=C_OK)
        kinds, nfiles = self._tpl_kinds
        if self._values_scope == "all" and kinds > 1:
            # v3.0：读遍全部表时顺便看出的——表其实不是同一套模板，跨文件合并时用户该心里有数
            self._values_status.configure(
                text=self._values_status.cget("text")
                     + f"　·　这 {nfiles} 个表其实是 {kinds} 种模板（详见处理详情）")
        # 「扫描全部」只要还在抽样态就一直可用：已限定名单的用户也可能想核对全量
        if from_sample:
            self._scanall_btn.pack(side="left", padx=(6, 0))
        self._values_btn.pack(side="left", padx=(6, 0))
        self._refresh_panel_hint()

    def _scan_all_values(self):
        """读遍输入里全部表格，把「抽样分组」换成完整分组清单。

        做成显式按钮而不是扫描时自动跑：代价是每个文件都要完整读一遍，
        400 个表的批次自动扫会让「选字段」这一步凭空卡好几分钟。
        """
        if self._scanning_all or self._running:
            return
        col = self._split_var.get()
        p = self._input_var.get().strip()
        if col in ("", COL_PLACEHOLDER) or not p or not os.path.isdir(p):
            return
        if self._scan_n > 30 and not messagebox.askyesno(
                "扫描全部分组",
                f"要列全分组，需要把 {self._scan_n} 个表各读一遍，可能要几分钟。\n"
                "（只是读取，不会改动任何文件）\n\n现在开始？"):
            return
        self._scanning_all = True
        self._scan_all_token += 1
        token = self._scan_all_token
        cfg = self._collect_config()
        out_now = self._output_var.get().strip()
        self._scanall_btn.configure(state="disabled")
        self._values_status.configure(text="正在读取全部表格的分组…", text_color=C_MUTED)

        def work():
            from core.splitter import scan_values, template_groups, template_lines
            # 文件清单的过滤规则与 run_split / _scan_input 保持一致，
            # 否则「扫描出的分组」和「真正拆出的分组」又会对不上
            out_abs = os.path.normcase(os.path.abspath(out_now)) if out_now else None
            targets = []
            for root, dirs, fs in os.walk(p):
                if out_abs:
                    ra = os.path.normcase(os.path.abspath(root))
                    if ra == out_abs or ra.startswith(out_abs + os.sep):
                        dirs[:] = []
                        continue
                for fn in sorted(fs):
                    lf = fn.lower()
                    if (lf.endswith(('.xlsx', '.xls')) and not fn.startswith('~$')
                            and not lf.endswith('__tmp__.xlsx')):
                        targets.append(os.path.join(root, fn))
            vals, failed, pairs = set(), 0, []
            for i, fp in enumerate(targets, 1):
                if token != self._scan_all_token:
                    return          # 用户换了输入/字段，本轮结果已作废
                try:
                    v, names = scan_values(fp, cfg, col)     # 一遍读出分组和这个表的模板
                    vals.update(v)
                    pairs.append((fp, names))
                except Exception as e:
                    failed += 1
                    self._ui_q.put(("log", f"⚠ 读取「{os.path.basename(fp)}」的分组失败：{e}"))
                self._ui_q.put(("scanall_prog", (token, i, len(targets))))
            groups = template_groups(pairs)
            tpl = (len(groups), sum(len(g["files"]) for g in groups),
                   template_lines(groups) if len(groups) > 1 else [])
            self._ui_q.put(("scanall", (token, col, sorted(vals), len(targets), failed, tpl)))

        threading.Thread(target=work, daemon=True).start()

    def _on_scan_all_prog(self, payload):
        token, i, total = payload
        if token != self._scan_all_token:
            return
        self._values_status.configure(
            text=f"正在读取全部表格的分组… {i}/{total}", text_color=C_MUTED)

    def _on_scan_all(self, payload):
        """全量扫描回填：此后摘要行显示的是全量结论，不再带抽样警告。"""
        token, col, vals, nfiles, failed, (kinds, ngrouped, tpl_lines) = payload
        self._scanning_all = False
        self._scanall_btn.configure(state="normal")
        if token != self._scan_all_token or col != self._split_var.get():
            return
        self._values_scope = "all"
        self._tpl_kinds = (kinds, ngrouped)
        self._set_values(vals)
        note = f"（其中 {failed} 个读取失败，见上面的 ⚠）" if failed else ""
        self._log(f"✓ 已读遍 {nfiles} 个表：「{col}」共有 {len(vals)} 个分组{note}")
        if kinds > 1:
            self._log(f"📋 这 {ngrouped} 个表其实是 {kinds} 种模板（含「{col}」的 sheet 列名不完全一样）：")
            for line in tpl_lines:
                self._log(line)
            self._log("  勾了「跨文件合并汇总」时按列名对齐；列名大半对不上的表单独放进「sheet 名_2」，不会错位"
                      if self._merge_var.get() else
                      "  没勾「跨文件合并汇总」，各表各拆，不受影响")

    def _pick_template(self):
        """手动指定样本表 —— 字段下拉和分组清单都从它读取。"""
        init = os.path.dirname(self._tpl_path) if self._tpl_path else self._input_var.get().strip()
        path = filedialog.askopenfilename(
            title="选一个有代表性的表格（字段和分组从它读取）",
            initialdir=init if init and os.path.isdir(init) else None,
            filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if not path:
            return
        cfg = self._collect_config()
        cur_in = self._input_var.get().strip()
        self._set_scan_status("正在识别表格里的字段…", C_MUTED)

        def work():
            try:
                from core.splitter import list_columns
                cols = list_columns(path, cfg)
            except Exception as e:
                cols = []
                self._ui_q.put(("log", f"⚠ 扫描「{os.path.basename(path)}」失败：{e}"))
            # 复用 _on_scan：文件数/体积/输入类型沿用上一次扫描的结果，只换样本表
            self._ui_q.put(("scan", (cols, self._scan_n, os.path.basename(path),
                                     self._scan_is_dir, cur_in, path, self._scan_bytes)))

        threading.Thread(target=work, daemon=True).start()

    def _toggle_values(self):
        """展开 / 收起取值勾选面板。"""
        if self._values_panel is not None:
            self._close_values_panel()
            return
        self._values_open = True
        self._values_btn.configure(text="▾ 收起分组")
        panel = ctk.CTkFrame(self._field_card)
        panel.grid(row=5, column=0, padx=12, pady=(0, 8), sticky="ew")
        panel.grid_columnconfigure(0, weight=1)
        bar = ctk.CTkFrame(panel, fg_color="transparent")
        bar.grid(row=0, column=0, padx=8, pady=(6, 0), sticky="w")
        _flat_button(bar, text="全选", width=46,
                     command=lambda: self._set_all_values(True)).pack(side="left")
        _flat_button(bar, text="清空", width=46,
                     command=lambda: self._set_all_values(False)).pack(side="left")
        # 「全选＝不筛选」的既有语义在抽样场景下会咬人：样本表里恰好就是用户
        # 想要的那几组时，全勾等于什么都没筛，其他表的分组照样会拆出来。
        # 这个按钮让「我只要这几组」变成一个能表达出来的动作。
        self._lock_btn = _flat_button(bar, text="只拆这几组", width=104,
                                      command=self._lock_to_listed)
        self._lock_btn.pack(side="left", padx=(4, 0))
        self._values_hint = ctk.CTkLabel(bar, text="", font=ctk.CTkFont(size=11),
                                         text_color=C_MUTED, wraplength=420,
                                         justify="left")
        self._values_hint.pack(side="left", padx=8)
        self._refresh_panel_hint()
        box = ctk.CTkScrollableFrame(panel, height=140)
        box.grid(row=1, column=0, padx=8, pady=(2, 8), sticky="ew")
        for v, var in self._value_vars.items():
            ctk.CTkCheckBox(box, text=v, variable=var).pack(anchor="w")
        self._values_panel = panel

    def _refresh_panel_hint(self):
        """勾选面板里的按钮与提示随「已限定 / 抽样」状态更新（面板开着才有意义）。

        面板是按当时状态一次性搭起来的，而状态会在面板开着时变（点了「只拆这 N 组」，
        或点「全选」取消限定），提示不跟着变就会自相矛盾。
        """
        if self._values_panel is None or self._lock_btn is None or self._values_hint is None:
            return
        n = len(self._value_vars)
        if self._values_locked:
            self._lock_btn.configure(text="✓ 已限定", state="disabled")
            txt = "已限定只拆勾选的这些分组，对全部表都生效；点「全选」可取消限定"
        else:
            self._lock_btn.configure(text=f"只拆这 {n} 组", state="normal")
            if self._tpl_is_sample and self._values_scope != "all":
                txt = ("只拆勾选的分组。注意：下面这些分组只来自样本表，"
                       "全勾＝不做筛选（其他表里的分组也会照样拆出来）")
            else:
                txt = "只拆勾选的分组；全选＝拆分全部"
        self._values_hint.configure(text=txt)

    def _close_values_panel(self):
        if self._values_panel is not None:
            self._values_panel.destroy()
            self._values_panel = None
            self._lock_btn = self._values_hint = None
        self._values_open = False
        self._values_btn.configure(text="▸ 查看/勾选分组")

    def _set_all_values(self, on):
        # 「全选 / 清空」是用户表达「不做筛选」的动作，同时解除已限定状态——
        # 否则点了全选，界面还挂着「已限定」，两边说的不是一回事。
        self._values_locked = False
        # 逐个 set 会逐个触发 trace → 逐次重绘摘要行；取值多时明显发卡。
        # 批量期间挂起同步，循环结束后统一同步一次。
        self._bulk_sync = True
        try:
            for var in self._value_vars.values():
                var.set(on)
        finally:
            self._bulk_sync = False
        self._sync_selvals()

    def _lock_to_listed(self):
        """把当前列出的分组**显式**写进「只拆这些取值」（抽样场景专用）。

        与「全选」的区别：全选等价于留空、等价于不筛选；这个按钮写的是一份
        具体名单，拆分时对全部文件生效——正是「我只要城区这几个网格」的表达。
        """
        vals = list(self._value_vars.keys())
        if not vals:
            return
        self._bulk_sync = True
        try:
            for var in self._value_vars.values():
                var.set(True)
        finally:
            self._bulk_sync = False
        self._values_locked = True     # 显式状态：此后即使全勾，名单也不会被清空
        self._sync_selvals()
        shown = "、".join(vals[:6]) + ("…" if len(vals) > 6 else "")
        self._log(f"✓ 已限定只拆这 {len(vals)} 个分组：{shown}")

    def _sync_selvals(self):
        """勾选状态 + 已限定状态 → selected_values（与高级设置文本框始终一致）。

        取消勾选任何一组＝明确要筛选，自动进入已限定；全不选则回到「不筛选」
        （留空＝拆分时自动枚举全部取值，与 core 的既有语义一致）。已限定时即使
        全勾也照样写出名单——这正是「只拆这 N 组」能顶住后续勾选操作的原因。
        """
        if not self._value_vars or self._bulk_sync:
            return
        checked = [v for v, var in self._value_vars.items() if var.get()]
        if not checked:
            self._values_locked = False
        elif len(checked) < len(self._value_vars):
            self._values_locked = True
        self._selvals_var.set(", ".join(checked) if self._values_locked else "")
        self._refresh_values_ui()   # 摘要行上的「🔒 已限定」角标要跟着变

    def _build_adv_area(self, parent):
        self._adv_btn = _flat_button(parent, text="▸ 高级设置（一般用不到）", width=200,
                                     anchor="w", command=self._toggle_adv)
        self._adv_btn.grid(row=2, column=0, padx=8, pady=(0, 4), sticky="w")
        self._adv_frame = self._build_advanced(parent)   # 展开时 grid 到 row=3

    def _build_advanced(self, parent):
        f = ctk.CTkFrame(parent)
        f.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(f, text="以下选项保持默认即可；只有扫描不到字段、或需要归并取值时才需要调整。",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED).grid(
            row=0, column=0, columnspan=2, padx=12, pady=(8, 2), sticky="w")

        ctk.CTkLabel(f, text="表头识别").grid(row=1, column=0, padx=12, pady=8, sticky="w")
        mode_map = {"auto": "自动", "row": "指定行号", "keyword": "关键词"}
        self._header_seg = ctk.CTkSegmentedButton(f, values=["自动", "指定行号", "关键词"],
                                                  command=self._on_header_mode)
        self._header_seg.set(mode_map.get(self.cfg.get("header_mode", "auto"), "自动"))
        self._header_seg.grid(row=1, column=1, padx=8, pady=8, sticky="w")

        self._hrow_frame = ctk.CTkFrame(f, fg_color="transparent")
        self._hrow_frame.grid(row=2, column=0, columnspan=2, padx=12, pady=2, sticky="w")
        ctk.CTkLabel(self._hrow_frame, text="表头行号（1 基）：").grid(row=0, column=0)
        self._hrow_var = ctk.StringVar(value=str(self.cfg.get("header_row", 1)))
        ctk.CTkEntry(self._hrow_frame, textvariable=self._hrow_var, width=80).grid(row=0, column=1, padx=6)

        self._kw_frame = ctk.CTkFrame(f, fg_color="transparent")
        self._kw_frame.grid(row=3, column=0, columnspan=2, padx=12, pady=2, sticky="ew")
        self._kw_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self._kw_frame, text="必含关键词A（逗号分隔）").grid(row=0, column=0, padx=4, pady=4, sticky="w")
        self._grid_keys_var = ctk.StringVar(value=", ".join(self.cfg.get("grid_keys", [])))
        ctk.CTkEntry(self._kw_frame, textvariable=self._grid_keys_var).grid(row=0, column=1, padx=6, pady=4, sticky="ew")
        ctk.CTkLabel(self._kw_frame, text="必含关键词B（逗号分隔）").grid(row=1, column=0, padx=4, pady=4, sticky="w")
        self._id_keys_var = ctk.StringVar(value=", ".join(self.cfg.get("id_keys", [])))
        ctk.CTkEntry(self._kw_frame, textvariable=self._id_keys_var).grid(row=1, column=1, padx=6, pady=4, sticky="ew")

        ctk.CTkLabel(f, text="跳过值（逗号分隔）").grid(row=4, column=0, padx=12, pady=8, sticky="w")
        self._skip_var = ctk.StringVar(value=", ".join(self.cfg.get("skip_values", [])))
        ctk.CTkEntry(f, textvariable=self._skip_var).grid(row=4, column=1, padx=8, pady=8, sticky="ew")

        ctk.CTkLabel(f, text="只拆这些取值（逗号分隔，留空=全部）").grid(row=5, column=0, padx=12, pady=8, sticky="w")
        self._selvals_var = ctk.StringVar(value=", ".join(self.cfg.get("selected_values", [])))
        ctk.CTkEntry(f, textvariable=self._selvals_var).grid(row=5, column=1, padx=8, pady=8, sticky="ew")

        ctk.CTkLabel(f, text="取值归并映射（JSON，选填）").grid(row=6, column=0, padx=12, pady=(8, 2), sticky="nw")
        self._alias_box = ctk.CTkTextbox(f, height=64, wrap="none")
        self._alias_box.grid(row=6, column=1, padx=8, pady=(8, 2), sticky="ew")
        alias = self.cfg.get("value_alias_map", {})
        if alias:
            self._alias_box.insert("1.0", json.dumps(alias, ensure_ascii=False, indent=2))

        opts = ctk.CTkFrame(f, fg_color="transparent")
        opts.grid(row=7, column=0, columnspan=2, padx=8, pady=8, sticky="w")
        self._exact_var    = ctk.BooleanVar(value=self.cfg.get("exact_match", True))
        self._merge_var    = ctk.BooleanVar(value=self.cfg.get("merge_across_files", True))
        self._preserve_var = ctk.BooleanVar(value=self.cfg.get("preserve_format", True))
        self._keep_formula_var = ctk.BooleanVar(value=self.cfg.get("keep_formulas", False))
        self._auto_open_var = ctk.BooleanVar(value=self.cfg.get("auto_open_output", True))
        self._keep_hidden_var = ctk.BooleanVar(value=self.cfg.get("keep_hidden", False))
        ctk.CTkCheckBox(opts, text="精确匹配", variable=self._exact_var).grid(row=0, column=0, padx=8, pady=4)
        ctk.CTkCheckBox(opts, text="跨文件合并汇总", variable=self._merge_var).grid(row=0, column=1, padx=8, pady=4)
        ctk.CTkCheckBox(opts, text="完成后打开输出", variable=self._auto_open_var).grid(row=0, column=2, padx=8, pady=4)
        # 默认：隐藏的 sheet 不发、隐藏列清空（仍隐藏）、隐藏的行照常发（多半是筛选）——别人点开就能看见的东西不该藏在里面
        ctk.CTkCheckBox(opts, text="隐藏的 sheet 和列也照样发出", variable=self._keep_hidden_var).grid(
            row=0, column=3, padx=8, pady=4)
        # 「保留格式」的勾选框已上移到第 ③ 步的「拆分速度」段选（v2.9），这里不再重复
        self._keep_formula_box = ctk.CTkCheckBox(
            opts, text="保留公式（收件人可见计算过程，仅同行公式，需选「标准」速度）",
            variable=self._keep_formula_var)
        self._keep_formula_box.grid(row=1, column=0, columnspan=4, padx=8, pady=(0, 4), sticky="w")
        ctk.CTkLabel(opts, text="提示：「跨文件合并汇总」把各表的同名 sheet 按列名合到一起（列顺序不同也对得上），"
                                "所有结果留在内存里最后统一写盘，宽表 + 大批量时明显更吃内存；默认的「按原表各自拆分」更稳。",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED,
                     wraplength=560, justify="left").grid(
            row=2, column=0, columnspan=4, padx=8, pady=(0, 6), sticky="w")

        self._on_header_mode(self._header_seg.get())
        return f

    # 拆分 ② 卡片里的「发出去时打码」：识别到个人信息列才出现，证件号、银行卡号默认打勾
    def _scan_split_desens(self, tpl_path, p):
        if not tpl_path:
            self._split_ds_row.grid_remove()
            return
        cfg = self._header_cfg()

        def work():
            cols = []
            try:
                from core.desens import inspect
                for s in inspect(tpl_path, "内部下发", cfg):
                    if not s["hidden"]:
                        cols += [(c["column"], c["kind"]) for c in s["columns"]]
            except Exception:
                cols = []
            self._ui_q.put(("split_desens", (p, list(dict.fromkeys(cols)))))

        threading.Thread(target=work, daemon=True).start()

    def _on_split_desens(self, payload):
        p, cols = payload
        if p != self._input_var.get().strip():
            return
        self._split_ds_cols = cols
        if not cols:
            self._split_ds_row.grid_remove()
            return
        shown = "、".join(f"{c}（{k}）" if c != k else c for c, k in cols[:4]) + (" 等" if len(cols) > 4 else "")
        self._split_ds_lbl.configure(text=f"🛡 发现个人信息列：{shown}")
        strong = any(k in ("身份证号", "银行卡号") for _c, k in cols)
        rec = self._recipe_active if (self._recipe_active or {}).get("mode") == "excel" else None
        want = (rec or {}).get("settings", {}).get("split_desens") if rec else None
        if want is not None:
            self._split_ds_var.set(bool(want))
            if want:
                self._split_ds_scene.set(want)
        else:
            self._split_ds_var.set(strong)
        self._split_ds_row.grid()
        self._on_split_desens_toggle()

    # 拆分 ② 卡片里的「通报表」一行：样本表有序号列或合计行才出现（v3.2）
    def _scan_report_hints(self, tpl_path, p):
        if not tpl_path:
            self._rpt_row.grid_remove()
            return
        cfg = self._header_cfg()

        def work():
            try:
                from core.splitter import report_hints
                hints = report_hints(tpl_path, cfg)
            except Exception:
                hints = {"seq_column": "", "total_rows": 0}
            self._ui_q.put(("report_hints", (p, hints)))

        threading.Thread(target=work, daemon=True).start()

    def _on_report_hints(self, payload):
        p, hints = payload
        if p != self._input_var.get().strip():
            return
        self._rpt_hints = hints
        seq, tot = hints.get("seq_column"), hints.get("total_rows")
        if not (seq or tot):
            self._rpt_row.grid_remove()
            return
        found = "、".join(([f"「{seq}」列"] if seq else []) + (["合计行"] if tot else []))
        self._rpt_lbl.configure(text=f"📋 这张表有{found}，像是通报表")
        if seq:
            self._renumber_cb.grid()
        else:
            self._renumber_cb.grid_remove()
        if tot:
            self._totals_cb.grid()
        else:
            self._totals_cb.grid_remove()
        self._rpt_row.grid()

    def _report_value(self, key):
        """renumber_seq / group_totals：提示行没出现（样本表没有序号列 / 合计行）时一律 False，和以前一样。"""
        if key == "renumber_seq":
            return bool(self._rpt_hints.get("seq_column")) and self._renumber_var.get()
        return bool(self._rpt_hints.get("total_rows")) and self._totals_var.get()

    def _on_split_desens_toggle(self):
        on = self._split_ds_var.get()
        self._split_ds_scene.configure(state="normal" if on else "disabled")
        scene = self._split_ds_scene.get()
        self._split_ds_tip.configure(
            text=(DS_SCENE_TEXT.get(scene, "") + "；拆完复检，还有完整号码就不打 ZIP") if on else
            "不打码：拆出去的表里是完整的号码。要打码就勾上（发群用就把用途选成「发群通报」）")

    def _split_desens_value(self):
        """拆分时的 config['desens']：没出现提示行、或没勾上都是空（不脱敏，行为和以前一样）。"""
        if not self._split_ds_cols or not self._split_ds_var.get():
            return ""
        return self._split_ds_scene.get()

    def _on_speed_change(self, label):
        """速度模式 → preserve_format，并联动「保留公式」（它必须先有保留格式）。"""
        std = (label != SPEED_FAST)
        self._preserve_var.set(std)
        if std:
            self._speed_hint.configure(
                text="完整保留表头、字体、颜色、列宽——收件人拿到的表和原表长得一样。"
                     "代价是逐格复制格式，文件多时会明显慢一些。")
        else:
            self._speed_hint.configure(
                text="只输出数据（纯文本/数字），不保留格式与公式，速度快很多。"
                     "适合对方只要数、不看排版的场景。")
        self._update_formula_state()

    def _update_formula_state(self):
        """极速模式下「保留公式」无从谈起（它依赖保留格式那条写入路径）：
        直接置灰并取消勾选，而不是留着一个勾了也不生效的框。"""
        box = getattr(self, "_keep_formula_box", None)
        if box is None:
            return
        if self._preserve_var.get():
            box.configure(state="normal")
        else:
            self._keep_formula_var.set(False)
            box.configure(state="disabled")

    # ── 自适应 / 折叠 ───────────────────────────────────
    def _update_input_ui(self):
        """输入类型跟随路径自动推断：是文件夹才显示批量选项。"""
        if os.path.isdir(self._input_var.get().strip()):
            self._batch_frame.grid()
        else:
            self._batch_frame.grid_remove()

    def _toggle_adv(self):
        self._adv_open = not self._adv_open
        if self._adv_open:
            self._adv_frame.grid(row=3, column=0, padx=8, pady=(0, 10), sticky="ew")
            self._adv_btn.configure(text="▾ 高级设置（一般用不到）")
        else:
            self._adv_frame.grid_remove()
            self._adv_btn.configure(text="▸ 高级设置（一般用不到）")

    def _on_header_mode(self, label):
        if label == "指定行号":
            self._hrow_frame.grid()
            self._kw_frame.grid_remove()
        elif label == "关键词":
            self._hrow_frame.grid_remove()
            self._kw_frame.grid()
        else:
            self._hrow_frame.grid_remove()
            self._kw_frame.grid_remove()

    # ── 选择输入 / 输出 ─────────────────────────────────
    def _pick_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if path:
            self._input_var.set(os.path.normpath(path))
            self._after_pick()

    def _pick_folder(self):
        path = filedialog.askdirectory()
        if path:
            self._input_var.set(os.path.normpath(path))
            self._after_pick()

    def _after_pick(self):
        self._update_input_ui()
        self._suggest_output()
        self._scan_input()

    def _on_path_edited(self):
        """手动改路径框（回车/失焦）后同步界面；路径没变就什么都不做。"""
        p = self._input_var.get().strip()
        if p == self._scanned_path:
            return
        self._update_input_ui()
        if p and os.path.exists(p):
            self._suggest_output()
            self._scan_input()
        elif p:
            self._scanned_path = p
            self._set_in_status("⚠ 找不到这个路径，请检查有没有写错", C_WARN)

    @staticmethod
    def _auto_output_for(p):
        """输入路径对应的默认输出目录；无法推导时返回空串。

        文件 → 同目录下「拆分结果」；文件夹 → 文件夹里面的「拆分结果」——结果永远出现在
        用户数据旁边，找得到、不污染上层目录。放输入里面是安全的：run_split 递归扫描时
        会整体跳过 output_root 子树（重复运行也不会把旧结果吃回去）。
        """
        if os.path.isdir(p):
            return os.path.join(os.path.normpath(p), "拆分结果")
        if os.path.isfile(p):
            return os.path.join(os.path.dirname(p), "拆分结果")
        return ""

    def _suggest_output(self):
        """按输入自动推荐输出目录；用户手动定过（当前值≠上次推荐值）就不再覆盖。"""
        cur = self._output_var.get().strip()
        if cur and cur != self._out_auto:
            return
        auto = self._auto_output_for(self._input_var.get().strip())
        if auto:
            self._output_var.set(auto)
            self._out_auto = auto

    # ── 扫描输入（文件数 + 字段），子线程执行 ───────────
    def _scan_input(self):
        p = self._input_var.get().strip()
        self._scanned_path = p
        if not p or not os.path.exists(p):
            self._set_in_status(self._idle_input_text() if not p
                                else "⚠ 找不到这个路径，请检查有没有写错",
                                C_MUTED if not p else C_WARN)
            return
        is_dir = os.path.isdir(p)
        self._set_in_status("正在读取…", C_MUTED)
        self._set_scan_status("正在识别表格里的字段…", C_MUTED)
        cfg = self._collect_config()
        out_now = self._output_var.get().strip()

        def work():
            # 计数镜像 run_split 的规则：输出目录子树整体跳过（结果放在输入里时数字才对）
            out_abs = os.path.normcase(os.path.abspath(out_now)) if out_now else None
            tpl, n, total_bytes = None, 0, 0
            if is_dir:
                for root, dirs, files in os.walk(p):
                    if out_abs:
                        ra = os.path.normcase(os.path.abspath(root))
                        if ra == out_abs or ra.startswith(out_abs + os.sep):
                            dirs[:] = []
                            continue
                    dirs.sort()         # Linux 上 os.walk 不排序，样本表要和 Windows 挑同一个
                    for fn in sorted(files):
                        if fn.lower().endswith(('.xlsx', '.xls')) and not fn.startswith('~$'):
                            n += 1
                            try:
                                total_bytes += os.path.getsize(os.path.join(root, fn))
                            except OSError:
                                pass
                            if tpl is None:
                                tpl = os.path.join(root, fn)
            else:
                tpl, n = p, 1
                try:
                    total_bytes = os.path.getsize(p)
                except OSError:
                    pass
            cols = []
            if tpl:
                try:
                    from core.splitter import list_columns
                    cols = list_columns(tpl, cfg)
                except Exception as e:
                    # 不静默吞掉：扫描失败原因（文件被占用/损坏等）打进日志，
                    # 否则界面只会显示笼统的「未能识别字段」，用户无从判断到底是
                    # 真的表头识别不出来，还是文件被 Excel/网盘同步占用等其它原因。
                    cols = []
                    self._ui_q.put(("log", f"⚠ 扫描「{os.path.basename(tpl)}」失败：{e}"))
            self._ui_q.put(("scan", (cols, n, os.path.basename(tpl) if tpl else "",
                                     is_dir, p, tpl or "", total_bytes)))

        threading.Thread(target=work, daemon=True).start()

    def _on_scan(self, payload):
        cols, n, tpl_name, is_dir, p, tpl_full, total_bytes = payload
        if p != self._input_var.get().strip():
            return    # 结果已过期（用户又换了输入）
        self._tpl_path = tpl_full or None    # 取值枚举复用这份模板文件
        # 本次扫描的规模先记下来：取值摘要要靠它说清「N 组来自 M 个表里的 1 个」，
        # 「换样本表」也要沿用这些数字（它只换模板，不重新统计整个文件夹）
        self._scan_n, self._scan_bytes, self._scan_is_dir = n, total_bytes, is_dir
        self._tpl_is_sample = bool(is_dir and n > 1)
        self._values_scope = "sample"
        self._scan_all_token += 1            # 让进行中的全量扫描结果作废
        if self._tpl_is_sample:
            self._tpl_btn.pack(side="left", padx=(6, 0))
        else:
            self._tpl_btn.pack_forget()
        # 换了输入，旧的取值预览随之作废
        self._values_all, self._value_vars = [], {}
        self._refresh_values_ui()
        mb = total_bytes / 1048576
        size_txt = f"，共 {mb:.0f} MB" if mb >= 1 else ""
        heavy = "（量较大，预计需要几分钟）" if (n > 50 or mb > 200) else ""
        name = os.path.basename(os.path.normpath(p)) or p
        if is_dir and n == 0:
            self._set_in_status(f"⚠ 文件夹「{name}」里没有找到 Excel 文件（.xlsx / .xls）", C_WARN)
            self._set_scan_status("", C_MUTED)
            return
        if is_dir:
            self._set_in_status(f"✓ 文件夹「{name}」：找到 {n} 个 Excel{size_txt}，全部一起拆{heavy}", C_OK)
        else:
            self._set_in_status(f"✓ 已选择：{name}{size_txt}{heavy}", C_OK)
        self._on_columns(cols, tpl_name)
        self._split_ds_cols = []
        self._scan_split_desens(tpl_full, p)
        self._rpt_hints = {"seq_column": "", "total_rows": 0}
        self._scan_report_hints(tpl_full, p)

    def _on_columns(self, cols, tpl_name=None):
        self._columns = cols
        if cols:
            self._split_menu.configure(values=cols)
            rec = self._recipe_active if (self._recipe_active or {}).get("mode") == "excel" else None
            if self._split_var.get() not in cols and not (rec and rec["settings"].get("split_column")):
                self._split_var.set(self._recommend_split(cols))
            self._pcol_menu.configure(values=cols)
            if self._pcol_var.get() not in cols and not (rec and rec["settings"].get("person_column")):
                self._pcol_var.set(self._recommend_person(cols))
            src = f"（来自「{tpl_name}」）" if tpl_name else ""
            warns = []
            if rec:
                from core.recipe import check_columns
                warns = check_columns(rec, cols)
            if rec and warns:
                self._set_scan_status("⚠ " + "；".join(warns) + "。请在下拉框里改选，或换一张表", C_WARN)
                for w in warns:
                    self._log(f"⚠️ 方案「{rec['name']}」：{w}")
            else:
                self._set_scan_status(f"✓ 识别到 {len(cols)} 个字段{src}，确认下拉框选的对不对", C_OK)
            self._on_split_col_change(None)   # 字段已定 → 自动枚举取值（v2.7 取值预览）
        else:
            self._warn_scan_failed()

    def _warn_scan_failed(self):
        """扫描不到字段时的提示：「关键词」模式下关键词没匹配上是个常见且好定位的
        原因（比如换了一批表头写法不同的表格，还留着上次填的旧关键词），单独
        点名出来，比统一甩一句「未能识别字段」让用户少走弯路。"""
        if self._header_seg.get() == "关键词":
            kws = [k for k in (self._grid_keys_var.get().strip(), self._id_keys_var.get().strip()) if k]
            if kws:
                if not self._adv_open:
                    self._toggle_adv()
                self._set_scan_status(
                    f"未能识别字段：当前「表头识别」用的是「关键词」模式，但表格前几行里"
                    f"没找到关键词「{' / '.join(kws)}」——已展开下方「高级设置」，"
                    f"核对这两个关键词是否适用于这份表格，或改回「自动」识别", C_WARN)
                return
        self._set_scan_status("未能识别字段：点「🔄 重新识别」重试，或到「高级设置」手动指定表头行", C_WARN)

    @staticmethod
    def _recommend_split(cols):
        """猜一个最像「分组」的字段做默认值，猜不到就用第一个（用户仍需自己确认）。"""
        for kw in ("部门", "网格", "区域", "分公司", "门店", "班组", "科室", "组织", "单位", "类别", "分类"):
            for c in cols:
                if kw in str(c):
                    return c
        return cols[0]

    @staticmethod
    def _recommend_person(cols):
        for kw in ("姓名", "名字", "工号", "负责人", "经办"):
            for c in cols:
                if kw in str(c):
                    return c
        return cols[0]

    def _find_stale_results(self, input_dir, output_root):
        """输入文件夹里「不被本次输出目录覆盖」的历史拆分结果目录。

        结果默认放在输入里面、由 run_split 按 output_root 跳过；但如果用户后来把输出
        改到了别处，留在输入里的旧结果就会被当成数据拆一遍（静默重复）。开跑前找出来确认。
        """
        out = os.path.normcase(os.path.abspath(output_root))
        hits = []
        for root, dirs, _ in os.walk(input_dir):
            for d in list(dirs):
                full = os.path.normcase(os.path.abspath(os.path.join(root, d)))
                if full == out or full.startswith(out + os.sep):
                    dirs.remove(d)    # 本次输出覆盖的子树，run_split 会跳过，不算
                    continue
                if d == "拆分结果" or d.endswith("_拆分结果") or re.fullmatch(r"\d{8}结果(_\d+)?", d):
                    hits.append(os.path.join(root, d))
                    dirs.remove(d)    # 命中即整树算一个，不再往里翻
        return hits

    @staticmethod
    def _find_open_locks(path, cap=4):
        """检测输入里正被 Excel/WPS 打开的文件（存在同名锁文件）。

        两种锁文件格式都认：Windows 的 `~$名字.xlsx`，和 Linux 版 WPS /
        LibreOffice 的 `.~lock.名字.xlsx#`（见 platform_compat.lock_owner_name）——
        只认前者的话，这条预警在麒麟上永远不会触发。

        返回原文件名列表（最多 cap 个）；目录扫描限前 2000 个目录，防止巨型目录卡界面。
        """
        hits = []
        if os.path.isfile(path):
            d, base = os.path.dirname(path), os.path.basename(path)
            if any(os.path.exists(os.path.join(d, cand))
                   for cand in ("~$" + base, ".~lock." + base + "#")):
                hits.append(base)
            return hits
        n_walk = 0
        for _root, _dirs, files in os.walk(path):
            n_walk += 1
            if n_walk > 2000 or len(hits) >= cap:
                break
            names = {f for f in files if f.lower().endswith(('.xlsx', '.xls'))}
            locked = {n for n in (lock_owner_name(f) for f in files) if n}
            for f in sorted(locked & names):
                hits.append(f)
                if len(hits) >= cap:
                    break
        return hits

    def _set_in_status(self, text, color=C_MUTED):
        self._in_status.configure(text=text, text_color=color)

    def _set_scan_status(self, text, color=C_MUTED):
        self._status.configure(text=text, text_color=color)

    def _parse_alias(self):
        raw = self._alias_box.get("1.0", "end").strip()
        if not raw:
            return {}, True
        try:
            obj = json.loads(raw)
            if isinstance(obj, dict):
                return obj, True
        except Exception:
            pass
        return {}, False

    def _start(self):
        if self._mode == "pdf":
            return self._start_pdf()
        if self._mode == "collect":
            return self._start_collect("check" if self._col_check_var.get() else "merge")
        if self._mode == "desens":
            return self._start_desens()
        cfg = self._collect_config()
        cfg["desens"] = cfg["split_desens"]       # 拆分时顺带脱敏：空 = 不脱敏，和以前一样
        if not cfg["input_path"] or not os.path.exists(cfg["input_path"]):
            messagebox.showwarning("先选择表格", "请先在第 ① 步选择要拆的 Excel 文件或文件夹。")
            return
        if not cfg["split_column"]:
            messagebox.showwarning("先选择拆分字段",
                                   "请在第 ② 步选择「拆分字段」——按哪一列拆分。\n如果下拉框还没有内容，点「🔄 扫描字段」。")
            return
        if not cfg["output_path"]:
            self._suggest_output()
            cfg["output_path"] = self._output_var.get().strip()
        if not cfg["output_path"]:
            messagebox.showwarning("先选择保存位置", "请在第 ③ 步选择结果保存到哪个文件夹。")
            return
        # 输出目录若等于输入文件夹本身或在它上层，核心扫描时会把全部文件当输出跳过 → 提前拦住
        if os.path.isdir(cfg["input_path"]):
            try:
                inp = os.path.normcase(os.path.abspath(cfg["input_path"]))
                out = os.path.normcase(os.path.abspath(cfg["output_path"]))
                if inp == out or inp.startswith(out + os.sep):
                    messagebox.showwarning("换个保存位置",
                                           "结果保存位置不能选输入文件夹自己或它的上层文件夹，\n否则找不到要拆的文件。建议直接用自动推荐的位置。")
                    return
            except Exception:
                pass
            # 输入里有历史拆分结果、而本次输出没盖住它 → 会被当数据重复拆，先问一声
            try:
                stale = self._find_stale_results(cfg["input_path"], cfg["output_path"])
            except Exception:
                stale = []
            if stale:
                shown = "\n".join(stale[:3]) + ("\n…" if len(stale) > 3 else "")
                if not messagebox.askyesno(
                        "发现以前的拆分结果",
                        f"输入文件夹里有以前拆出来的结果：\n{shown}\n\n"
                        "这次运行会把里面的文件也当成数据一起拆，内容可能重复。\n"
                        "建议先把它们移走/删除，或把保存位置改回默认。\n\n仍要继续吗？"):
                    return
        # ★ 抽样提醒（v2.9）：②里的「N 组」只来自一个样本表，而这一跑是对全部
        # 文件生效的——不筛选取值时真实分组数几乎总是更多。让用户在这里就知道，
        # 而不是跑完 20 分钟看到 29 个分组才发现和预览的 6 组不是一回事。
        if (self._tpl_is_sample and self._values_scope != "all"
                and not cfg["selected_values"] and self._values_all):
            ans = messagebox.askyesnocancel(
                "分组数可能比预览的多",
                f"第 ② 步显示的 {len(self._values_all)} 个分组只来自样本表"
                f"「{os.path.basename(self._tpl_path or '')}」，\n"
                f"这次要拆的是全部 {self._scan_n} 个表，其他表里可能还有别的分组，"
                "实际拆出的会更多。\n\n"
                "· 「是」＝照这样拆全部分组\n"
                "· 「否」＝先读遍所有表列全分组，我再勾选要哪几个\n"
                "· 「取消」＝什么都不做")
            if ans is None:
                return
            if not ans:
                self._scan_all_values()
                return
        # 误选保护（A1）：取值预览发现分组异常多时，开跑前再确认一次
        if len(self._values_all) > 50:
            if not messagebox.askyesno(
                    "分组特别多",
                    f"将按「{cfg['split_column']}」拆出 {len(self._values_all)} 个分组（文件）。\n"
                    "取值异常多常见于选错了字段（比如选到工号 / 姓名列）。\n\n仍要继续吗？"):
                return
        # 占用预检（B2）：输入里被 Excel/WPS 打开的文件（~$ 锁），读到的是未保存的旧内容
        locks = self._find_open_locks(cfg["input_path"])
        if locks:
            shown = "、".join(locks[:3]) + (" 等" if len(locks) > 3 else "")
            if not messagebox.askyesno(
                    "文件正被打开",
                    f"检测到以下文件正被 Excel/WPS 打开：{shown}\n"
                    "拆分读到的是「打开前保存的内容」，刚改的数不会生效。\n\n建议先关闭这些文件。仍要继续吗？"):
                return
        # 保存位置提前建好：位置只读/无权限时现在就报，而不是跑到一半失败
        try:
            os.makedirs(cfg["output_path"], exist_ok=True)
        except OSError as e:
            messagebox.showerror("无法创建保存位置",
                                 f"没法在这里创建结果文件夹：\n{cfg['output_path']}\n\n{e}\n\n"
                                 "请点「浏览」换一个能写入的位置（比如桌面或文档）。")
            return
        _, alias_ok = self._parse_alias()
        if not alias_ok:
            messagebox.showwarning("高级设置有误", "「高级设置 → 取值归并映射」不是合法 JSON，请修正或清空。")
            return

        self._enter_running("⏳ 正在拆分…", "已开始，正在扫描输入…")

        def run():
            try:
                from core.splitter import run_split
                output_path = run_split(
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
