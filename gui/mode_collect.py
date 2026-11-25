"""
ExcelRouter - 图形界面：「Excel 汇总」模式。

①文件夹 / 一个文件 ②先查收再汇总 + 结果卡片，以及开始汇总 / 查收。

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

from gui.platform_compat import open_folder
from gui.common import (
    COL_IN_TIP, EXPECT_TIP, C_OK, C_WARN, C_ERR, C_MUTED, PRIMARY, PRIMARY_HOVER, OK_BANNER_BG,
    CARD_BORDER, _ghost_button, _flat_button, save_config,
)


class CollectModeMixin:
    # ── Excel 汇总模式（v2.10）：默认就是普通汇总；② 打开后先查收（谁没交、谁漏填）再汇总 ──
    def _build_collect_steps(self, parent):
        card, _ = self._step_card(parent, 0, 1, "选择要汇总的表格")
        btns = ctk.CTkFrame(card, fg_color="transparent")
        btns.grid(row=1, column=0, padx=12, pady=(6, 2), sticky="w")
        ctk.CTkButton(btns, text="📁 选文件夹", width=150,
                      command=self._pick_collect_folder).pack(side="left")
        _ghost_button(btns, text="📄 选一个文件", width=130,
                      command=self._pick_collect_file).pack(side="left", padx=(10, 0))
        self._col_in_var = ctk.StringVar(value=self.cfg.get("collect_input_path", ""))
        entry = ctk.CTkEntry(card, textvariable=self._col_in_var,
                             placeholder_text="也可以把文件夹直接拖到这里")
        entry.grid(row=2, column=0, padx=12, pady=(6, 2), sticky="ew")
        entry.bind("<Return>",   lambda e: self._scan_collect())
        entry.bind("<FocusOut>", lambda e: self._scan_collect())
        self._enable_drop(card, self._on_drop_collect)
        self._enable_drop(entry, self._on_drop_collect)
        self._col_in_status = ctk.CTkLabel(
            card, text=COL_IN_TIP, font=ctk.CTkFont(size=11), text_color=C_MUTED,
            anchor="w", justify="left", wraplength=640)
        self._col_in_status.grid(row=3, column=0, padx=12, pady=(0, 10), sticky="w")

        card2, _ = self._step_card(parent, 1, 2, "查谁没交、谁漏填（选填）")
        self._col_check_var = ctk.BooleanVar(value=bool(self.cfg.get("collect_check")))
        ctk.CTkSwitch(card2, text="这些表是发出去让大家填的：先查收，再汇总", variable=self._col_check_var,
                      command=self._on_check_toggle, progress_color=PRIMARY).grid(
            row=1, column=0, padx=12, pady=(6, 0), sticky="w")
        ctk.CTkLabel(card2, text="打开后点「查收」：几秒列出谁没交、谁交了空表、谁哪一列漏了哪几行，"
                                 "生成能直接发到群里的通报；不打开就是普通的合并汇总",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED, anchor="w", justify="left",
                     wraplength=640).grid(row=2, column=0, padx=12, pady=(2, 10), sticky="w")
        form = self._col_form = ctk.CTkFrame(card2, fg_color="transparent")
        form.grid(row=3, column=0, padx=12, pady=(0, 10), sticky="ew")
        form.grid_columnconfigure(1, weight=1)

        def field(box, row, label, key, hint, width=None):
            ctk.CTkLabel(box, text=label).grid(row=row * 2, column=0, sticky="w", pady=(4, 0))
            var = ctk.StringVar(value=str(self.cfg.get(key, "") or ""))
            ctk.CTkEntry(box, textvariable=var, width=width or 140).grid(
                row=row * 2, column=1, padx=(8, 0), pady=(4, 0), sticky="w" if width else "ew")
            tip = ctk.CTkLabel(box, text=hint, font=ctk.CTkFont(size=11), text_color=C_MUTED,
                               anchor="w", justify="left", wraplength=520)
            tip.grid(row=row * 2 + 1, column=1, padx=(8, 0), sticky="w")
            return var, tip

        self._col_expect_var, self._col_expect_tip = field(form, 0, "应交名单", "collect_expect", EXPECT_TIP)
        _ghost_button(form, text="从表格读", width=80,
                      command=self._pick_collect_names).grid(row=0, column=2, padx=(8, 0), pady=(4, 0))
        self._col_due_var, _ = field(form, 1, "截止时间", "collect_due",
                                     "写进通报和催办话术，比如「10月15日 18:00」", width=220)
        self._col_more_btn = _flat_button(form, text="▸ 更多设置", width=90, anchor="w",
                                          command=self._toggle_col_more)
        self._col_more_btn.grid(row=4, column=0, columnspan=3, pady=(6, 0), sticky="w")
        more = self._col_more = ctk.CTkFrame(form, fg_color="transparent")
        more.grid(row=5, column=0, columnspan=3, sticky="ew")
        more.grid_columnconfigure(1, weight=1)
        self._col_req_var, _ = field(more, 0, "必须填的列", "collect_required",
                                     "留空自动定：有下发记录就用下发时的填写列，没有就看大多数人都填了的列")
        self._col_by_var, _ = field(more, 1, "按哪一列分人", "collect_by",
                                    "只在 ① 选了一个文件（大家各填各行的一张总表）时用，比如「网格」「客户经理」",
                                    width=180)
        self._col_title_var, _ = field(more, 2, "通报标题", "collect_title", "留空用文件夹名", width=220)
        self._col_rec_var, _ = field(more, 3, "下发记录", "collect_record",
                                     "用本工具拆分、并填了「让对方填写的列」才有，能查出谁改了不该改的；刚拆过会自动带上")
        _ghost_button(more, text="选择", width=80,
                      command=self._pick_collect_record).grid(row=6, column=2, padx=(8, 0), pady=(4, 0))
        more.grid_remove()
        if not self._col_check_var.get():
            form.grid_remove()

        # 结果卡片：不编号——它是③的结果，不是一个步骤。查收后显示通报，汇总后显示汇总说明
        res = self._col_result = ctk.CTkFrame(parent, corner_radius=12, border_width=1, border_color=CARD_BORDER)
        res.grid(row=2, column=0, padx=8, pady=(0, 10), sticky="ew")
        res.grid_columnconfigure(0, weight=1)
        self._col_result_title = ctk.CTkLabel(res, text="", font=ctk.CTkFont(size=14, weight="bold"))
        self._col_result_title.grid(row=0, column=0, padx=14, pady=(10, 2), sticky="w")
        self._col_notice_box = ctk.CTkTextbox(res, height=200, wrap="word", font=ctk.CTkFont(size=12))
        self._col_notice_box.grid(row=1, column=0, padx=12, pady=(4, 6), sticky="ew")
        self._col_notice_box.configure(state="disabled")
        self._col_check_btns = ctk.CTkFrame(res, fg_color="transparent")
        self._col_check_btns.grid(row=2, column=0, padx=12, pady=(0, 12), sticky="w")
        ctk.CTkButton(self._col_check_btns, text="📋 复制通报", width=110, fg_color=PRIMARY,
                      hover_color=PRIMARY_HOVER, command=self._copy_notice).pack(side="left")
        _ghost_button(self._col_check_btns, text="📄 打开进度表", width=110,
                      command=self._open_collect_progress).pack(side="left", padx=(10, 0))
        self._col_merge_btn = _ghost_button(self._col_check_btns, text="📊 汇总已交的表", width=130,
                                            command=lambda: self._start_collect("merge"))
        self._col_merge_btn.pack(side="left", padx=(10, 0))
        self._col_merged_btns = ctk.CTkFrame(res, fg_color="transparent")
        self._col_merged_btns.grid(row=2, column=0, padx=12, pady=(0, 12), sticky="w")
        ctk.CTkButton(self._col_merged_btns, text="📄 打开汇总表", width=120, fg_color=PRIMARY,
                      hover_color=PRIMARY_HOVER, command=self._open_merged).pack(side="left")
        # 要往上报、发给外单位：另存一份个人信息清空的脱敏版（v3.1）
        self._col_desens_btn = _ghost_button(self._col_merged_btns, text="🛡 另存脱敏版（上报用）", width=170,
                                             command=self._desens_merged)
        self._col_desens_btn.pack(side="left", padx=(10, 0))
        self._col_merged_btns.grid_remove()
        res.grid_remove()
        p = self._col_in_var.get().strip()
        if p and os.path.exists(p):
            self._scan_collect()

    def _on_check_toggle(self):
        """② 的开关：开 = 查收再汇总，关 = 普通汇总。主按钮文案和提示跟着变。"""
        if self._col_check_var.get():
            self._col_form.grid()
        else:
            self._col_form.grid_remove()
        self._col_result.grid_remove()        # 上一次的结果和现在的模式可能对不上，先收起
        self._col_info.pop("auto_on", None)   # 用户自己拨过开关，「已替你打开」的提示就不对了
        if self._mode == "collect":
            self._apply_mode()
        self._refresh_collect_status()

    def _toggle_col_more(self):
        if self._col_more.winfo_ismapped():
            self._col_more.grid_remove()
            self._col_more_btn.configure(text="▸ 更多设置")
        else:
            self._col_more.grid()
            self._col_more_btn.configure(text="▾ 更多设置")

    def _pick_collect_folder(self):
        p = filedialog.askdirectory()
        if p:
            self._col_in_var.set(os.path.normpath(p))
            self._scan_collect()

    def _pick_collect_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if p:
            self._col_in_var.set(os.path.normpath(p))
            self._scan_collect()

    def _on_drop_collect(self, event):
        paths = self._parse_drop_paths(getattr(event, "data", ""))
        if self._drop_recipe(paths):
            return
        rec = next((p for p in paths if p.lower().endswith(".json")), None)
        if rec:                                    # 不是方案的 .json：当下发记录用
            self._col_rec_var.set(os.path.normpath(rec))
            self._col_check_var.set(True)
            self._on_check_toggle()
            self._col_scanned = None
            self._scan_collect()
            return
        target = next((p for p in paths
                       if os.path.isdir(p) or p.lower().endswith((".xlsx", ".xlsm", ".xls"))), None)
        if target is None:
            self._col_in_status.configure(text="⚠ 请拖入装表格的文件夹，或一个 Excel 文件", text_color=C_WARN)
            return
        self._col_in_var.set(os.path.normpath(target))
        self._scan_collect()

    def _pick_collect_record(self):
        p = filedialog.askopenfilename(filetypes=[("下发记录", "*.json")])
        if p:
            self._col_rec_var.set(os.path.normpath(p))
            self._col_scanned = None
            self._scan_collect()

    def _pick_collect_names(self):
        """从一张名单表里读应交名单（表头含「姓名 / 网格 / 单位 …」的列，没有就取第一列）。"""
        p = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if not p:
            return
        try:
            from core.collect import read_name_list
            names = read_name_list(p)
        except Exception as e:
            messagebox.showerror("读不了名单", f"这张表读不了：{e}")
            return
        if not names:
            messagebox.showwarning("没读到名字", "这张表里没读到名字，请确认第一列或「姓名」列里有内容。")
            return
        self._col_expect_var.set("、".join(names))
        self._col_expect_tip.configure(text=f"✓ 从「{os.path.basename(p)}」读到 {len(names)} 个名字",
                                       text_color=C_OK)

    def _scan_collect(self):
        """子线程清点：几个表、有没有下发记录 / 名单表（经 UI 泵回填，机制同 _scan_input）。"""
        p = self._col_in_var.get().strip()
        if not p or p == self._col_scanned:
            return
        self._col_scanned = p
        if not os.path.exists(p):
            self._col_info = {}
            self._col_in_status.configure(text="⚠ 找不到这个文件夹或文件，请检查路径", text_color=C_WARN)
            return
        self._col_in_status.configure(text="正在清点…", text_color=C_MUTED)
        chosen = self._col_rec_var.get().strip()

        def work():
            info: dict[str, Any] = {"path": p}
            try:
                from core import collect
                base = p if os.path.isdir(p) else os.path.dirname(p)
                files, lists = collect.returned_files(p, os.path.join(base, collect.RESULT_DIR))
                info["n"] = len(files) + (len(lists) if os.path.isdir(p) else 0)
                info["lists"] = [os.path.basename(x) for x in lists]
                near = collect.find_record(p)
                rp = chosen if chosen and os.path.isfile(chosen) else near
                rec = collect.load_record(rp) if rp else None
                if rec:
                    stems = [os.path.splitext(os.path.basename(f[1]))[0] for f in files]
                    info["record"] = {
                        "path": rp, "near": rp == near, "title": rec.get("title", ""),
                        "created": str(rec.get("created", ""))[5:10], "names": list(rec["groups"]),
                        # 用文件名粗看一下这份记录和这个文件夹是不是一批（改了名的认不出，只用来提示）
                        "hits": sum(1 for g in rec["groups"] if any(g in s for s in stems))}
            except Exception as e:
                info["error"] = str(e)
            self._ui_q.put(("collect_scan", info))

        threading.Thread(target=work, daemon=True).start()

    def _on_collect_scan(self, info):
        if info["path"] != self._col_in_var.get().strip():
            return    # 结果已过期
        self._col_info = info
        rec = info.get("record")
        # 检测到这一批的下发记录或名单表：多半是在收大家填的表，替用户打开 ②
        if (rec and (rec["near"] or rec["hits"]) or info.get("lists")) and not self._col_check_var.get():
            self._col_check_var.set(True)
            self._on_check_toggle()
            info["auto_on"] = "下发记录" if rec else "名单表"
        self._refresh_collect_status()

    def _refresh_collect_status(self):
        """① 的状态行和 ② 应交名单的提示：随清点结果和 ② 的开关一起变。"""
        info = self._col_info
        if not info:
            return
        if info.get("error"):
            self._col_in_status.configure(text=f"⚠ 读不了这个位置：{info['error']}", text_color=C_WARN)
            return
        check = self._col_check_var.get()
        name = os.path.basename(info["path"])
        if os.path.isfile(info["path"]):
            text = (f"✓ 一张总表「{name}」：在 ②「更多设置」里填按哪一列分人" if check
                    else f"✓ 会把「{name}」里的各个 sheet 合成一张")
            ok = True
        elif info.get("n"):
            text, ok = f"✓ 找到 {info['n']} 个表，汇总结果放在这个文件夹里的「汇总结果」", True
        else:
            text, ok = "⚠ 文件夹里还没有表格（.xlsx / .xls）", False
        if info.get("auto_on"):
            text += f"；有这一批的{info['auto_on']}，已替你打开第 ② 步的查收"
        self._col_in_status.configure(text=text, text_color=C_OK if ok else C_WARN)
        rec = info.get("record")
        tip, color = EXPECT_TIP, C_MUTED
        if rec:
            shown = "、".join(rec["names"][:4]) + (" 等" if len(rec["names"]) > 4 else "")
            label = f"「{rec['title']}」" if rec["title"] else ""
            tip = f"留空就按下发记录{label}（{rec['created']} 拆分）：应交 {len(rec['names'])} 个，{shown}"
            color = C_OK
            if not rec["near"] and not rec["hits"]:
                tip = f"⚠ 下发记录{label}的分组名和这个文件夹里的文件名对不上，可能不是这一批；在「更多设置」里换掉或清空"
                color = C_WARN
        elif info.get("lists"):
            tip, color = "留空就用文件夹里的名单表：" + "、".join(info["lists"]), C_OK
        if not self._col_expect_var.get().strip():
            self._col_expect_tip.configure(text=tip, text_color=color)

    def _start_collect(self, action):
        """action：merge（普通汇总，或查收后汇总已交的表）/ check（查收）。结果写进输入旁的「汇总结果」。"""
        cfg = self._collect_config()
        p = cfg["collect_input_path"]
        check = cfg["collect_check"]
        if not p or not os.path.exists(p):
            messagebox.showwarning("先选择表格", "请先在第 ① 步选择装表格的文件夹（或一个 Excel 文件）。")
            return
        if check and os.path.isfile(p) and not cfg["collect_by"]:
            messagebox.showwarning("填一下按哪一列分人",
                                   "选的是一张总表：请在第 ②「更多设置」里填「按哪一列分人」（比如 网格、客户经理），\n"
                                   "我按这一列的取值逐人查。\n\n只想把它的各个 sheet 合成一张的话，关掉第 ② 步的开关就行。")
            return
        locks = self._find_open_locks(p)
        if locks:
            shown = "、".join(locks[:3]) + (" 等" if len(locks) > 3 else "")
            if not messagebox.askyesno(
                    "文件正被打开",
                    f"检测到以下文件正被 Excel/WPS 打开：{shown}\n"
                    "读到的是「打开前保存的内容」，刚改的不会算进去。\n\n建议先关闭这些文件。仍要继续吗？"):
                return
        run_cfg = dict(cfg, input_path=p, output_path="", log_notice=True)   # 日志框里直接显示通报
        if action == "check":
            self._enter_running("⏳ 正在查收…", "已开始，正在清点收回来的表…")
        else:
            self._enter_running("⏳ 正在汇总…", "已开始，先查收一遍再汇总…" if check else "已开始，正在读取表格…")
        self._col_merge_btn.configure(state="disabled")
        self._col_result.grid_remove()

        def run():
            try:
                from core.collect import check_collect, merge_tables, run_collect
                fn = check_collect if action == "check" else (run_collect if check else merge_tables)
                output_path = fn(run_cfg,
                                 log_fn=lambda m: self._ui_q.put(("log", m)),
                                 progress_fn=lambda v: self._ui_q.put(("progress", v)),
                                 stop_flag=lambda: self._stop_flag)
            except Exception as e:
                self._ui_q.put(("log", f"\n❌ 运行出错：{e}"))
                output_path = None
            self._ui_q.put(("done", (dict(cfg, _collect_action=action), output_path)))

        threading.Thread(target=run, daemon=True).start()

    def _show_collect_result(self, title, text, checked):
        self._col_result_title.configure(text=title)
        # 框高随内容：汇总说明常常只有三四行，固定高度会空出一大块
        self._col_notice_box.configure(height=min(240, 30 + 20 * (text.count("\n") + 1)))
        self._col_notice_box.configure(state="normal")
        self._col_notice_box.delete("1.0", "end")
        self._col_notice_box.insert("1.0", text)
        self._col_notice_box.configure(state="disabled")
        if checked:
            self._col_merged_btns.grid_remove()
            self._col_check_btns.grid()
        else:
            self._col_check_btns.grid_remove()
            self._col_merged_btns.grid()
        self._col_result.grid()
        self.after(150, lambda: self._body._parent_canvas.yview_moveto(1.0))

    def _on_collect_done(self, cfg, output_path):
        """Excel 汇总完成：查收 → 通报 + 进度表；汇总 → 汇总说明 + 打开汇总表。"""
        self._col_merge_btn.configure(state="normal")
        s = self._last_summary or {}
        action = cfg.get("_collect_action", "merge")
        if self._stop_flag or s.get("stopped"):
            self._run_status.configure(text="⏹ 已停止，这一轮没有生成结果", text_color=C_WARN)
            return
        if not output_path:
            if s.get("mode") in ("collect", "merge"):      # 正常跑完，只是没有能合并的数据
                self._run_status.configure(text="⚠ 没有生成汇总表：没有能合并的数据（详见处理详情）",
                                           text_color=C_WARN)
                self._toggle_log(show=True)
                return
            self._run_status.configure(
                text="❌ 没有完成：原因见下方「处理详情」；可点「复制日志」连同反馈一起发给作者",
                text_color=C_ERR)
            self._copylog_btn.grid()
            self._toggle_log(show=True)
            return
        self._last_output = output_path if os.path.isdir(output_path) else os.path.dirname(output_path)
        self._open_btn.grid()
        try:
            save_config({k: v for k, v in cfg.items() if not k.startswith("_")})
        except Exception:
            pass
        lines, warn = [], False
        if s.get("mode") in ("collect_check", "collect"):
            exp, sub, ok = s.get("expected"), s.get("submitted", 0), s.get("ok", 0)
            done_all = (ok == exp) if exp is not None else (sub and ok == sub)
            warn = not done_all
            problems = [f"{lab} {s[k]}" for k, lab in (("missing", "未交"), ("empty", "空表"), ("partial", "漏填"),
                                                       ("invalid", "填写有误"), ("changed", "有改动"), ("bad", "没法查"), ("extra", "名单外"))
                        if s.get(k)]
            head = f"📊 应交 {exp} · 已交 {sub} · 合格 {ok}" if exp is not None else f"📊 已收 {sub} 份 · 合格 {ok} 份"
            lines.append(head + ("（" + " · ".join(problems) + "）" if problems else ""))
        if action == "check":
            text = ""
            try:
                with open(s.get("notice") or "", encoding="utf-8-sig") as f:
                    text = f.read()
            except OSError:
                pass
            self._col_notice = text
            self._col_progress = s.get("progress", "")
            self._col_merge_btn.configure(text="📊 汇总已交的表" if warn else "📊 汇总成一张")
            self._show_collect_result("📣 查收结果（可直接复制到群里）",
                                      text or "（通报没写成文件，请看下方「处理详情」）", True)
            self._run_status.configure(
                text="✅ 查收完成：通报在下方，点「复制通报」发到群里" if warn else "✅ 全部交齐、填写完整，可以汇总了",
                text_color=C_OK)
            lines.append(f"📁 {self._last_output}")
        else:
            self._col_merged = output_path
            self._col_notice = s.get("report_text", "")
            self._show_collect_result("✅ 汇总完成", self._col_notice or "（详见处理详情）", False)
            self._run_status.configure(text="✅ 汇总完成！点「打开汇总表」查看", text_color=C_OK)
            lines.insert(0, f"📄 汇总 {s.get('files', '?')} 个表 · {s.get('rows', '?')} 行 → {os.path.basename(output_path)}")
            lines.append(f"📁 {self._last_output}")
            self._log("💬 用得顺手或踩了坑？点「反馈建议」匿名告诉作者（1 分钟）。")
        self._summary_lbl.configure(text="\n".join(lines), text_color=C_WARN if warn else C_OK)
        self._summary_banner.configure(fg_color=("gray90", "gray20") if warn else OK_BANNER_BG)
        self._show_summary_banner()
        if self._warn_count:
            self._log_btn.configure(text=self._log_btn_text())

    def _copy_notice(self):
        if not self._col_notice:
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(self._col_notice)
            self._run_status.configure(text="📋 通报已复制，可直接粘贴到群里；逐人催办在通报末尾", text_color=C_OK)
        except Exception:
            pass

    def _open_collect_progress(self):
        p = self._col_progress
        if not p or not os.path.exists(p):
            self._run_status.configure(text="⚠ 进度表不在了，请重新查收一次", text_color=C_WARN)
            return
        if not open_folder(p):
            self._run_status.configure(text=f"⚠ 打不开进度表，文件在：{p}", text_color=C_WARN)

    def _open_merged(self):
        p = self._col_merged
        if not p or not os.path.exists(p):
            self._run_status.configure(text="⚠ 汇总表不在了，请重新汇总一次", text_color=C_WARN)
            return
        if not open_folder(p):
            self._run_status.configure(text=f"⚠ 打不开汇总表，文件在：{p}", text_color=C_WARN)

    @staticmethod
    def _recommend_col(cols, keywords):
        for kw in keywords:
            for c in cols:
                if kw.lower() in str(c).lower():
                    return c
        return None
