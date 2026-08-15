"""
ExcelRouter - 图形界面

三步式单窗口布局（面向非技术办公用户，按任务顺序组织界面）：
  ① 选择要拆的表格：「选文件 / 选文件夹」两个直选按钮（类型由路径自动推断），
     选完自动统计文件数、后台扫描字段、自动推荐输出目录；
  ② 按哪个字段拆分：拆分字段下拉（扫描后按关键词智能预选）＋批量选项
     （仅文件夹显示：每组 ZIP / 同时拆到人）；
  ③ 开始拆分：输出目录＋主按钮＋进度＋状态行，固定在窗口底部随时可达。
  「▸ 高级设置」折叠在步骤区末尾；「▸ 处理详情」日志默认收起，出错自动展开。

拆分在子线程里跑，日志/进度经 queue 由主线程 UI 泵（_pump_ui，每 100ms 批量刷新）
更新界面——子线程绝不直接碰 Tk 控件（v2.4 机制，勿回退）。

Copyright (c) 2026 AbeLin
MIT License
"""

import os
import re
import sys
import json
import queue
import threading
import subprocess
import webbrowser
from typing import Any

import customtkinter as ctk
from tkinter import filedialog, messagebox


# =====================================================
# 路径解析（兼容 PyInstaller --onefile 打包）
# =====================================================

def _is_frozen():
    return getattr(sys, 'frozen', False)


def _resource_dir():
    if _is_frozen():
        return sys._MEIPASS   # type: ignore[attr-defined]  # PyInstaller 运行时注入
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _app_dir():
    if _is_frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


DEFAULT_CONFIG_PATH = os.path.join(_resource_dir(), "config", "default_config.json")
USER_CONFIG_PATH    = os.path.join(_app_dir(), "user_config.json")

COL_PLACEHOLDER = "（选择表格后自动识别）"
PDF_COL_PLACEHOLDER = "（选清单后自动识别）"
PDF_RECV_NONE = "（不需要）"

# UI 泵用的「本轮没有此类消息」哨兵（不能用 None：扫描失败时 payload 可能为空）
_MISSING = object()

APP_VERSION = "2.6.0"
# 匿名反馈问卷地址（问卷 URL 确定后替换此处即可，一行改动 + 打 tag 发版）
FEEDBACK_URL = "https://f.wps.cn/g/pBOAWUQc/"

# 状态文字颜色（浅色模式, 深色模式）
C_OK    = ("#15803d", "#4ade80")
C_WARN  = ("#b45309", "#fbbf24")
C_ERR   = ("#b91c1c", "#f87171")
C_MUTED = ("gray40", "gray60")
ACCENT  = ("#3B8ED0", "#1F6AA5")   # 与 ctk blue 主题一致，步骤徽章/主按钮同源

def _ghost_button(parent, **kw):
    """描边次要按钮：比主按钮弱一级的操作（浏览 / 扫描字段 / 打开文件夹等）。"""
    return ctk.CTkButton(parent, fg_color="transparent", border_width=1,
                         border_color=("gray60", "gray40"),
                         text_color=("gray15", "gray85"),
                         hover_color=("gray90", "gray25"), **kw)


def _flat_button(parent, **kw):
    """纯文字按钮：折叠开关、页脚工具位等最低视觉权重的操作。"""
    return ctk.CTkButton(parent, fg_color="transparent",
                         text_color=("gray20", "gray80"),
                         hover_color=("gray85", "gray25"), **kw)

# 兜底默认配置（通用、不绑定任何业务）
FALLBACK_CONFIG = {
    "input_path": "", "output_path": "",
    "header_mode": "auto", "header_row": 1,
    "grid_keys": [], "id_keys": [],
    "split_column": "", "selected_values": [],
    "person_column": "", "to_person": False, "person_file_filter": [],
    "value_alias_map": {},
    "skip_values": ["合计", "小计", "总计", "平均", ""],
    "merge_across_files": False,
    "make_zip": True,
    "exact_match": True,
    "preserve_format": True,
    "keep_formulas": False,
    "auto_open_output": True,
    # ── PDF 加密分发（v2.6）──
    "ui_mode": "excel",
    "pdf_input_paths": [],
    "pdf_mapping_path": "",
    "pdf_grid_column": "",
    "pdf_password_column": "",
    "pdf_receiver_column": "",
    "pdf_watermark": True,
    "pdf_watermark_text": "{grid} {date}",
    "pdf_watermark_opacity": 0.15,
    "pdf_watermark_angle": 45,
}

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


def load_config():
    for path in (USER_CONFIG_PATH, DEFAULT_CONFIG_PATH):
        if os.path.exists(path):
            try:
                with open(path, encoding='utf-8') as f:
                    cfg = json.load(f)
                merged = dict(FALLBACK_CONFIG)
                merged.update(cfg)
                return merged
            except Exception:
                continue
    return dict(FALLBACK_CONFIG)


def save_config(cfg):
    with open(USER_CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        # 窗口标题只放简名：完整品牌名+副标已经在正文首行显示一次，
        # 标题栏再写一遍会让人觉得“这句话重复了”。
        self.title(f"ExcelRouter v{APP_VERSION}")
        self._set_window_icon()
        self.geometry("820x800")
        self.minsize(720, 640)
        self._stop_flag = False
        self._running = False
        self.cfg = load_config()
        self._mode = self.cfg.get("ui_mode", "excel")   # "excel" / "pdf"
        self._pdf_paths = [p for p in self.cfg.get("pdf_input_paths", []) if os.path.exists(p)]
        self._scanned_map_path = None   # 最近扫描过的映射清单路径（防过期结果）
        self._columns = []
        self._adv_open = False
        self._log_open = False
        self._scanned_path = None     # 最近一次扫描的输入路径（防重复扫描/过期结果）
        self._out_auto = ""           # 最近一次自动推荐的输出目录（用户没改过才允许覆盖）
        self._last_output = None      # 最近一次成功运行的输出目录（「打开输出文件夹」用）
        self._ui_q = queue.Queue()            # 子线程 → 主线程的消息队列
        self._prog_indeterminate = False      # 进度条当前是否处于不定态动画
        self._build_ui()
        self._update_input_ui()
        self._apply_mode()
        p = self.cfg.get("input_path", "")
        saved_out = self._output_var.get().strip()
        if not saved_out:
            self._suggest_output()
        elif saved_out and os.path.normcase(saved_out) == os.path.normcase(self._auto_output_for(p) or ""):
            self._out_auto = saved_out    # 上次存的就是自动值：换输入时允许跟着更新
        if p and os.path.exists(p):
            self._scan_input()        # 上次的输入还在：启动即自动扫描，打开就能直接开始
        mp = self.cfg.get("pdf_mapping_path", "")
        if mp and os.path.exists(mp):
            self._scan_mapping()      # 上次的映射清单还在：同样启动即恢复列下拉
        self.after(100, self._pump_ui)

    def _set_window_icon(self):
        """设置标题栏 / 任务栏 / 最小化时的窗口图标。

        `--icon app.ico`（build.bat / CI）只改 exe 文件本身的图标，运行起来的
        Tk 窗口图标要靠 iconbitmap() 单独设置，否则标题栏/任务栏还是 Tk 默认
        羽毛图标——这就是“exe 图标换了，但程序里和最小化时的图标没跟着换”的
        原因。app.ico 需要被 --add-data 打包进去才能在冻结后找到（见 build.bat）。
        """
        for base in (_resource_dir(), _app_dir()):
            ico = os.path.join(base, "app.ico")
            if os.path.exists(ico):
                try:
                    self.iconbitmap(ico)
                    return
                except Exception:
                    continue

    # ── UI 构建 ──────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_header()                                   # row 0 品牌区（含模式切换）
        self._body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self._body.grid(row=1, column=0, padx=12, pady=(4, 0), sticky="nsew")
        self._body.grid_columnconfigure(0, weight=1)
        # 两个模式各占一个 Frame（同一格），切换时 grid()/grid_remove()
        self._excel_frame = ctk.CTkFrame(self._body, fg_color="transparent")
        self._excel_frame.grid(row=0, column=0, sticky="ew")
        self._excel_frame.grid_columnconfigure(0, weight=1)
        self._pdf_frame = ctk.CTkFrame(self._body, fg_color="transparent")
        self._pdf_frame.grid(row=0, column=0, sticky="ew")
        self._pdf_frame.grid_columnconfigure(0, weight=1)
        self._build_step_input(self._excel_frame)              # ① 选表格
        self._build_step_field(self._excel_frame)              # ② 选字段（含批量选项）
        self._build_adv_area(self._excel_frame)                # ▸ 高级设置
        self._build_pdf_step_input(self._pdf_frame)            # PDF ① 选 PDF
        self._build_pdf_step_map(self._pdf_frame)              # PDF ② 选映射清单
        self._build_action(self)                               # row 2 ③ 开始（固定底部）
        self._build_bottom(self)                               # row 3-5 工具条 / 日志 / 页脚

    def _build_header(self):
        # 品牌文案已定稿（副标/标语/特性行），只调整排版：作者信息移到页脚
        head = ctk.CTkFrame(self, fg_color="transparent")
        head.grid(row=0, column=0, padx=20, pady=(14, 4), sticky="ew")
        brand = ctk.CTkFrame(head, fg_color="transparent")
        brand.pack(anchor="w")
        ctk.CTkLabel(brand, text="ExcelRouter",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(side="left")
        ctk.CTkLabel(brand, text="Excel 智能拆分工具",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=("gray25", "gray75")).pack(side="left", padx=(10, 0), pady=(7, 0))
        ctk.CTkLabel(head, text="整个文件夹一键拆完：按部门、区域、工号等字段自动拆分，打包分发",
                     font=ctk.CTkFont(size=12),
                     text_color=("gray30", "gray70")).pack(anchor="w", pady=(4, 0))
        ctk.CTkLabel(head, text="保留原格式 · 跨文件自动合并 · 单个文件也能拆 · 可同时拆到每个人",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(anchor="w", pady=(1, 0))
        # 功能模式切换：Excel 拆分（默认） / PDF 按网格加密分发
        self._mode_seg = ctk.CTkSegmentedButton(head, values=["Excel 拆分", "PDF 加密分发"],
                                                command=self._on_mode_change)
        self._mode_seg.set("PDF 加密分发" if self._mode == "pdf" else "Excel 拆分")
        self._mode_seg.pack(anchor="w", pady=(8, 0))

    def _step_card(self, parent, row, num, title, padx=8, pady=(0, 10)):
        """带编号徽章的步骤卡片：编号即真实操作顺序，是界面的导航主线。

        返回 (card, 标题 Label)：标题引用给需要随模式改文案的卡片用（如③操作卡）。
        """
        card = ctk.CTkFrame(parent)
        card.grid(row=row, column=0, padx=padx, pady=pady, sticky="ew")
        card.grid_columnconfigure(0, weight=1)
        bar = ctk.CTkFrame(card, fg_color="transparent")
        bar.grid(row=0, column=0, padx=12, pady=(10, 2), sticky="w")
        ctk.CTkLabel(bar, text=str(num), width=26, height=26, corner_radius=13,
                     fg_color=ACCENT, text_color="white",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        title_lbl = ctk.CTkLabel(bar, text=title,
                                 font=ctk.CTkFont(size=14, weight="bold"))
        title_lbl.pack(side="left", padx=(8, 0))
        return card, title_lbl

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

        self._in_status = ctk.CTkLabel(card, text="还没有选择文件（也可以把路径粘贴到上面的输入框）",
                                       font=ctk.CTkFont(size=11), text_color=C_MUTED,
                                       anchor="w", justify="left")
        self._in_status.grid(row=3, column=0, padx=12, pady=(0, 10), sticky="w")

    def _build_step_field(self, parent):
        card, _ = self._step_card(parent, 1, 2, "按哪个字段拆分")

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.grid(row=1, column=0, padx=12, pady=(6, 2), sticky="w")
        ctk.CTkLabel(row, text="拆分字段").pack(side="left")
        self._split_var = ctk.StringVar(value=self.cfg.get("split_column") or COL_PLACEHOLDER)
        self._split_menu = ctk.CTkOptionMenu(row, variable=self._split_var,
                                             values=[self._split_var.get()], width=220)
        self._split_menu.pack(side="left", padx=(8, 8))
        _ghost_button(row, text="🔄 扫描字段", width=104,
                      command=self._scan_input).pack(side="left")

        ctk.CTkLabel(card, text="拆分字段的每个取值各生成一个文件：比如按「部门」拆 → 销售部.xlsx、财务部.xlsx…",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED).grid(
            row=2, column=0, padx=12, pady=(2, 0), sticky="w")
        self._status = ctk.CTkLabel(card, text="", font=ctk.CTkFont(size=11),
                                    text_color=C_MUTED, anchor="w")
        self._status.grid(row=3, column=0, padx=12, pady=(0, 6), sticky="w")

        self._batch_frame = self._build_batch(card)
        self._batch_frame.grid(row=4, column=0, padx=12, pady=(2, 12), sticky="ew")
        self._batch_frame.grid_remove()   # 仅输入为文件夹时显示（_update_input_ui 控制）

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
        ctk.CTkCheckBox(opts, text="精确匹配", variable=self._exact_var).grid(row=0, column=0, padx=8, pady=4)
        ctk.CTkCheckBox(opts, text="跨文件合并汇总", variable=self._merge_var).grid(row=0, column=1, padx=8, pady=4)
        ctk.CTkCheckBox(opts, text="保留格式", variable=self._preserve_var).grid(row=0, column=2, padx=8, pady=4)
        ctk.CTkCheckBox(opts, text="完成后打开输出", variable=self._auto_open_var).grid(row=0, column=3, padx=8, pady=4)
        ctk.CTkCheckBox(opts, text="保留公式（收件人可见计算过程，仅同行公式，需先勾选保留格式）",
                        variable=self._keep_formula_var).grid(
            row=1, column=0, columnspan=4, padx=8, pady=(0, 4), sticky="w")

        self._on_header_mode(self._header_seg.get())
        return f

    # ── PDF 加密分发模式的步骤卡 ─────────────────────────
    def _build_pdf_step_input(self, parent):
        card, _ = self._step_card(parent, 0, 1, "选择要分发的 PDF")

        btns = ctk.CTkFrame(card, fg_color="transparent")
        btns.grid(row=1, column=0, padx=12, pady=(6, 2), sticky="w")
        ctk.CTkButton(btns, text="📄 选 PDF 文件（可多选）", width=190,
                      command=self._pick_pdfs).pack(side="left")
        _ghost_button(btns, text="清空", width=60,
                      command=self._clear_pdfs).pack(side="left", padx=(10, 0))

        self._pdf_in_status = ctk.CTkLabel(card, text="", font=ctk.CTkFont(size=11),
                                           text_color=C_MUTED, anchor="w", justify="left")
        self._pdf_in_status.grid(row=2, column=0, padx=12, pady=(0, 10), sticky="w")
        self._update_pdf_in_status()

    def _build_pdf_step_map(self, parent):
        card, _ = self._step_card(parent, 1, 2, "选择密码映射清单（Excel）")

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

        ctk.CTkLabel(card, text="清单里每行一个网格：网格名、专属密码，可选接收人（方便照着群发）",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED).grid(
            row=2, column=0, padx=12, pady=(2, 0), sticky="w")

        crow = ctk.CTkFrame(card, fg_color="transparent")
        crow.grid(row=3, column=0, padx=12, pady=(6, 2), sticky="w")
        ctk.CTkLabel(crow, text="网格列").pack(side="left")
        self._pdfgrid_var = ctk.StringVar(value=self.cfg.get("pdf_grid_column") or PDF_COL_PLACEHOLDER)
        self._pdfgrid_menu = ctk.CTkOptionMenu(crow, variable=self._pdfgrid_var,
                                               values=[self._pdfgrid_var.get()], width=140)
        self._pdfgrid_menu.pack(side="left", padx=(6, 12))
        ctk.CTkLabel(crow, text="密码列").pack(side="left")
        self._pdfpwd_var = ctk.StringVar(value=self.cfg.get("pdf_password_column") or PDF_COL_PLACEHOLDER)
        self._pdfpwd_menu = ctk.CTkOptionMenu(crow, variable=self._pdfpwd_var,
                                              values=[self._pdfpwd_var.get()], width=140)
        self._pdfpwd_menu.pack(side="left", padx=(6, 12))
        ctk.CTkLabel(crow, text="接收人列").pack(side="left")
        self._pdfrecv_var = ctk.StringVar(value=self.cfg.get("pdf_receiver_column") or PDF_RECV_NONE)
        self._pdfrecv_menu = ctk.CTkOptionMenu(crow, variable=self._pdfrecv_var,
                                               values=[self._pdfrecv_var.get()], width=140)
        self._pdfrecv_menu.pack(side="left", padx=(6, 0))

        self._pdf_map_status = ctk.CTkLabel(card, text="", font=ctk.CTkFont(size=11),
                                            text_color=C_MUTED, anchor="w")
        self._pdf_map_status.grid(row=4, column=0, padx=12, pady=(0, 6), sticky="w")

        wm = ctk.CTkFrame(card)
        wm.grid(row=5, column=0, padx=12, pady=(2, 12), sticky="ew")
        self._pdfwm_var = ctk.BooleanVar(value=self.cfg.get("pdf_watermark", True))
        ctk.CTkCheckBox(wm, text="加网格专属水印（泄露可溯源）",
                        variable=self._pdfwm_var,
                        command=self._update_wm_state).grid(row=0, column=0, padx=12, pady=(8, 4), sticky="w")
        wrow = ctk.CTkFrame(wm, fg_color="transparent")
        wrow.grid(row=1, column=0, padx=(38, 12), pady=(0, 8), sticky="w")
        ctk.CTkLabel(wrow, text="水印文字").pack(side="left")
        self._pdfwmtext_var = ctk.StringVar(value=self.cfg.get("pdf_watermark_text", "{grid} {date}"))
        self._pdfwmtext_entry = ctk.CTkEntry(wrow, textvariable=self._pdfwmtext_var, width=220)
        self._pdfwmtext_entry.pack(side="left", padx=(6, 8))
        ctk.CTkLabel(wrow, text="{grid}＝网格名  {date}＝日期",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED).pack(side="left")
        self._update_wm_state()

    def _update_wm_state(self):
        self._pdfwmtext_entry.configure(state="normal" if self._pdfwm_var.get() else "disabled")

    def _update_pdf_in_status(self):
        n = len(self._pdf_paths)
        if not n:
            self._pdf_in_status.configure(text="还没有选择 PDF；每个网格都会拿到全部所选文件（各自密码+水印）",
                                          text_color=C_MUTED)
            return
        names = [os.path.basename(p) for p in self._pdf_paths]
        shown = "、".join(names[:3]) + ("…" if n > 3 else "")
        self._pdf_in_status.configure(text=f"✓ 已选 {n} 个：{shown}", text_color=C_OK)

    def _pick_pdfs(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF 文件", "*.pdf")])
        if paths:
            self._pdf_paths = [os.path.normpath(p) for p in paths]
            self._update_pdf_in_status()
            self._suggest_pdf_output()

    def _clear_pdfs(self):
        self._pdf_paths = []
        self._update_pdf_in_status()

    def _suggest_pdf_output(self):
        """PDF 模式的默认输出：第一个 PDF 所在目录下「分发结果」。规则同 _suggest_output。"""
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
            try:
                from core.pdf_dist import list_mapping_columns
                cols = list_mapping_columns(p)
            except Exception:
                cols = []
            self._ui_q.put(("pdf_scan", (cols, p)))

        threading.Thread(target=work, daemon=True).start()

    def _on_pdf_scan(self, payload):
        cols, p = payload
        if p != self._pdf_map_var.get().strip():
            return    # 结果已过期
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
        recv_values = [PDF_RECV_NONE] + cols
        self._pdfrecv_menu.configure(values=recv_values)
        # 「（不需要）」是未选状态的占位，也参与推荐；用户真不需要可再手动选回去
        if self._pdfrecv_var.get() not in cols:
            self._pdfrecv_var.set(self._recommend_col(cols, ("接收", "姓名", "负责人", "联系")) or PDF_RECV_NONE)
        self._pdf_map_status.configure(text=f"✓ 识别到 {len(cols)} 列，确认三个下拉选的对不对", text_color=C_OK)

    @staticmethod
    def _recommend_col(cols, keywords):
        for kw in keywords:
            for c in cols:
                if kw.lower() in str(c).lower():
                    return c
        return None

    # ── 模式切换 ─────────────────────────────────────────
    def _on_mode_change(self, label):
        self._mode = "pdf" if label == "PDF 加密分发" else "excel"
        self._apply_mode()

    def _apply_mode(self):
        """按当前模式显示对应步骤区并更新③卡片文案；运行中禁止切换按钮已够复杂，
        简单起见运行中也允许切界面，但开始按钮在运行中本来就是禁用态。"""
        if self._mode == "pdf":
            self._excel_frame.grid_remove()
            self._pdf_frame.grid()
            self._action_title.configure(text="开始分发")
            if not self._running:
                self._start_btn.configure(text="🔐 开始分发")
                self._run_status.configure(
                    text="完成 ① ② 后点「开始分发」；每个网格生成专属密码+水印的副本，并输出分发清单",
                    text_color=C_MUTED)
        else:
            self._pdf_frame.grid_remove()
            self._excel_frame.grid()
            self._action_title.configure(text="开始拆分")
            if not self._running:
                self._start_btn.configure(text="▶ 开始拆分")
                self._run_status.configure(
                    text="完成 ① ② 后直接点「开始拆分」；保存位置不用改，结果会放进自动创建的「拆分结果」文件夹",
                    text_color=C_MUTED)

    def _build_action(self, parent):
        card, self._action_title = self._step_card(parent, 2, 3, "开始拆分", padx=20, pady=(6, 0))

        orow = ctk.CTkFrame(card, fg_color="transparent")
        orow.grid(row=1, column=0, padx=12, pady=(4, 2), sticky="ew")
        orow.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(orow, text="结果保存到").grid(row=0, column=0, sticky="w")
        self._output_var = ctk.StringVar(value=self.cfg.get("output_path", ""))
        ctk.CTkEntry(orow, textvariable=self._output_var).grid(row=0, column=1, padx=8, sticky="ew")
        _ghost_button(orow, text="浏览", width=60,
                      command=self._browse_output).grid(row=0, column=2)

        btns = ctk.CTkFrame(card, fg_color="transparent")
        btns.grid(row=2, column=0, padx=12, pady=(8, 2), sticky="w")
        self._start_btn = ctk.CTkButton(btns, text="▶ 开始拆分", width=200, height=38,
                                        font=ctk.CTkFont(size=15, weight="bold"),
                                        command=self._start)
        self._start_btn.grid(row=0, column=0)
        self._stop_btn = ctk.CTkButton(btns, text="⏹ 停止", width=90,
                                       fg_color="gray40", hover_color="gray30",
                                       command=self._stop)
        self._stop_btn.grid(row=0, column=1, padx=(10, 0))
        self._stop_btn.grid_remove()          # 只在运行时出现
        self._open_btn = _ghost_button(btns, text="📂 打开输出文件夹", width=140,
                                       command=self._open_output)
        self._open_btn.grid(row=0, column=2, padx=(10, 0))
        self._open_btn.grid_remove()          # 成功后出现

        self._progress = ctk.CTkProgressBar(card)
        self._progress.grid(row=3, column=0, padx=12, pady=(8, 2), sticky="ew")
        self._progress.set(0)
        self._run_status = ctk.CTkLabel(card, text="完成 ① ② 后直接点「开始拆分」；保存位置不用改，结果会放进自动创建的「拆分结果」文件夹",
                                        font=ctk.CTkFont(size=11), text_color=C_MUTED, anchor="w")
        self._run_status.grid(row=4, column=0, padx=12, pady=(0, 10), sticky="w")

    def _build_bottom(self, parent):
        bar = ctk.CTkFrame(parent, fg_color="transparent")
        bar.grid(row=3, column=0, padx=20, pady=(4, 0), sticky="ew")
        bar.grid_columnconfigure(1, weight=1)
        self._log_btn = _flat_button(bar, text="▸ 处理详情", width=100, anchor="w",
                                     command=self._toggle_log)
        self._log_btn.grid(row=0, column=0, sticky="w")
        right = ctk.CTkFrame(bar, fg_color="transparent")
        right.grid(row=0, column=2, sticky="e")
        _flat_button(right, text="💬 反馈建议", width=90,
                     command=self._open_feedback).pack(side="left")
        _flat_button(right, text="保存配置", width=80,
                     command=self._save_cfg).pack(side="left", padx=(6, 0))

        self._log_box = ctk.CTkTextbox(parent, state="disabled", wrap="none", height=190)
        # 展开时 grid 到 row=4（_toggle_log 控制），默认收起

        # 「数据不出本地」是核心卖点，字号/颜色比作者署名这类次要信息更突出
        footer = ctk.CTkFrame(parent, fg_color="transparent")
        footer.grid(row=5, column=0, padx=22, pady=(2, 8), sticky="w")
        ctk.CTkLabel(footer, text="🔒 数据仅在本机处理，不上传任何服务器",
                     font=ctk.CTkFont(size=12, weight="bold"), text_color=C_OK).pack(anchor="w")
        ctk.CTkLabel(footer, text="作者 AbeLin · MIT 开源",
                     font=ctk.CTkFont(size=10), text_color="gray").pack(anchor="w", pady=(1, 0))

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

    def _toggle_log(self, show=None):
        show = (not self._log_open) if show is None else show
        if show == self._log_open:
            return
        self._log_open = show
        if show:
            self._log_box.grid(row=4, column=0, padx=20, pady=(2, 0), sticky="ew")
            self._log_btn.configure(text="▾ 处理详情")
        else:
            self._log_box.grid_remove()
            self._log_btn.configure(text="▸ 处理详情")

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

    def _browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self._output_var.set(os.path.normpath(path))

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
            self._set_in_status("还没有选择文件（也可以把路径粘贴到上面的输入框）" if not p
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
            tpl, n = None, 0
            if is_dir:
                for root, dirs, files in os.walk(p):
                    if out_abs:
                        ra = os.path.normcase(os.path.abspath(root))
                        if ra == out_abs or ra.startswith(out_abs + os.sep):
                            dirs[:] = []
                            continue
                    for fn in sorted(files):
                        if fn.lower().endswith(('.xlsx', '.xls')) and not fn.startswith('~$'):
                            n += 1
                            if tpl is None:
                                tpl = os.path.join(root, fn)
            else:
                tpl, n = p, 1
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
            self._ui_q.put(("scan", (cols, n, os.path.basename(tpl) if tpl else "", is_dir, p)))

        threading.Thread(target=work, daemon=True).start()

    def _on_scan(self, payload):
        cols, n, tpl_name, is_dir, p = payload
        if p != self._input_var.get().strip():
            return    # 结果已过期（用户又换了输入）
        name = os.path.basename(os.path.normpath(p)) or p
        if is_dir and n == 0:
            self._set_in_status(f"⚠ 文件夹「{name}」里没有找到 Excel 文件（.xlsx / .xls）", C_WARN)
            self._set_scan_status("", C_MUTED)
            return
        if is_dir:
            self._set_in_status(f"✓ 文件夹「{name}」：找到 {n} 个 Excel，全部一起拆", C_OK)
        else:
            self._set_in_status(f"✓ 已选择：{name}", C_OK)
        self._on_columns(cols, tpl_name)

    def _on_columns(self, cols, tpl_name=None):
        self._columns = cols
        if cols:
            self._split_menu.configure(values=cols)
            if self._split_var.get() not in cols:
                self._split_var.set(self._recommend_split(cols))
            self._pcol_menu.configure(values=cols)
            if self._pcol_var.get() not in cols:
                self._pcol_var.set(self._recommend_person(cols))
            src = f"（来自「{tpl_name}」）" if tpl_name else ""
            self._set_scan_status(f"✓ 识别到 {len(cols)} 个字段{src}，确认下拉框选的对不对", C_OK)
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
        self._set_scan_status("未能识别字段：点「🔄 扫描字段」重试，或到「高级设置」手动指定表头行", C_WARN)

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
                if d == "拆分结果" or d.endswith("_拆分结果") or re.fullmatch(r"\d{8}结果", d):
                    hits.append(os.path.join(root, d))
                    dirs.remove(d)    # 命中即整树算一个，不再往里翻
        return hits

    def _set_in_status(self, text, color=C_MUTED):
        self._in_status.configure(text=text, text_color=color)

    def _set_scan_status(self, text, color=C_MUTED):
        self._status.configure(text=text, text_color=color)

    # ── 配置收集 ───────────────────────────────────────
    def _split_list(self, text):
        return [k.strip() for k in text.split(",") if k.strip()]

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

    def _collect_config(self):
        mode_map = {"自动": "auto", "指定行号": "row", "关键词": "keyword"}
        try:
            header_row = int(self._hrow_var.get())
        except (TypeError, ValueError):
            header_row = 1

        split_col = "" if self._split_var.get() == COL_PLACEHOLDER else self._split_var.get()
        pcol = "" if self._pcol_var.get() == COL_PLACEHOLDER else self._pcol_var.get()
        alias, _ = self._parse_alias()

        recv = self._pdfrecv_var.get()
        return {
            "input_path":  self._input_var.get().strip(),
            "output_path": self._output_var.get().strip(),
            "ui_mode":     self._mode,
            "pdf_input_paths": list(self._pdf_paths),
            "pdf_mapping_path": self._pdf_map_var.get().strip(),
            "pdf_grid_column": "" if self._pdfgrid_var.get() == PDF_COL_PLACEHOLDER else self._pdfgrid_var.get(),
            "pdf_password_column": "" if self._pdfpwd_var.get() == PDF_COL_PLACEHOLDER else self._pdfpwd_var.get(),
            "pdf_receiver_column": "" if recv in (PDF_RECV_NONE, PDF_COL_PLACEHOLDER) else recv,
            "pdf_watermark": self._pdfwm_var.get(),
            "pdf_watermark_text": self._pdfwmtext_var.get().strip() or "{grid} {date}",
            "pdf_watermark_opacity": self.cfg.get("pdf_watermark_opacity", 0.15),
            "pdf_watermark_angle": self.cfg.get("pdf_watermark_angle", 45),
            "header_mode": mode_map.get(self._header_seg.get(), "auto"),
            "header_row":  header_row,
            "grid_keys":   self._split_list(self._grid_keys_var.get()),
            "id_keys":     self._split_list(self._id_keys_var.get()),
            "split_column": split_col,
            "selected_values": self._split_list(self._selvals_var.get()),
            "person_column": pcol,
            "to_person":   self._person_var.get(),
            "person_file_filter": self._split_list(self._pfilter_var.get()),
            "value_alias_map": alias,
            "skip_values": self._split_list(self._skip_var.get()),
            "merge_across_files": self._merge_var.get(),
            "make_zip":    self._zip_var.get(),
            "exact_match": self._exact_var.get(),
            "preserve_format": self._preserve_var.get(),
            "keep_formulas": self._keep_formula_var.get(),
            "auto_open_output": self._auto_open_var.get(),
        }

    def _save_cfg(self):
        try:
            save_config(self._collect_config())
            messagebox.showinfo("已保存", f"配置已保存到：\n{USER_CONFIG_PATH}\n下次启动自动加载。\n（每次成功拆分后也会自动记住当前配置）")
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存配置：{e}")

    # ── 运行 ───────────────────────────────────────────
    def _log(self, msg):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _stop_indeterminate(self):
        """进度条从不定态动画切回真实进度模式。"""
        if self._prog_indeterminate:
            self._prog_indeterminate = False
            self._progress.stop()
            self._progress.configure(mode="determinate")

    def _pump_ui(self):
        """主线程 UI 泵：每 100ms 批量消费子线程消息（日志/进度/扫描/完成）。

        子线程绝不直接碰 Tk 控件；日志批量合并插入，进度只取最新值——
        避免高频 after(0) 回调打满主线程事件队列（大文件时界面假死的诱因之一）。
        """
        logs, progress = [], None
        done: Any = _MISSING
        scan: Any = _MISSING
        pdf_scan: Any = _MISSING
        try:
            while True:
                kind, payload = self._ui_q.get_nowait()
                if kind == "log":
                    logs.append(payload)
                elif kind == "progress":
                    progress = payload
                elif kind == "done":
                    done = payload
                elif kind == "scan":
                    scan = payload
                elif kind == "pdf_scan":
                    pdf_scan = payload
        except queue.Empty:
            pass
        if logs:
            joined = "\n".join(logs)
            self._log(joined)
            if self._running:
                # 状态行镜像最新一条日志：不展开「处理详情」也能看到进展
                last = next((ln.strip() for ln in reversed(joined.splitlines()) if ln.strip()), "")
                if last:
                    self._run_status.configure(text=("⏳ " + last)[:70], text_color=C_MUTED)
        if progress is not None:
            self._stop_indeterminate()
            self._progress.set(progress)
        if scan is not _MISSING:
            self._on_scan(scan)
        if pdf_scan is not _MISSING:
            self._on_pdf_scan(pdf_scan)
        if done is not _MISSING:
            self._on_done(*done)
        self.after(100, self._pump_ui)

    def _enter_running(self, busy_text, first_log):
        """进入运行态的公共 UI 准备：两种模式共用（按钮/进度/日志/状态行）。"""
        self._stop_flag = False
        self._running = True
        self._last_output = None
        self._start_btn.configure(state="disabled", text=busy_text)
        self._stop_btn.grid()
        self._stop_btn.configure(state="normal")
        self._open_btn.grid_remove()
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")
        # 点击的瞬间就要看到反馈：先用不定态动画，第一条真实进度回来后自动切换（见 _pump_ui）
        self._prog_indeterminate = True
        self._progress.configure(mode="indeterminate")
        self._progress.start()
        self._run_status.configure(text=first_log, text_color=C_MUTED)
        self._log("▶ " + first_log)

    def _start_pdf(self):
        cfg = self._collect_config()
        if not cfg["pdf_input_paths"]:
            messagebox.showwarning("先选择 PDF", "请先在第 ① 步选择要分发的 PDF 文件。")
            return
        missing = [p for p in cfg["pdf_input_paths"] if not os.path.exists(p)]
        if missing:
            messagebox.showwarning("PDF 不存在", "以下文件找不到了，请重新选择：\n" + "\n".join(missing[:3]))
            return
        if not cfg["pdf_mapping_path"] or not os.path.exists(cfg["pdf_mapping_path"]):
            messagebox.showwarning("先选择映射清单", "请在第 ② 步选择「网格 → 密码」的 Excel 映射清单。")
            return
        if not cfg["pdf_grid_column"] or not cfg["pdf_password_column"]:
            messagebox.showwarning("先选择列", "请在第 ② 步选好「网格列」和「密码列」。\n如果下拉框没有内容，重新选一次映射清单。")
            return
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

    def _start(self):
        if self._mode == "pdf":
            return self._start_pdf()
        cfg = self._collect_config()
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

    def _stop(self):
        self._stop_flag = True
        self._log("正在停止，请等待当前文件处理完成...")

    def _open_feedback(self):
        """打开匿名反馈页；顺手把版本号复制进剪贴板，方便用户粘贴到问卷。"""
        try:
            self.clipboard_clear()
            self.clipboard_append(f"ExcelRouter v{APP_VERSION}")
        except Exception:
            pass
        webbrowser.open(FEEDBACK_URL)
        self._log("💬 已打开反馈页，版本号已复制到剪贴板，粘贴到问卷即可。")

    def _open_output(self):
        path = self._last_output or self._output_var.get().strip()
        if path and os.path.exists(path):
            try:
                subprocess.Popen(f'explorer "{os.path.normpath(path)}"')
            except Exception:
                pass

    def _on_done(self, cfg, output_path):
        self._running = False
        is_pdf = cfg.get("ui_mode") == "pdf"
        self._start_btn.configure(state="normal",
                                  text="🔐 开始分发" if self._mode == "pdf" else "▶ 开始拆分")
        self._stop_btn.grid_remove()
        self._stop_indeterminate()
        self._progress.set(1 if output_path else 0)
        if output_path:
            self._last_output = output_path
            self._open_btn.grid()
            if self._stop_flag:
                self._run_status.configure(text="⏹ 已停止：处理完的部分已保存，可点「打开输出文件夹」查看",
                                           text_color=C_WARN)
            else:
                self._run_status.configure(text="✅ 分发完成！结果已保存" if is_pdf else "✅ 拆分完成！结果已保存",
                                           text_color=C_OK)
                self._log("💬 用得顺手或踩了坑？点「反馈建议」匿名告诉作者（1 分钟）。")
            try:
                save_config(cfg)      # 静默记住本次配置：下次打开即用（失败运行不存，避免存坏参数）
            except Exception:
                pass
        else:
            self._run_status.configure(text="❌ 没有完成：下方「处理详情」里有原因", text_color=C_ERR)
            self._toggle_log(show=True)
        if cfg.get("auto_open_output") and output_path and os.path.exists(output_path):
            try:
                subprocess.Popen(f'explorer "{os.path.normpath(output_path)}"')
            except Exception:
                pass


def run():
    app = App()
    app.mainloop()
