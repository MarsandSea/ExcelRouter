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
import json
import queue
import webbrowser
from typing import Any

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox

from gui.common import (
    _resource_dir, _app_dir, USER_CONFIG_PATH, COL_PLACEHOLDER, SPEED_STD, SPEED_FAST,
    PDF_COL_PLACEHOLDER, MODE_LABELS, DS_SCENES, _MISSING, C_OK, C_WARN, C_ERR, C_MUTED,
    PRIMARY, PRIMARY_HOVER, OK_BANNER_BG, CARD_BORDER, _init_fonts, _ghost_button, _flat_button,
    load_config, save_config,
)
from gui.mode_collect import CollectModeMixin
from gui.mode_desens import DesensModeMixin
from gui.mode_pdf import PdfModeMixin
from gui.mode_split import SplitModeMixin
from gui.platform_compat import open_folder, parse_drop_paths, set_window_icon

# 拖拽支持（v2.7）：tkinterdnd2 为可选依赖——打包环境缺它时静默退回无拖拽，
# 不影响其余任何功能。CTk + TkinterDnD 的混编写法是社区通行方案（本地 spike 验证过）。
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _DND_OK = True
except Exception:
    TkinterDnD = None  # type: ignore[assignment]
    DND_FILES = None   # type: ignore[assignment]
    _DND_OK = False

if _DND_OK:
    class _RootBase(ctk.CTk, TkinterDnD.DnDWrapper):  # type: ignore[misc]
        """customtkinter 根窗口 + tkdnd 拖放能力的混编基类。"""
else:
    _RootBase = ctk.CTk  # type: ignore[assignment]  # 缺依赖时退回纯 CTk


APP_VERSION = "3.2.2"
# 匿名反馈问卷地址（问卷 URL 确定后替换此处即可，一行改动 + 打 tag 发版）
FEEDBACK_URL = "https://f.wps.cn/g/pBOAWUQc/"
# 在线 FAQ（「❓ 常见问题」按钮）：链接 Gitee 镜像而非 GitHub——国内办公网络访问
# GitHub 不稳定，用户点帮助按钮却打不开比没有按钮更糟。内容与 GitHub 同步发版。
FAQ_URL = "https://gitee.com/Marsandsea/Excelrouter/blob/main/docs/FAQ.md"

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(SplitModeMixin, CollectModeMixin, DesensModeMixin, PdfModeMixin, _RootBase):
    def __init__(self):
        super().__init__()
        # 字体统一必须在任何组件创建前、root 建立后执行：
        # tkfont.families() 需要 root；而组件的 CTkFont 在构造时读主题值，晚了改无效。
        _init_fonts()
        # 启动期（_build_ui 之前）产生的日志先攒着，_log_box 建好后统一排空
        self._pending_logs = []
        # _DND_OK = tkinterdnd2 这个 Python 包 import 成功；
        # self._dnd_ok = 它的原生 tkdnd 库真的加载成功。两者必须分开：
        # 麒麟等国产化 Linux 上 .so 的 glibc/libX11 ABI 常对不上，_require() 会抛
        # RuntimeError，没有 try 包住时会直接把 __init__ 打穿、进程退出。
        # 拖拽只是锦上添花，绝不能让它拦住整个程序启动。
        self._dnd_ok = False
        if _DND_OK:
            try:
                # 加载 tkdnd 二进制（守卫条件已保证 TkinterDnD 非 None）
                self.TkdndVersion = TkinterDnD._require(self)  # type: ignore[union-attr]
                self._dnd_ok = True
            except Exception as e:
                self.TkdndVersion = None
                self._pending_logs.append(
                    f"ℹ 拖拽功能不可用（{e}），不影响使用，请点「选文件 / 选文件夹」按钮。")
        # 窗口标题只放简名：完整品牌名+副标已经在正文首行显示一次，
        # 标题栏再写一遍会让人觉得“这句话重复了”。
        self.title(f"ExcelRouter v{APP_VERSION}")
        self._set_window_icon()
        self._stop_flag = False
        self._running = False
        self.cfg = load_config()
        self._init_geometry()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._first_run = not os.path.exists(USER_CONFIG_PATH) and not self.cfg.get("input_path")
        # 界面缩放（大字号）：启动时应用；运行时切换只存配置、下次启动生效
        scale = self.cfg.get("ui_scale", 100)
        if isinstance(scale, int) and scale != 100:
            ctk.set_widget_scaling(scale / 100)
        # 深浅色（v2.7.1）：支持手动固定，不再只能跟随系统
        mode = self.cfg.get("appearance_mode", "system")
        if mode in ("light", "dark"):
            ctk.set_appearance_mode(mode)
        self._mode = self.cfg.get("ui_mode", "excel")   # "excel" / "pdf" / "collect"
        if self._mode not in MODE_LABELS:
            self._mode = "excel"
        self._col_scanned = None      # Excel 汇总：最近清点过的输入（防重复 / 过期结果）
        self._col_info: dict[str, Any] = {}   # 最近一次清点的结果（① 状态行、② 名单提示用）
        self._col_notice = ""         # 结果卡片里的文字（查收通报 / 汇总说明，「复制」用）
        self._col_progress = ""       # 最近一次查收的进度表路径
        self._col_merged = ""         # 最近一次汇总表路径（「打开汇总表」用）
        self._ds_scanned = None       # 数据脱敏：最近识别过的输入（防重复 / 过期结果）
        self._ds_info: dict[str, Any] = {}   # 最近一次识别的结果（只有列名、类型和计数）
        self._ds_overrides: dict[str, str] = {}   # 逐列调整：{列名: 处理方式}，和用途默认一样的不记
        self._ds_report = ""          # 最近一次脱敏报告路径
        self._split_ds_cols = []      # 拆分：样本表里认出的个人信息列 [(列名, 类型)]
        self._rpt_hints = {"seq_column": "", "total_rows": 0}   # 拆分：样本表里的序号列 / 合计行（v3.2 通报表）
        self._pdf_paths = [p for p in self.cfg.get("pdf_input_paths", []) if os.path.exists(p)]
        self._scanned_map_path = None   # 最近扫描过的映射清单路径（防过期结果）
        self._columns = []
        self._adv_open = False
        self._log_open = False
        self._scanned_path = None     # 最近一次扫描的输入路径（防重复扫描/过期结果）
        self._out_auto = ""           # 最近一次自动推荐的输出目录（用户没改过才允许覆盖）
        self._last_output = None      # 最近一次成功运行的输出目录（「打开输出文件夹」用）
        self._tpl_path = None         # 最近一次扫描识别出的模板文件全路径（取值枚举用）
        self._values_all = []         # 当前拆分字段枚举到的全部取值
        self._value_vars = {}         # 取值 -> BooleanVar（勾选结果写回 selected_values）
        self._values_open = False     # 取值勾选面板展开状态
        self._values_panel = None     # 展开中的勾选面板
        self._enum_token = 0          # 取值枚举序号：防止过期结果回填
        self._enum_after = None       # 防抖 after id
        # 取值枚举的「作用域」（v2.9）：_tpl_path 只是批量里的一个样本文件，
        # 它枚举出的取值天然可能不全，而拆分是对全部文件生效的。这几个状态
        # 就是为了让界面能说清「这 N 组是哪来的」，不再把抽样结果当全量结论。
        self._scan_n = 0              # 最近一次扫描到的 Excel 文件数
        self._scan_bytes = 0          # 同上，总字节（换样本文件时沿用）
        self._scan_is_dir = False     # 同上，输入是否为文件夹
        self._tpl_is_sample = False   # _tpl_path 只是多个文件里的一个
        self._values_scope = "sample"  # 取值来源：sample / all（已全量扫描）
        self._scan_all_token = 0      # 全量取值扫描序号：防过期结果回填
        self._scanning_all = False    # 全量取值扫描进行中
        self._tpl_kinds = (0, 0)      # 全量扫描看出的 (几种模板, 几个表)；不止一种时取值摘要里提一句
        self._bulk_sync = False       # 全选/清空批量改勾选：合并成一次同步
        # 「已限定名单」是一个**显式状态**，不能靠「是否全勾」反推：样本表里恰好
        # 就是用户要的那几组时，全勾既可能是「不筛选」也可能是「只要这几组」。
        # 早先靠勾选数反推，用户点完「只拆这 N 组」再随手取消一个又勾回来，
        # 名单就被悄悄清空、退回拆全部——正是这次要修的名实不符。**别改回反推。**
        self._values_locked = False
        self._lock_btn = None
        self._recipe_active = None     # 套用过的方案：方案里的字段在表里找不到时只提醒，不自动换成推荐列         # 勾选面板里的「只拆这 N 组」按钮（面板开着才有）
        self._values_hint = None      # 勾选面板里的提示文字（随锁定/抽样状态变）
        self._warn_count = 0          # 本轮运行日志里 ⚠ 行数（完成后给「处理详情」打角标）
        self._last_summary = None     # core 收尾的 [SUMMARY] JSON（完成摘要用）
        self._parked = {}             # 切走模式时收起的上一轮结果：{模式: (露着的控件名, 状态行文字, 颜色)}
        self._pdf_grid_rows = 0      # 映射清单网格行数（分发预览用）
        self._pdf_n_cols = 0          # 映射清单列数（分发预览用）
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
        if self._first_run and not p:
            # 首启一句话引导（D3）：三步卡片本身就是向导，不再做独立新手页
            tip = "点「📄 选一个 Excel 文件」"
            if self._dnd_ok:
                tip += "，或直接把文件 / 文件夹拖进窗口"
            self._set_in_status(f"第一次用？{tip}即可开始", C_MUTED)
        for _m in self._pending_logs:      # 排空启动期攒下的日志（此时 _log_box 已就绪）
            self._log(_m)
        self._pending_logs = []
        self.after(100, self._pump_ui)

    def _init_geometry(self):
        """窗口初始尺寸：优先恢复上次关闭时的大小位置；首次按屏幕高度自适应。

        目标用户大量使用 1366×768 的老笔记本，固定 820×800 会让窗口下沿
        （含主按钮的操作区）超出屏幕——这是小屏适配的实际 bug，高度必须
        按可用屏高压下来（滚动区会负责吸收内容）。
        """
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        h = min(800, max(600, sh - 90))
        default = f"820x{h}"
        saved = str(self.cfg.get("window_geometry", ""))
        m = re.fullmatch(r"(\d+)x(\d+)([+-]\d+)([+-]\d+)", saved)
        if m and int(m.group(1)) <= sw and int(m.group(2)) <= sh:
            self.geometry(saved)
        else:
            self.geometry(default)
        self.minsize(720, min(640, h))

    def _on_close(self):
        """关闭时静默记住窗口大小位置（下次启动恢复现场），失败不影响退出。"""
        try:
            cfg = self._collect_config()
            cfg["window_geometry"] = self.geometry()
            cfg["ui_scale"] = self.cfg.get("ui_scale", 100)
            save_config(cfg)
        except Exception:
            pass
        self.destroy()

    def _set_window_icon(self):
        """设置标题栏 / 任务栏 / 最小化时的窗口图标。

        `--icon app.ico`（build.bat / CI）只改 exe 文件本身的图标，运行起来的
        Tk 窗口图标要靠 iconbitmap() 单独设置，否则标题栏/任务栏还是 Tk 默认
        羽毛图标——这就是“exe 图标换了，但程序里和最小化时的图标没跟着换”的
        原因。app.ico 需要被 --add-data 打包进去才能在冻结后找到（见 build.bat）。

        Linux（X11）上 iconbitmap() 只认 XBM 位图，传 .ico 必抛 TclError，
        所以还要用 app.png + iconphoto 兜底（同样需要 --add-data 打包进去）。
        """
        bases = (_resource_dir(), _app_dir())
        set_window_icon(self,
                        [os.path.join(b, "app.ico") for b in bases],
                        [os.path.join(b, "app.png") for b in bases])

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
        self._collect_frame = ctk.CTkFrame(self._body, fg_color="transparent")
        self._collect_frame.grid(row=0, column=0, sticky="ew")
        self._collect_frame.grid_columnconfigure(0, weight=1)
        self._desens_frame = ctk.CTkFrame(self._body, fg_color="transparent")
        self._desens_frame.grid(row=0, column=0, sticky="ew")
        self._desens_frame.grid_columnconfigure(0, weight=1)
        self._build_step_input(self._excel_frame)              # ① 选表格
        self._build_step_field(self._excel_frame)              # ② 选字段（含批量选项）
        self._build_adv_area(self._excel_frame)                # ▸ 高级设置
        self._build_pdf_step_input(self._pdf_frame)            # PDF ① 选 PDF
        self._build_pdf_step_map(self._pdf_frame)              # PDF ② 选映射清单
        self._build_collect_steps(self._collect_frame)         # 汇总 ①② + 结果卡片
        self._build_desens_steps(self._desens_frame)           # 脱敏 ①② + 结果卡片
        self._build_action(self)                               # row 2 ③ 开始（固定底部）
        self._build_bottom(self)                               # row 3-5 工具条 / 日志 / 页脚

    def _build_header(self):
        # 品牌文案已定稿（副标/标语/特性行），只调整排版：作者信息移到页脚
        head = ctk.CTkFrame(self, fg_color="transparent")
        head.grid(row=0, column=0, padx=20, pady=(14, 4), sticky="ew")
        brand = ctk.CTkFrame(head, fg_color="transparent")
        brand.pack(anchor="w")
        # 品牌色块（v2.7.1）：标题左侧的圆角主色标记，最轻量的品牌记忆点
        ctk.CTkFrame(brand, width=6, height=30, corner_radius=3,
                     fg_color=PRIMARY).pack(side="left", padx=(0, 10), pady=(2, 0))
        ctk.CTkLabel(brand, text="ExcelRouter",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(side="left")
        ctk.CTkLabel(brand, text="Excel 智能拆分汇总工具",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=("gray25", "gray75")).pack(side="left", padx=(10, 0), pady=(7, 0))
        ctk.CTkLabel(head, text="一键按部门、网格、工号拆分并打包分发；收回来一键汇总，还能查出谁没交、谁漏填",
                     font=ctk.CTkFont(size=12),
                     text_color=("gray30", "gray70")).pack(anchor="w", pady=(4, 0))
        ctk.CTkLabel(head, text="保留原格式 · 多表合并汇总 · 查收通报 · 文件加密分发",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(anchor="w", pady=(1, 0))
        # 功能模式切换：Excel 拆分（默认） / Excel 汇总（可选先查收） / 文件按网格加密分发
        self._mode_seg = ctk.CTkSegmentedButton(head, values=list(MODE_LABELS.values()),
                                                command=self._on_mode_change,
                                                selected_color=PRIMARY,
                                                selected_hover_color=PRIMARY_HOVER,
                                                height=30, font=ctk.CTkFont(size=12))
        self._mode_seg.set(MODE_LABELS[self._mode])
        self._mode_seg.pack(anchor="w", pady=(8, 0))

    def _step_card(self, parent, row, num, title, padx=8, pady=(0, 10)):
        """带编号徽章的步骤卡片：编号即真实操作顺序，是界面的导航主线。

        返回 (card, 标题 Label)：标题引用给需要随模式改文案的卡片用（如③操作卡）。
        """
        card = ctk.CTkFrame(parent, corner_radius=12, border_width=1,
                            border_color=CARD_BORDER)
        card.grid(row=row, column=0, padx=padx, pady=pady, sticky="ew")
        card.grid_columnconfigure(0, weight=1)
        bar = ctk.CTkFrame(card, fg_color="transparent")
        bar.grid(row=0, column=0, padx=12, pady=(10, 2), sticky="w")
        ctk.CTkLabel(bar, text=str(num), width=26, height=26, corner_radius=13,
                     fg_color=PRIMARY, text_color="white",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        title_lbl = ctk.CTkLabel(bar, text=title,
                                 font=ctk.CTkFont(size=14, weight="bold"))
        title_lbl.pack(side="left", padx=(8, 0))
        return card, title_lbl

    # ── 拖拽（v2.7，tkinterdnd2 可选依赖）────────────────────
    def _enable_drop(self, widget, handler):
        """给控件注册文件拖放；tkinterdnd2 缺失时静默跳过（不影响其余功能）。"""
        if not self._dnd_ok:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", handler)
        except Exception:
            pass

    @staticmethod
    def _parse_drop_paths(data):
        """解析 tkdnd 的拖放数据（Windows 原生路径 / X11 的 file:// URI 都能吃）。"""
        return parse_drop_paths(data)

    # ── 模式切换 ─────────────────────────────────────────
    # ③卡片里只属于上一轮的结果控件（完成摘要、打开输出文件夹、存方案提示、复制日志）
    _RESULT_WIDGETS = ("_summary_banner", "_open_btn", "_recipe_tip", "_copylog_btn")

    def _on_mode_change(self, label):
        """切模式。上一轮的结果只属于跑它的那个模式：切走时收起、切回来原样放出。

        不收起的话，拆分的完成摘要会挂在「开始分发」下面（对不上），而且③卡片被摘要
        撑高后，上面的步骤区会被挤到几乎没有，再带着原来的滚动位置——看上去只剩第 3 步。
        """
        old = self._mode
        self._mode = next((k for k, v in MODE_LABELS.items() if v == label), "excel")
        switched = old != self._mode and not self._running
        if switched:
            shown = tuple(n for n in self._RESULT_WIDGETS if getattr(self, n).winfo_manager())
            for n in shown:
                getattr(self, n).grid_remove()
            self._parked[old] = (shown, self._run_status.cget("text"), self._run_status.cget("text_color"))
        self._apply_mode()
        if switched and self._mode in self._parked:
            shown, text, color = self._parked.pop(self._mode)
            for n in shown:
                getattr(self, n).grid()
            self._run_status.configure(text=text, text_color=color)
        # 切回有结果的模式就停在结果上（结果在步骤下面，等布局刷新完再拉到底），其余从头看起
        if self._summary_banner.winfo_manager():
            self.after(150, lambda: self._body._parent_canvas.yview_moveto(1.0))
        else:
            self._body._parent_canvas.yview_moveto(0)

    def _start_text(self):
        if self._mode == "collect":
            return "🔍 查收（看谁没交、谁漏填）" if self._col_check_var.get() else "▶ 开始汇总"
        return {"pdf": "🔐 开始分发", "desens": "🛡 开始脱敏"}.get(self._mode, "▶ 开始拆分")

    def _apply_mode(self):
        """按当前模式显示对应步骤区并更新③卡片文案；运行中禁止切换按钮已够复杂，
        简单起见运行中也允许切界面，但开始按钮在运行中本来就是禁用态。"""
        for fr in (self._excel_frame, self._pdf_frame, self._collect_frame, self._desens_frame):
            fr.grid_remove()
        # Excel 汇总的结果固定放在输入旁的「汇总结果」：收起保存位置；拆分速度只有拆分模式有意义
        for w, show in ((self._out_row, self._mode not in ("collect", "desens")),
                        (self._speed_row, self._mode == "excel"),
                        (self._speed_hint, self._mode == "excel")):
            if show:
                w.grid()
            else:
                w.grid_remove()
        if self._mode == "collect":
            self._collect_frame.grid()
            check = self._col_check_var.get()
            self._action_title.configure(text="查收" if check else "开始汇总")
            if not self._running:
                self._start_btn.configure(text=self._start_text())
                self._run_status.configure(
                    text="收回几份就点一次「查收」；交齐后在结果里点「汇总成一张」" if check else
                    "选好表格直接点「开始汇总」：同一套模板的表合成一张，每行标上来自哪个文件，列顺序不同也能对齐",
                    text_color=C_MUTED)
        elif self._mode == "desens":
            self._desens_frame.grid()
            self._action_title.configure(text="开始脱敏")
            if not self._running:
                self._start_btn.configure(text="🛡 开始脱敏")
                self._run_status.configure(
                    text="选好表格和用途后点「开始脱敏」：结果放在表格旁边的「脱敏结果」，原表不改，处理完自动复检",
                    text_color=C_MUTED)
        elif self._mode == "pdf":
            self._pdf_frame.grid()
            self._action_title.configure(text="开始分发")
            if not self._running:
                self._start_btn.configure(text="🔐 开始分发")
                self._run_status.configure(
                    text="完成 ① ② 后点「开始分发」；每个网格生成一份带专属密码 / 水印的副本，并输出分发清单。"
                         "只是个别列有号码、证件号？先用「数据脱敏」打码，发出去就不用再发密码",
                    text_color=C_MUTED)
        else:
            self._excel_frame.grid()
            self._action_title.configure(text="开始拆分")
            if not self._running:
                self._start_btn.configure(text="▶ 开始拆分")
                self._run_status.configure(
                    text="完成 ① ② 后直接点「开始拆分」；保存位置不用改，结果会放进自动创建的「拆分结果」文件夹",
                    text_color=C_MUTED)

    def _build_action(self, parent):
        card, self._action_title = self._step_card(parent, 2, 3, "开始拆分", padx=20, pady=(6, 0))

        orow = self._out_row = ctk.CTkFrame(card, fg_color="transparent")
        orow.grid(row=1, column=0, padx=12, pady=(4, 2), sticky="ew")
        orow.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(orow, text="结果保存到").grid(row=0, column=0, sticky="w")
        self._output_var = ctk.StringVar(value=self.cfg.get("output_path", ""))
        ctk.CTkEntry(orow, textvariable=self._output_var).grid(row=0, column=1, padx=8, sticky="ew")
        _ghost_button(orow, text="浏览", width=60,
                      command=self._browse_output).grid(row=0, column=2)

        # 速度模式（v2.9）：从「高级设置 → 保留格式」提到主操作区。
        # 这是唯一能把耗时拉开数倍的选项，藏在折叠面板里等于不存在——用户只会
        # 得出「这工具拆得慢」的结论，而不会去翻高级设置找到它（真实反馈）。
        # 这里是 preserve_format 的**唯一**界面入口，高级设置里的那个勾选框已移除，
        # 别再加回去：同一个开关两处控制，用户永远不确定哪个生效。
        srow = self._speed_row = ctk.CTkFrame(card, fg_color="transparent")
        srow.grid(row=2, column=0, padx=12, pady=(8, 0), sticky="w")
        ctk.CTkLabel(srow, text="拆分速度").pack(side="left")
        self._speed_seg = ctk.CTkSegmentedButton(
            srow, values=[SPEED_STD, SPEED_FAST], command=self._on_speed_change,
            selected_color=PRIMARY, selected_hover_color=PRIMARY_HOVER,
            height=30, font=ctk.CTkFont(size=12))
        self._speed_seg.set(SPEED_STD if self._preserve_var.get() else SPEED_FAST)
        self._speed_seg.pack(side="left", padx=(8, 0))
        self._speed_hint = ctk.CTkLabel(card, text="", font=ctk.CTkFont(size=11),
                                        text_color=C_MUTED, anchor="w", justify="left",
                                        wraplength=640)
        self._speed_hint.grid(row=3, column=0, padx=12, pady=(2, 0), sticky="w")
        self._on_speed_change(self._speed_seg.get())

        btns = ctk.CTkFrame(card, fg_color="transparent")
        btns.grid(row=4, column=0, padx=12, pady=(8, 2), sticky="w")
        # 主按钮是全页唯一核心动作（v2.7.1）：加高配粗 + 品牌主色，保证第一眼锁定
        self._start_btn = ctk.CTkButton(btns, text="▶ 开始拆分", width=220, height=44,
                                        corner_radius=10,
                                        font=ctk.CTkFont(size=15, weight="bold"),
                                        fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
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

        self._progress = ctk.CTkProgressBar(card, corner_radius=6,
                                           progress_color=PRIMARY)
        self._progress.grid(row=5, column=0, padx=12, pady=(8, 2), sticky="ew")
        self._progress.set(0)
        self._run_status = ctk.CTkLabel(card, text="完成 ① ② 后直接点「开始拆分」；保存位置不用改，结果会放进自动创建的「拆分结果」文件夹",
                                        font=ctk.CTkFont(size=11), text_color=C_MUTED, anchor="w")
        self._run_status.grid(row=6, column=0, padx=12, pady=(0, 2), sticky="w")
        # 完成摘要（v2.7，A4；v2.7.1 横幅化）：浅绿底圆角横幅承载「完成瞬间」的
        # 明确成功反馈——摘要内容与逻辑不变，只加视觉承载。
        # 放在滚动区里、步骤下面（和汇总的结果卡片一样），不放进固定的③卡片：
        # 摘要有五六行，放在③里会把上面的步骤区挤到只剩十几像素，老笔记本上看着就只剩第 3 步
        self._summary_banner = ctk.CTkFrame(self._body, corner_radius=10,
                                            fg_color=OK_BANNER_BG)
        self._summary_banner.grid(row=1, column=0, padx=8, pady=(0, 6), sticky="ew")
        self._summary_banner.grid_remove()
        self._summary_lbl = ctk.CTkLabel(self._summary_banner, text="",
                                         font=ctk.CTkFont(size=12, weight="bold"),
                                         text_color=C_OK, anchor="w", justify="left")
        self._summary_lbl.grid(row=0, column=0, padx=14, pady=8, sticky="w")
        # 跑完后的小提示：这次的设置存成方案，下个月一键套用（v3.0）
        self._recipe_tip = ctk.CTkFrame(self._body, fg_color="transparent")
        self._recipe_tip.grid(row=2, column=0, padx=4, pady=(0, 6), sticky="w")
        _flat_button(self._recipe_tip, text="📋 把这次的设置存成方案，下个月一键用", width=260, anchor="w",
                     command=self._save_recipe).pack(side="left")
        _flat_button(self._recipe_tip, text="不再提示", width=70, command=self._recipe_tip_never).pack(side="left")
        self._recipe_tip.grid_remove()

    # ── 方案（v3.0）：把一次设置存成小文件，下个月一键套用、发给同事也能用 ──
    def _recipe_dir(self):
        from core.recipe import RECIPE_DIR
        return os.path.join(_app_dir(), RECIPE_DIR)

    def _recipe_menu(self):
        """页脚「📋 方案」：存为方案 / 打开方案 / 最近用过的。"""
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="存为方案…", command=self._save_recipe)
        menu.add_command(label="打开方案…", command=self._open_recipe)
        recent = [p for p in self.cfg.get("recent_recipes", []) if isinstance(p, str) and os.path.isfile(p)]
        if recent:
            menu.add_separator()
            for p in recent:
                menu.add_command(label=os.path.splitext(os.path.basename(p))[0],
                                 command=lambda p=p: self._apply_recipe(p))
        btn = self._recipe_btn
        try:
            menu.tk_popup(btn.winfo_rootx(), btn.winfo_rooty() - 4 - 22 * (len(recent) + 3))
        finally:
            menu.grab_release()

    def _remember_recipe(self, path=None, tip_off=None):
        """最近用过的方案（最多 5 个）和「不再提示」写进配置文件；不动上次成功运行记下的其余设置。"""
        from core.recipe import RECENT_MAX
        if path:
            path = os.path.normpath(path)
            recent = [path] + [p for p in self.cfg.get("recent_recipes", []) if os.path.normpath(p) != path]
            self.cfg["recent_recipes"] = recent[:RECENT_MAX]
        if tip_off is not None:
            self.cfg["recipe_tip_off"] = tip_off
        try:
            saved = load_config()
            saved["recent_recipes"] = self.cfg.get("recent_recipes", [])
            saved["recipe_tip_off"] = self.cfg.get("recipe_tip_off", False)
            save_config(saved)
        except Exception:
            pass

    def _recipe_input_name(self):
        if self._mode == "collect":
            return self._col_in_var.get().strip()
        if self._mode == "desens":
            return self._ds_in_var.get().strip()
        if self._mode == "pdf":
            return self._pdf_map_var.get().strip()
        return self._input_var.get().strip()

    def _save_recipe(self):
        from core import recipe as R
        cfg = self._collect_config()
        base = os.path.basename(os.path.normpath(self._recipe_input_name())) if self._recipe_input_name() else ""
        default = os.path.splitext(base)[0] or R.MODE_NAMES[self._mode]
        os.makedirs(self._recipe_dir(), exist_ok=True)
        path = filedialog.asksaveasfilename(title="存为方案", initialdir=self._recipe_dir(),
                                            initialfile=R.recipe_filename(default),
                                            defaultextension=".json", filetypes=[("方案", "*.json")])
        if not path:
            return
        name = os.path.splitext(os.path.basename(path))[0]
        cols = self._columns if self._mode == "excel" else None
        try:
            rec = R.make_recipe(cfg, self._mode, name, columns=cols, input_name=self._recipe_input_name(),
                                app=APP_VERSION)
            R.save_recipe(rec, path)
        except Exception as e:
            messagebox.showerror("没存成", f"方案没存成：{e}")
            return
        self._remember_recipe(path)
        self._recipe_tip.grid_remove()
        tip = ""
        if rec["settings"].get("collect_expect"):
            from core.collect import split_names
            tip = f"（里面有应交名单 {len(split_names(rec['settings']['collect_expect']))} 个，发给同事前确认可以给对方看）"
        self._log(f"📋 已存为方案「{name}」：{R.summary(rec)}。文件在：{path}")
        self._run_status.configure(text=f"📋 已存为方案「{name}」，下次点「📋 方案」就能一键套用{tip}",
                                   text_color=C_OK)

    def _open_recipe(self):
        path = filedialog.askopenfilename(title="打开方案", initialdir=self._recipe_dir()
                                          if os.path.isdir(self._recipe_dir()) else None,
                                          filetypes=[("方案", "*.json")])
        if path:
            self._apply_recipe(path)

    def _drop_recipe(self, paths):
        """拖进来的是方案文件就打开它，返回 True；不是方案的 .json（如下发记录）返回 False。"""
        from core.recipe import is_recipe
        p = next((p for p in paths if p.lower().endswith(".json") and is_recipe(p)), None)
        if p:
            self._apply_recipe(p)
            return True
        return False

    def _apply_recipe(self, path):
        """套用方案：切到对应模式、填好各个控件；已经选了表就当场比一比列对不对得上。路径不动，截止时间清空。"""
        from core import recipe as R
        if self._running:
            messagebox.showinfo("正在运行", "等这一轮跑完再套用方案")
            return
        try:
            rec, warns = R.load_recipe(path)
        except R.RecipeError as e:
            messagebox.showerror("方案打不开", str(e))
            return
        new, _changed, notes = R.apply_recipe(self._collect_config(), rec)
        self.cfg.update(new)
        self._fill_from_recipe(rec["settings"], rec["mode"])
        self._recipe_active = rec
        self._mode_seg.set(MODE_LABELS[rec["mode"]])
        self._on_mode_change(MODE_LABELS[rec["mode"]])   # 和点模式按钮走同一条路（上一轮结果跟着收起）
        self._remember_recipe(path)
        self._log(f"📋 已套用方案「{rec['name']}」（{R.MODE_NAMES[rec['mode']]}）：{R.summary(rec)}")
        for w in warns + notes:
            self._log(f"  ℹ️ {w}")
        msg = f"📋 已套用方案「{rec['name']}」"
        if rec["mode"] == "excel" and self._columns:
            self._on_columns(self._columns)          # 已经选了表：当场检查方案里的字段在不在
        elif rec["mode"] == "excel":
            msg += "：现在选这个月的表格就行"
        elif rec["mode"] == "desens":
            msg += "：选这次要脱敏的表格"
            if self._ds_info:
                self._render_desens_cols()
        elif rec["mode"] == "collect":
            msg += "：选收回来的文件夹，截止时间请重新填"
            p = self._col_in_var.get().strip()
            if p and os.path.exists(p):
                self._col_scanned = None
                self._scan_collect()
        else:
            msg += "：选这次要分发的文件和清单"
        self._run_status.configure(text=msg, text_color=C_OK)

    def _fill_from_recipe(self, st, mode):
        """把方案里的设置填进各个控件（只填方案里有的键）。"""
        def has(k):
            return k in st
        if mode == "excel":
            if has("split_column"):
                self._split_var.set(st["split_column"] or COL_PLACEHOLDER)
            if has("person_column"):
                self._pcol_var.set(st["person_column"] or COL_PLACEHOLDER)
            if has("to_person"):
                self._person_var.set(st["to_person"])
                self._update_person_state()
            if has("selected_values"):
                self._selvals_var.set(", ".join(st["selected_values"]))
                self._values_locked = bool(st["selected_values"])
            for k, var in (("grid_keys", self._grid_keys_var), ("id_keys", self._id_keys_var),
                           ("person_file_filter", self._pfilter_var), ("skip_values", self._skip_var)):
                if has(k):
                    var.set(", ".join(st[k]))
            for k, var in (("make_zip", self._zip_var), ("merge_across_files", self._merge_var),
                           ("exact_match", self._exact_var), ("keep_hidden", self._keep_hidden_var)):
                if has(k):
                    var.set(st[k])
            if has("preserve_format"):
                label = SPEED_STD if st["preserve_format"] else SPEED_FAST
                self._speed_seg.set(label)
                self._on_speed_change(label)
            if has("keep_formulas") and self._preserve_var.get():
                self._keep_formula_var.set(st["keep_formulas"])
            if has("value_alias_map"):
                self._alias_box.delete("1.0", "end")
                if st["value_alias_map"]:
                    self._alias_box.insert("1.0", json.dumps(st["value_alias_map"], ensure_ascii=False, indent=2))
            if has("fill_columns"):
                self._fill_var.set(st["fill_columns"])
            self._due_var.set("")
            if has("renumber_seq"):
                self._renumber_var.set(st["renumber_seq"])
            if has("group_totals"):
                self._totals_var.set(st["group_totals"])
            if has("split_desens") and self._split_ds_cols:
                self._split_ds_var.set(bool(st["split_desens"]))
                if st["split_desens"]:
                    self._split_ds_scene.set(st["split_desens"])
                self._on_split_desens_toggle()
        if mode == "desens":
            if has("desens_scene") and st["desens_scene"] in DS_SCENES:
                self._ds_scene_seg.set(st["desens_scene"])
                self._on_desens_scene(st["desens_scene"])
            if has("desens_notes"):
                self._ds_notes_var.set(st["desens_notes"])
            if has("desens_columns"):
                self._ds_extra_var.set(st["desens_columns"])
            if has("keep_hidden"):
                self._keep_hidden_var.set(st["keep_hidden"])
        if mode in ("excel", "collect", "desens"):
            if has("header_mode"):
                label = {"auto": "自动", "row": "指定行号", "keyword": "关键词"}[st["header_mode"]]
                self._header_seg.set(label)
                self._on_header_mode(label)
            if has("header_row"):
                self._hrow_var.set(str(st["header_row"]))
        if mode == "collect":
            for k, var in (("collect_expect", self._col_expect_var), ("collect_required", self._col_req_var),
                           ("collect_by", self._col_by_var), ("collect_title", self._col_title_var)):
                if has(k):
                    var.set(st[k])
            self._col_due_var.set("")
            self._col_rec_var.set("")
            if has("collect_check"):
                self._col_check_var.set(st["collect_check"])
                self._on_check_toggle()
        if mode == "pdf":
            for k, var in (("pdf_grid_column", self._pdfgrid_var), ("pdf_password_column", self._pdfpwd_var)):
                if has(k):
                    var.set(st[k] or PDF_COL_PLACEHOLDER)
            for k, var in (("pdf_encrypt", self._pdfenc_var), ("pdf_watermark", self._pdfwm_var)):
                if has(k):
                    var.set(st[k])
            if not self._pdfenc_var.get() and not self._pdfwm_var.get():
                self._pdfenc_var.set(True)
            if has("pdf_watermark_text"):
                self._pdfwmtext_var.set(st["pdf_watermark_text"])

    def _offer_recipe(self, cfg):
        """成功跑完一次：结果区出现「存成方案」的小链接。当前设置已经是最近某个方案、或点过「不再提示」时不出现。"""
        from core import recipe as R
        mode = cfg.get("ui_mode", "excel")
        if self.cfg.get("recipe_tip_off") or mode not in R.RECIPE_KEYS:
            return
        mine = R.make_recipe(cfg, mode, "x")["settings"]
        for p in self.cfg.get("recent_recipes", []):
            try:
                rec, _w = R.load_recipe(p)
            except (R.RecipeError, TypeError):
                continue
            if rec["mode"] == mode and rec["settings"] == mine:
                return
        self._recipe_tip.grid()

    def _recipe_tip_never(self):
        self._recipe_tip.grid_remove()
        self._remember_recipe(tip_off=True)

    def _build_bottom(self, parent):
        bar = ctk.CTkFrame(parent, fg_color="transparent")
        bar.grid(row=3, column=0, padx=20, pady=(4, 0), sticky="ew")
        bar.grid_columnconfigure(1, weight=1)
        self._log_btn = _flat_button(bar, text="▸ 处理详情", width=100, anchor="w",
                                     command=self._toggle_log)
        self._log_btn.grid(row=0, column=0, sticky="w")
        self._copylog_btn = _flat_button(bar, text="📋 复制日志", width=90, anchor="w",
                                         command=self._copy_log)
        self._copylog_btn.grid(row=0, column=1, sticky="w", padx=(4, 0))
        self._copylog_btn.grid_remove()      # 只在失败时出现（配合反馈问卷降低求助门槛）
        right = ctk.CTkFrame(bar, fg_color="transparent")
        right.grid(row=0, column=2, sticky="e")
        self._recipe_btn = _flat_button(right, text="📋 方案", width=70, command=self._recipe_menu)
        self._recipe_btn.pack(side="left")
        _flat_button(right, text="❓ 常见问题", width=90,
                     command=lambda: webbrowser.open(FAQ_URL)).pack(side="left")
        _flat_button(right, text="💬 反馈建议", width=90,
                     command=self._open_feedback).pack(side="left")
        self._appearance_btn = _flat_button(right, text=self._appearance_btn_text(),
                                            width=96, command=self._cycle_appearance)
        self._appearance_btn.pack(side="left", padx=(6, 0))
        self._scale_btn = _flat_button(
            right, text=self._scale_btn_text(), width=90, command=self._toggle_scale)
        self._scale_btn.pack(side="left", padx=(6, 0))
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

    def _log_btn_text(self):
        """日志折叠钮文案：本轮有 ⚠ 提醒时带角标，把藏在折叠日志里的警告提到可见。"""
        arrow = "▾" if self._log_open else "▸"
        badge = f"（⚠ {self._warn_count}）" if self._warn_count else ""
        return f"{arrow} 处理详情{badge}"

    def _toggle_log(self, show=None):
        show = (not self._log_open) if show is None else show
        if show == self._log_open:
            return
        self._log_open = show
        if show:
            self._log_box.grid(row=4, column=0, padx=20, pady=(2, 0), sticky="ew")
        else:
            self._log_box.grid_remove()
        self._log_btn.configure(text=self._log_btn_text())

    def _browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self._output_var.set(os.path.normpath(path))

    # ── 配置收集 ───────────────────────────────────────
    def _split_list(self, text):
        return [k.strip() for k in text.split(",") if k.strip()]

    def _collect_config(self):
        mode_map = {"自动": "auto", "指定行号": "row", "关键词": "keyword"}
        try:
            header_row = int(self._hrow_var.get())
        except (TypeError, ValueError):
            header_row = 1

        split_col = "" if self._split_var.get() == COL_PLACEHOLDER else self._split_var.get()
        pcol = "" if self._pcol_var.get() == COL_PLACEHOLDER else self._pcol_var.get()
        alias, _ = self._parse_alias()

        return {
            "input_path":  self._input_var.get().strip(),
            "output_path": self._output_var.get().strip(),
            "ui_mode":     self._mode,
            "pdf_input_paths": list(self._pdf_paths),
            "pdf_mapping_path": self._pdf_map_var.get().strip(),
            "pdf_grid_column": "" if self._pdfgrid_var.get() == PDF_COL_PLACEHOLDER else self._pdfgrid_var.get(),
            "pdf_password_column": "" if self._pdfpwd_var.get() == PDF_COL_PLACEHOLDER else self._pdfpwd_var.get(),
            "pdf_receiver_column": "",            # v2.10 起由 core 按列名自动认
            "pdf_encrypt": self._pdfenc_var.get(),
            "pdf_watermark": self._pdfwm_var.get(),
            "pdf_watermark_text": self._pdfwmtext_var.get().strip() or "{grid} {date}",
            "pdf_watermark_opacity": self.cfg.get("pdf_watermark_opacity", 0.15),
            "pdf_watermark_angle": self.cfg.get("pdf_watermark_angle", 45),
            "collect_input_path": self._col_in_var.get().strip(),
            "collect_expect": self._col_expect_var.get().strip(),
            "collect_required": self._col_req_var.get().strip(),
            "collect_by": self._col_by_var.get().strip(),
            "collect_title": self._col_title_var.get().strip(),
            "collect_due": self._col_due_var.get().strip(),
            "collect_record": self._col_rec_var.get().strip(),
            "collect_check": self._col_check_var.get(),
            "fill_columns": self._fill_var.get().strip(),
            "due": self._due_var.get().strip(),
            "split_desens": self._split_desens_value(),
            "renumber_seq": self._report_value("renumber_seq"),
            "group_totals": self._report_value("group_totals"),
            "desens_input_path": self._ds_in_var.get().strip(),
            "desens_scene": self._ds_scene_seg.get(),
            "desens_notes": self._ds_notes_var.get(),
            "desens_columns": self._ds_extra_var.get().strip(),
            # 界面个性化键随每次保存透传，避免跑一次成功后丢失（与 pdf 水印参数同理）
            "ui_scale": self.cfg.get("ui_scale", 100),
            "window_geometry": self.cfg.get("window_geometry", ""),
            "appearance_mode": self.cfg.get("appearance_mode", "system"),
            "recent_recipes": self.cfg.get("recent_recipes", []),
            "recipe_tip_off": self.cfg.get("recipe_tip_off", False),
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
            "keep_hidden": self._keep_hidden_var.get(),
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
        values: Any = _MISSING
        scanall: Any = _MISSING
        scanall_prog: Any = _MISSING
        col_scan: Any = _MISSING
        ds_scan: Any = _MISSING
        split_ds: Any = _MISSING
        rpt_hints: Any = _MISSING
        ds_copy: Any = _MISSING
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
                elif kind == "values":
                    values = payload
                elif kind == "scanall":
                    scanall = payload
                elif kind == "scanall_prog":
                    scanall_prog = payload
                elif kind == "collect_scan":
                    col_scan = payload
                elif kind == "desens_scan":
                    ds_scan = payload
                elif kind == "split_desens":
                    split_ds = payload
                elif kind == "report_hints":
                    rpt_hints = payload
                elif kind == "desens_copy":
                    ds_copy = payload
        except queue.Empty:
            pass
        if logs:
            # [SUMMARY] 机器可读行不直接展示：提取出来供完成摘要渲染（v2.7）；
            # 提取失败也不报错，退回原有纯日志行为
            out_lines = []
            for payload in logs:
                for line in str(payload).splitlines():
                    if line.startswith("[SUMMARY] "):
                        try:
                            self._last_summary = json.loads(line[len("[SUMMARY] "):])
                        except Exception:
                            pass
                        continue
                    if "⚠" in line:
                        self._warn_count += 1
                    out_lines.append(line)
            if out_lines:
                joined = "\n".join(out_lines)
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
        if values is not _MISSING:
            self._on_values(values)
        if scanall_prog is not _MISSING:
            self._on_scan_all_prog(scanall_prog)
        if scanall is not _MISSING:
            self._on_scan_all(scanall)
        if col_scan is not _MISSING:
            self._on_collect_scan(col_scan)
        if ds_scan is not _MISSING:
            self._on_desens_scan(ds_scan)
        if split_ds is not _MISSING:
            self._on_split_desens(split_ds)
        if rpt_hints is not _MISSING:
            self._on_report_hints(rpt_hints)
        if ds_copy is not _MISSING:
            self._on_desens_copy(ds_copy)
        if done is not _MISSING:
            self._on_done(*done)
        self.after(100, self._pump_ui)

    def _enter_running(self, busy_text, first_log):
        """进入运行态的公共 UI 准备：两种模式共用（按钮/进度/日志/状态行）。"""
        self._stop_flag = False
        self._running = True
        self._last_output = None
        self._warn_count = 0
        self._last_summary = None
        self._parked.clear()          # 结果控件要给这一轮用了，别的模式收起的旧结果作废
        self._summary_lbl.configure(text="")
        self._summary_banner.grid_remove()
        self._recipe_tip.grid_remove()
        self._copylog_btn.grid_remove()
        self._log_btn.configure(text=self._log_btn_text())
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

    def _stop(self):
        self._stop_flag = True
        self._log("正在停止，请等待当前文件处理完成...")

    def _copy_log(self):
        """把日志全文复制进剪贴板，方便粘贴到反馈问卷或发给维护者。"""
        try:
            self.clipboard_clear()
            self.clipboard_append(self._log_box.get("1.0", "end").strip())
            self._run_status.configure(text="📋 日志已复制，可直接粘贴到反馈问卷", text_color=C_OK)
        except Exception:
            pass

    def _appearance_btn_text(self):
        return {"system": "🌓 跟随系统", "light": "☀️ 浅色", "dark": "🌙 深色"}.get(
            self.cfg.get("appearance_mode", "system"), "🌓 跟随系统")

    def _cycle_appearance(self):
        """深浅色循环切换：系统 → 浅色 → 深色 → 系统；即时生效并持久化。"""
        order = ["system", "light", "dark"]
        cur = self.cfg.get("appearance_mode", "system")
        new = order[(order.index(cur) + 1) % len(order)]
        self.cfg["appearance_mode"] = new
        ctk.set_appearance_mode(new if new != "system" else "System")
        self._appearance_btn.configure(text=self._appearance_btn_text())
        try:
            save_config(self._collect_config())
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存外观设置：{e}")

    def _scale_btn_text(self):
        return "🔍 标准字号" if self.cfg.get("ui_scale", 100) != 100 else "🔍 大字号"

    def _toggle_scale(self):
        """大字号即时切换（v2.7.1 改）：set_widget_scaling 对已建控件即时生效，
        点了立刻看到变大/复原；115% 在当前布局余量内，无需重建界面。"""
        new = 100 if self.cfg.get("ui_scale", 100) != 100 else 115
        self.cfg["ui_scale"] = new
        ctk.set_widget_scaling(new / 100)
        self._scale_btn.configure(text=self._scale_btn_text())
        try:
            save_config(self._collect_config() | {"ui_scale": new})
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存设置：{e}")

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
        if not path:
            return
        if not os.path.exists(path):
            self._run_status.configure(text=f"⚠ 这个目录不在了：{path}", text_color=C_WARN)
            return
        if not open_folder(path):
            # 旧实现失败时悄无声息：用户点了按钮什么也没发生，比报错更让人困惑
            self._run_status.configure(text=f"⚠ 打不开文件管理器，结果就在：{path}",
                                       text_color=C_WARN)
            self._log(f"⚠ 无法调起文件管理器，请手动打开：{path}")

    def _render_summary(self, output_path, is_pdf):
        """完成摘要（A4）：用 core 收尾的 [SUMMARY] 渲染多行结果概览。

        摘要行是文本契约（[SUMMARY] + JSON），解析失败/缺字段时逐字段容错，
        哪项拿不到就不显示哪项，绝不影响完成提示本身。
        """
        s = self._last_summary or {}
        lines, has_warn = [], False
        if is_pdf:
            if s.get("grids_ok") is not None:
                t = f"📊 {s['grids_ok']} 个网格成功"
                if s.get("grids_fail"):
                    t += f"，{s['grids_fail']} 个失败（详见处理详情）"
                    has_warn = True
                if s.get("files"):
                    t += f"，共 {s['files']} 份副本"
                lines.append(t)
                if s.get("encrypt", True):
                    lines.append("🔒 分发清单含明文密码，只留给自己用，勿随文件一起发出去")
        else:
            if s.get("groups") is not None:
                t = f"📊 拆出 {s['groups']} 个分组 · {s.get('files', '?')} 个文件 · {s.get('rows', '?')} 行"
                if s.get("zips"):
                    t += f" · {s['zips']} 个 ZIP"
                lines.append(t)
                bad = [f"{s[k]} 个{lab}" for k, lab in
                       (("skipped_sheets", "sheet 被跳过"), ("failed_files", "文件读取失败"),
                        ("failed_saves", "文件保存失败")) if s.get(k)]
                if bad:
                    lines.append("⚠ " + "，".join(bad) + "（点「处理详情」看原因）")
                    has_warn = True
                # v3.0：行数对账——源表每一行去哪了，一眼对得上
                if s.get("source_rows") and not s.get("stopped"):
                    parts = [f"拆出 {s.get('routed', 0)}"] + [
                        f"{lab} {s[k]}" for k, lab in (("skip_rows", "合计等"), ("not_selected", "没选的分组"))
                        if s.get(k)]
                    if s.get("unassigned"):
                        parts.append(f"没写分组 {s['unassigned']}（见「请核对_无归属.xlsx」）")
                        has_warn = True
                    lines.append(f"🧮 源表 {s['source_rows']} 行 = " + " + ".join(parts))
                if s.get("hidden_sheets") or s.get("hidden_cols"):
                    if self._keep_hidden_var.get():
                        lines.append("🙈 隐藏的 sheet / 列照样发出了（高级设置里勾选了「隐藏的也照样发出」）")
                    else:
                        hid = [f"{s[k]} 个{lab}" for k, lab in
                               (("hidden_sheets", "隐藏 sheet 没发"), ("hidden_cols", "隐藏列的内容已清空")) if s.get(k)]
                        lines.append("🙈 " + "，".join(hid) + "（要一起发，在高级设置里勾选）")
                if s.get("hidden_rows"):
                    lines.append(f"ℹ️ 有 {s['hidden_rows']} 行原来是隐藏的（可能是筛选掉的），照常发出了")
                if s.get("templates"):
                    t = f"📋 这批表其实是 {s['templates']} 种模板，跨文件合并已按列名对齐"
                    if s.get("unaligned"):
                        t += f"；{s['unaligned']} 个 sheet 列名大半对不上，单独放进了「sheet 名_2」"
                        has_warn = True
                    lines.append(t)
                if s.get("groups") == 1 and not s.get("stopped") and not self._split_list(self._selvals_var.get()):
                    lines.append("⚠ 只拆出 1 组：每一行的拆分字段都一样，可能选错了字段")
                    has_warn = True
                d = s.get("desens")
                if d:
                    if d.get("ok"):
                        lines.append(f"🛡 打码（{d['scene']}）：{d['columns']} 列、{d['cells']} 个格子；复检通过，"
                                     "脱敏报告在结果文件夹里")
                    else:
                        lines.append(f"❌ 复检没通过：还有 {d.get('residual', 0)} 个格子有完整的号码，没有打 ZIP；"
                                     "见结果里的「未通过复检_请勿发出.txt」，这批先别发")
                        has_warn = True
                if s.get("issued"):
                    lines.append("✍️ 已标出填写列并写了下发记录；收回来后切到「Excel 汇总」，选回收文件夹点「查收」"
                             "（下发记录已自动带上）")
        if self._warn_count:
            self._log_btn.configure(text=self._log_btn_text())
        lines.append(f"📁 {output_path}")
        self._summary_lbl.configure(text="\n".join(lines))
        # 有跳过/失败时横幅退回中性底色，避免「绿底＋橙字」的违和
        self._summary_banner.configure(
            fg_color=("gray90", "gray20") if has_warn else OK_BANNER_BG)
        self._summary_lbl.configure(text_color=C_WARN if has_warn else C_OK)
        self._show_summary_banner()

    def _show_summary_banner(self):
        """露出完成摘要并把滚动区拉到底：摘要在步骤下面，不拉的话要自己往下翻才看得到。"""
        self._summary_banner.grid()
        self.after(150, lambda: self._body._parent_canvas.yview_moveto(1.0))

    def _on_done(self, cfg, output_path):
        self._running = False
        is_pdf = cfg.get("ui_mode") == "pdf"
        self._start_btn.configure(state="normal", text=self._start_text())
        self._stop_btn.grid_remove()
        self._stop_indeterminate()
        self._progress.set(1 if output_path else 0)
        if cfg.get("ui_mode") == "collect":
            self._on_collect_done(cfg, output_path)
            if output_path and not self._stop_flag:
                self._offer_recipe(cfg)
            return
        if cfg.get("ui_mode") == "desens":
            self._on_desens_done(cfg, output_path)
            return
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
                self._offer_recipe(cfg)
            self._render_summary(output_path, is_pdf)
            from core.collect import RECORD_NAME
            rec = os.path.join(output_path, RECORD_NAME)
            if not is_pdf and cfg.get("fill_columns") and os.path.isfile(rec):
                # 刚下发的这一批：Excel 汇总直接打开查收、用它的下发记录，不用再去②里找。
                # 上一批手填的名单 / 必填列 / 标题属于上一批，清掉，否则会盖过这份记录
                self._col_rec_var.set(rec)
                self._col_check_var.set(True)
                for var, key in ((self._col_expect_var, "collect_expect"), (self._col_req_var, "collect_required"),
                                 (self._col_title_var, "collect_title")):
                    var.set("")
                    cfg[key] = ""
                self._col_form.grid()
                self._col_scanned = None
                cfg.update(collect_record=rec, collect_check=True)
                if cfg.get("due") and not self._col_due_var.get().strip():
                    self._col_due_var.set(cfg["due"])
                    cfg["collect_due"] = cfg["due"]
            try:
                save_config(cfg)      # 静默记住本次配置：下次打开即用（失败运行不存，避免存坏参数）
            except Exception:
                pass
        else:
            if self._warn_count:
                self._log_btn.configure(text=self._log_btn_text())
            self._run_status.configure(
                text="❌ 没有完成：原因见下方「处理详情」；可点「复制日志」连同反馈一起发给作者",
                text_color=C_ERR)
            self._copylog_btn.grid()
            self._toggle_log(show=True)
        if cfg.get("auto_open_output") and output_path and os.path.exists(output_path):
            # 只记日志、不改状态行：完成横幅已经显示了 📁 输出路径，
            # 绿色横幅旁边再挂个黄警告是噪音。
            if not open_folder(output_path):
                self._log(f"ℹ 没能自动打开输出文件夹，结果在：{output_path}")


def run():
    app = App()
    app.mainloop()
