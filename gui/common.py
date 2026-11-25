"""
ExcelRouter - 图形界面公用部分：路径、文案常量、配色、按钮工厂、配置读写。

四个模式（gui/mode_*.py）和主窗口（gui/app.py）共用；本模块不建任何控件。

Copyright (c) 2026 AbeLin
MIT License
"""

import os
import sys
import json

import customtkinter as ctk

from gui.platform_compat import pick_ui_font


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
# 速度模式两档（v2.9）：文案里直接写清代价，用户不用点开帮助就能选
SPEED_STD  = "标准（保留格式）"
SPEED_FAST = "极速（纯数据）"
PDF_COL_PLACEHOLDER = "（选清单后自动识别）"
# 顶部模式切换：内部模式名 ↔ 按钮文字（ui_mode 存的是内部名；字典顺序 = 按钮顺序）。
# 「collect」对外叫「Excel 汇总」：默认是普通的合并汇总，查收（谁没交、谁漏填）是 ② 里的可选开关
# 「pdf」对外叫「文件加密分发」（v2.10 起也收 Word / Excel / PPT，内部名不改，ui_mode 已存进用户配置）
# 「desens」对外叫「数据脱敏」（v3.1）：单独给表格打码 / 换代号 / 清空，处理完复检
MODE_LABELS = {"excel": "Excel 拆分", "collect": "Excel 汇总", "desens": "数据脱敏", "pdf": "文件加密分发"}
COL_IN_TIP = "把要合并的表放进一个文件夹再选它；选一个文件时，把它的各个 sheet 合成一张"
DS_IN_TIP = "选一个表或一个文件夹；结果放在旁边的「脱敏结果」里，原表不改"
# 数据脱敏的用途与处理方式（与 core/desens.py 一致；GUI 启动时不 import core）
DS_SCENES = ("发群通报", "内部下发", "交给AI分析", "对外报送")
DS_SCENE_TEXT = {
    "发群通报": "认出来的个人信息全部打码：138****5678、张*、证件号留前 6 后 4、地址留到镇街",
    "内部下发": "只遮证件号和银行卡号，手机号、姓名、地址照发（客户经理外呼要完整号码）",
    "交给AI分析": "姓名、号码、证件、编码换成一致的代号（还能统计、关联，认不出人），地址到区县、日期到年",
    "对外报送": "个人信息整列清空，编码换成代号",
}
DS_METHODS = {"mask": "打码", "pseudo": "换代号", "generalize": "泛化", "clear": "清空", "keep": "照发"}
DS_METHOD_WORDS = {"打码": "mask", "换代号": "pseudo", "代号": "pseudo", "泛化": "generalize",
                   "清空": "clear", "照发": "keep", "不处理": "keep"}
EXPECT_TIP = "名字用顿号或逗号隔开，也可以点「从表格读」；留空时用下发记录，或文件夹里名为「应交名单」的表"
# 文件加密分发能收的文件（与 core.pdf_dist.SUPPORTED_EXTS 一致；GUI 不在启动时 import core）
DIST_EXTS = (".pdf", ".docx", ".xlsx", ".pptx", ".docm", ".xlsm", ".pptm")

# UI 泵用的「本轮没有此类消息」哨兵（不能用 None：扫描失败时 payload 可能为空）
_MISSING = object()

# 状态文字颜色（浅色模式, 深色模式）
C_OK    = ("#15803d", "#4ade80")
C_WARN  = ("#b45309", "#fbbf24")
C_ERR   = ("#b91c1c", "#f87171")
C_MUTED = ("gray40", "gray60")

# ── 品牌调色板（v2.7.1，浅色/深色双主题集中取值）──────────────
# 主色「移动蓝」：使用者是移动系办公人群，品牌联想自然；比 ctk 默认 blue
# (#3B8ED0) 更饱和明亮，主按钮/徽章/进度条统一从这套取色，形成记忆点。
# 完成态用明确的浅绿横幅承载——「显眼」最重要的时刻是完成瞬间。
PRIMARY        = ("#0E7FD1", "#38BDF8")   # 主按钮、步骤徽章、进度条、分段选择高亮
PRIMARY_HOVER  = ("#0B6BB0", "#7DD3FC")
OK_BANNER_BG   = ("#DCFCE7", "#14532D")   # 完成摘要横幅底色（文字仍用 C_OK）
CARD_BORDER    = ("gray82", "gray28")     # 步骤卡片 1px 淡边框
# ACCENT 保留为 PRIMARY 的别名：徽章等既有引用不逐处改动，取色来源已经换血
ACCENT  = PRIMARY

def _init_fonts():
    """把 CTk 默认字体族设为一款中文字体（tk 默认 Roboto 回退的中文渲染很差）。

    CTkFont 未显式传 family 时取 ThemeManager.theme["CTkFont"]["family"]，这是
    唯一真正的全局默认入口（实测改 FontManager._default_font 无效）。主题 dict 是
    set_default_color_theme 时载入的运行时副本，改它对已建/将建组件一致生效。

    ★ 调用时机：必须在 root 建立后（tkfont.families() 需要 root）、任何组件创建前
      （组件的 CTkFont 在构造时读主题值，晚了改无效）。

    候选顺序见 platform_compat._UI_FONT_PREFS：雅黑排第一，Windows 行为不变；
    麒麟 / UOS 上落到 Noto Sans CJK SC、文泉驿或方正字库。一个都没有就不动主题，
    交给 tk 自己的字体回退。**不要改 ["CTkFont"]["size"]** —— 页脚已有「大字号」
    开关，动默认字号会静默重排每个老用户的窗口。
    """
    try:
        import tkinter.font as tkfont
        fam = pick_ui_font(tkfont.families())
        if fam:
            ctk.ThemeManager.theme["CTkFont"]["family"] = fam
    except Exception:
        pass

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
    "renumber_seq": False,         # v3.2 通报表：拆完「序号」每组从 1 重新编号
    "group_totals": False,         # v3.2 通报表：每组重算合计行
    "keep_hidden": False,          # v3.0：隐藏的 sheet 和列也照样发出（默认不发 sheet、清空列）
    "auto_open_output": True,
    # ── 文件加密分发（v2.6 起 PDF，v2.10 起也收 Word/Excel/PPT；键名沿用 pdf_*）──
    "ui_mode": "excel",
    "pdf_input_paths": [],
    "pdf_mapping_path": "",
    "pdf_grid_column": "",
    "pdf_password_column": "",
    "pdf_receiver_column": "",
    "pdf_encrypt": True,
    "pdf_watermark": True,
    "pdf_watermark_text": "{grid} {date}",
    "pdf_watermark_opacity": 0.15,
    "pdf_watermark_angle": 45,
    # ── Excel 汇总（v2.10）：名单 / 必填列 / 标题等都是文字，core.collect.split_names 负责拆 ──
    "collect_input_path": "",
    "collect_check": False,    # ② 的开关：True = 先查收（谁没交、谁漏填）再汇总；False = 普通汇总
    "collect_expect": "",
    "collect_required": "",
    "collect_by": "",
    "collect_title": "",
    "collect_due": "",
    "collect_record": "",
    "fill_columns": "",       # 拆分时让对方填写的列（选填）：标黄、其余列锁定、写下发记录
    "due": "",
    # ── 方案（v3.0）：最近打开 / 存过的方案文件路径，跑完后「存成方案」的提示点过不再提示 ──
    "recent_recipes": [],
    "recipe_tip_off": False,
    # ── 界面个性化（v2.7）──
    "window_geometry": "",     # 上次关闭时的窗口大小位置；空 = 按屏幕自适应
    "ui_scale": 100,           # 界面缩放百分比（大字号=115），重启后生效
    # ── 界面个性化（v2.7.1）──
    "appearance_mode": "system",   # system / light / dark，页脚按钮循环切换
}


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
