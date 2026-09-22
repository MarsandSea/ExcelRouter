"""
ExcelRouter - 桌面平台兼容层
Copyright (c) 2026 AbeLin
MIT License

把「桌面外壳」形状的平台差异集中到这一个模块：打开文件管理器、解析拖放载荷、
挑中文 UI 字体、设窗口图标、认锁文件。Windows / Linux（银河麒麟 V10、统信 UOS）
两条路都走这里，其余代码不再散落 sys.platform 判断。

★ 本模块不得在模块级 import tkinter / customtkinter ——
  单元测试要能在没有显示器的环境（CI 容器、麒麟 ssh）里直接跑。
  需要 Tk 对象的函数一律把它当参数传进来。
"""

import os
import re
import sys
import shutil
import subprocess
import urllib.parse

IS_WINDOWS = sys.platform.startswith("win")
IS_MAC = sys.platform == "darwin"
IS_LINUX = not IS_WINDOWS and not IS_MAC


# =====================================================
# 打开输出文件夹
# =====================================================

# Linux 文件管理器回退链。麒麟 V10 桌面（UKUI）用的是 peony；
# 精简过的信创镜像可能连 xdg-utils 都没装，所以具体的文件管理器也逐个试。
_LINUX_OPENERS = (
    ["xdg-open"], ["gio", "open"], ["peony"], ["nautilus"],
    ["caja"], ["dolphin"], ["thunar"], ["nemo"],
)


def open_folder(path):
    """用系统文件管理器打开目录。成功返回 True，失败返回 False。

    ★ 调用方必须处理 False，给用户一句可见的提示 —— 旧实现失败时悄无声息，
      用户点了「打开输出文件夹」什么也没发生，比报错更让人困惑。
    """
    if not path or not os.path.exists(path):
        return False
    path = os.path.abspath(path)
    try:
        if IS_WINDOWS:
            # 必须传 list。旧实现是 Popen(f'explorer "{p}"') 这种裸字符串，
            # 在 POSIX 上会被当成「一个叫这么长名字的可执行文件」→ FileNotFoundError。
            # explorer 不认正斜杠，所以这一支保留 normpath。
            subprocess.Popen(["explorer", os.path.normpath(path)])
            return True
        if IS_MAC:
            subprocess.Popen(["open", path])
            return True
        for cmd in _LINUX_OPENERS:
            if shutil.which(cmd[0]):
                # start_new_session：不然文件管理器继承本进程的进程组，
                # ExcelRouter 一退出就把它一起带走了。
                subprocess.Popen(cmd + [path],
                                 stdout=subprocess.DEVNULL,
                                 stderr=subprocess.DEVNULL,
                                 start_new_session=True)
                return True
    except Exception:
        return False
    return False


# =====================================================
# 拖放载荷解析
# =====================================================

def _from_uri(token, *, windows=IS_WINDOWS):
    """把拖放载荷里的一个 token 还原成本地路径。

    Windows 的 tkdnd 直接发路径；X11/XDND 发的是百分号编码的 file:// URI
    （`file:///home/u/%E8%A1%A8.xlsx`），不解码就过不了下游的扩展名/isdir 检查。
    """
    t = (token or "").strip().strip("\r\n")
    if not t:
        return ""
    if t.startswith("file://"):
        parts = urllib.parse.urlsplit(t)
        path = urllib.parse.unquote(parts.path)
        host = (parts.netloc or "").lower()
        if host and host != "localhost":
            path = "//" + parts.netloc + path        # file://server/share → UNC
        if windows and re.match(r"^/[A-Za-z]:", path):
            path = path[1:]                          # file:///C:/x → C:/x
        return path
    if t.startswith("file:"):
        return urllib.parse.unquote(t[5:])           # 少数实现发单斜杠 file:/path
    return t


def parse_drop_paths(data, *, windows=IS_WINDOWS):
    """解析 tkdnd <<Drop>> 事件的数据，返回本地路径列表。

    两个平台的 tkdnd 都发 Tcl list（`{带空格的路径} 普通路径`），分词逻辑通用，
    `\\S+` 也已经能切开换行分隔的多文件。差异只在 X11 那边元素是 file:// URI。

    windows 关键字参数是为了可测性：能在 Linux 上测盘符分支、在 Windows 上测
    POSIX 分支，不用 monkeypatch 全局状态。
    """
    if not data:
        return []
    out = []
    for m in re.finditer(r"\{([^}]*)\}|(\S+)", data):
        raw = m.group(1) if m.group(1) is not None else m.group(2)
        p = _from_uri(raw, windows=windows)
        if p:
            out.append(p)
    return out


# =====================================================
# 中文 UI 字体
# =====================================================

# 按优先级排列。雅黑排最前，保证 Windows 行为与 v2.7.2 完全一致。
# tkfont.families() 在部分 Linux 上返回中文别名，所以中英文名都列。
_UI_FONT_PREFS = (
    "Microsoft YaHei UI", "Microsoft YaHei",              # Windows
    "Noto Sans CJK SC", "Noto Sans SC",                   # 麒麟 / UOS 默认
    "Source Han Sans SC", "Source Han Sans CN", "思源黑体",
    "WenQuanYi Zen Hei", "WenQuanYi Micro Hei",           # 文泉驿
    "文泉驿正黑", "文泉驿微米黑",
    "FZLanTingHei-R-GBK", "方正兰亭黑_GBK", "方正黑体_GBK",  # 麒麟预装方正字库
)


def pick_ui_font(families):
    """从已安装字体族里挑一款中文 UI 字体；一个都没有返回 None（交给 tk 自己回退）。"""
    fams = set(families or ())
    for name in _UI_FONT_PREFS:
        if name in fams:
            return name
    return None


# =====================================================
# 窗口图标
# =====================================================

def set_window_icon(win, ico_paths, png_paths):
    """设置窗口图标：先试 .ico/iconbitmap（Windows），再试 .png/iconphoto（X11）。

    iconbitmap() 在 X11 上只认 XBM 位图，传 .ico 必抛 TclError —— 所以麒麟上
    标题栏/任务栏会是 Tk 默认羽毛，必须用 iconphoto + PNG 兜底。
    """
    for ico in ico_paths:
        if os.path.exists(ico):
            try:
                win.iconbitmap(ico)
                return True
            except Exception:
                pass
    for png in png_paths:
        if os.path.exists(png):
            try:
                import tkinter as tk       # 局部导入：保持本模块可无显示器单测
                img = tk.PhotoImage(file=png)
                # ★ 必须持引用：PhotoImage 在 Python 侧被 GC 后，Tk 仍按名字引用它，
                #   图标会凭空消失。这是 Tk 最常见的坑。
                win._er_icon_ref = img
                win.iconphoto(True, img)
                return True
            except Exception:
                pass
    return False


# =====================================================
# 锁文件
# =====================================================

def lock_owner_name(filename):
    """锁文件 → 它锁住的原文件名；不是锁文件返回 None。

    ~$工资表.xlsx        → 工资表.xlsx      （Microsoft Excel / Windows 版 WPS）
    .~lock.工资表.xlsx#  → 工资表.xlsx      （LibreOffice / Linux 版 WPS）

    麒麟上用的是后一种格式，只认 `~$` 的话「文件正被打开」预警永远不会触发。
    """
    if not filename:
        return None
    if filename.startswith("~$"):
        return filename[2:] or None
    if filename.startswith(".~lock.") and filename.endswith("#"):
        return filename[len(".~lock."):-1] or None
    return None
