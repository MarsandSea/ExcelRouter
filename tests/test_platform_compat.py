"""
gui/platform_compat.py 单元测试
Copyright (c) 2026 AbeLin
MIT License

★ 本文件只 import gui.platform_compat，不碰 tkinter/customtkinter ——
  必须能在没有显示器的环境（CI 容器、麒麟 ssh）里直接 pytest。
"""

import pytest

from gui import platform_compat as pc


# =====================================================
# parse_drop_paths —— 最高价值：纯函数，且旧实现在 X11 上拒绝一切拖放
# =====================================================

def test_drop_windows_braced_and_plain():
    """Windows tkdnd：`{带空格的路径} 普通路径` 混合。"""
    got = pc.parse_drop_paths(r"{C:\a b\x.xlsx} C:\y.xlsx", windows=True)
    assert got == [r"C:\a b\x.xlsx", r"C:\y.xlsx"]


def test_drop_x11_percent_encoded_uri():
    """X11/XDND：百分号编码的 file:// URI —— 不解码就过不了下游扩展名检查。"""
    got = pc.parse_drop_paths("file:///home/u/%E8%A1%A8%20%E6%A0%BC.xlsx", windows=False)
    assert got == ["/home/u/表 格.xlsx"]


def test_drop_newline_separated_multiple():
    """多文件拖放常以换行分隔。"""
    got = pc.parse_drop_paths("file:///a/1.pdf\nfile:///a/2.pdf", windows=False)
    assert got == ["/a/1.pdf", "/a/2.pdf"]


def test_drop_braced_uri():
    """带空格的 URI 会被花括号包起来，两层都要剥。"""
    got = pc.parse_drop_paths("{file:///a/b%20c.xlsx}", windows=False)
    assert got == ["/a/b c.xlsx"]


def test_drop_bare_posix_path_unchanged():
    got = pc.parse_drop_paths("/home/u/x.xlsx", windows=False)
    assert got == ["/home/u/x.xlsx"]


def test_drop_localhost_authority_stripped():
    got = pc.parse_drop_paths("file://localhost/a/b.xlsx", windows=False)
    assert got == ["/a/b.xlsx"]


def test_drop_unc_authority_kept():
    """file://server/share → UNC 路径，authority 不能丢。"""
    got = pc.parse_drop_paths("file://server/share/x.xlsx", windows=True)
    assert got == ["//server/share/x.xlsx"]


def test_drop_drive_letter_depends_on_platform():
    """同一份载荷在两个平台上的正确结果不同 —— 这正是 windows= 参数存在的理由。"""
    uri = "file:///C:/a/b.xlsx"
    assert pc.parse_drop_paths(uri, windows=True) == ["C:/a/b.xlsx"]
    assert pc.parse_drop_paths(uri, windows=False) == ["/C:/a/b.xlsx"]


def test_drop_single_slash_file_scheme():
    got = pc.parse_drop_paths("file:/a/b%20c.xlsx", windows=False)
    assert got == ["/a/b c.xlsx"]


@pytest.mark.parametrize("data", ["", None])
def test_drop_empty(data):
    assert pc.parse_drop_paths(data) == []


# =====================================================
# pick_ui_font
# =====================================================

def test_font_yahei_wins_when_present():
    """Windows 行为不能变：雅黑在场时永远第一。"""
    assert pc.pick_ui_font(["Noto Sans CJK SC", "Microsoft YaHei UI"]) == "Microsoft YaHei UI"


def test_font_noto_wins_without_yahei():
    assert pc.pick_ui_font(["DejaVu Sans", "Noto Sans CJK SC"]) == "Noto Sans CJK SC"


def test_font_chinese_alias_matches():
    """tkfont.families() 在部分 Linux 上返回中文别名。"""
    assert pc.pick_ui_font(["文泉驿正黑"]) == "文泉驿正黑"


def test_font_none_when_no_cjk():
    assert pc.pick_ui_font(["Roboto", "DejaVu Sans"]) is None
    assert pc.pick_ui_font([]) is None
    assert pc.pick_ui_font(None) is None


# =====================================================
# lock_owner_name
# =====================================================

@pytest.mark.parametrize("name, expect", [
    ("~$工资表.xlsx", "工资表.xlsx"),              # Excel / Windows 版 WPS
    (".~lock.工资表.xlsx#", "工资表.xlsx"),        # LibreOffice / Linux 版 WPS
    ("工资表.xlsx", None),
    (".~lock.工资表.xlsx", None),                  # 少了结尾 #，不是锁文件
    ("~$", None),
    ("", None),
])
def test_lock_owner_name(name, expect):
    assert pc.lock_owner_name(name) == expect


# =====================================================
# open_folder
# =====================================================

def test_open_folder_rejects_missing_path(tmp_path):
    assert pc.open_folder(str(tmp_path / "并不存在")) is False
    assert pc.open_folder("") is False
    assert pc.open_folder(None) is False


def test_open_folder_passes_a_list_not_a_string(tmp_path, monkeypatch):
    """★ 这条正是本次修的 bug：旧实现传的是裸字符串 f'explorer \"{p}\"'，
    在 POSIX 上会被当成「一个叫这么长名字的可执行文件」→ FileNotFoundError。"""
    seen = {}

    def fake_popen(cmd, **kw):
        seen["cmd"] = cmd
        seen["kw"] = kw
        return object()

    monkeypatch.setattr(pc.subprocess, "Popen", fake_popen)
    monkeypatch.setattr(pc.shutil, "which", lambda c: "/usr/bin/" + c)
    assert pc.open_folder(str(tmp_path)) is True
    assert isinstance(seen["cmd"], list)
    assert str(tmp_path) in seen["cmd"][-1]


@pytest.mark.skipif(pc.IS_WINDOWS or pc.IS_MAC, reason="仅 Linux 走文件管理器回退链")
def test_open_folder_linux_fallback_order(tmp_path, monkeypatch):
    """xdg-open 缺席时应依次退到 gio → peony（麒麟 UKUI 的文件管理器）。"""
    calls = []
    monkeypatch.setattr(pc.subprocess, "Popen",
                        lambda cmd, **kw: calls.append((cmd, kw)) or object())
    # 只有 peony 存在：精简过的信创镜像常常没装 xdg-utils
    monkeypatch.setattr(pc.shutil, "which",
                        lambda c: "/usr/bin/peony" if c == "peony" else None)
    assert pc.open_folder(str(tmp_path)) is True
    cmd, kw = calls[0]
    assert cmd[0] == "peony"
    # 不设 start_new_session 的话，ExcelRouter 一退出会把文件管理器一起带走
    assert kw.get("start_new_session") is True


@pytest.mark.skipif(pc.IS_WINDOWS or pc.IS_MAC, reason="仅 Linux 走文件管理器回退链")
def test_open_folder_false_when_no_opener(tmp_path, monkeypatch):
    monkeypatch.setattr(pc.shutil, "which", lambda c: None)
    assert pc.open_folder(str(tmp_path)) is False


def test_open_folder_survives_popen_failure(tmp_path, monkeypatch):
    """文件管理器起不来也不能抛到调用方——GUI 那边只认 True/False。"""
    monkeypatch.setattr(pc.shutil, "which", lambda c: "/usr/bin/" + c)

    def boom(*a, **kw):
        raise OSError("no such thing")

    monkeypatch.setattr(pc.subprocess, "Popen", boom)
    assert pc.open_folder(str(tmp_path)) is False


def test_platform_flags_are_mutually_exclusive():
    assert sum([pc.IS_WINDOWS, pc.IS_MAC, pc.IS_LINUX]) == 1


def test_module_does_not_import_tkinter():
    """★ 保住「pytest 可无头运行」这条性质：本模块不许在模块级拉进 tkinter。"""
    import ast
    src = open(pc.__file__, encoding="utf-8").read()
    tree = ast.parse(src)
    top = []
    for node in tree.body:          # 只看模块级，函数体内的局部 import 是允许的
        if isinstance(node, ast.Import):
            top += [a.name for a in node.names]
        elif isinstance(node, ast.ImportFrom):
            top.append(node.module or "")
    assert not [m for m in top if m.split(".")[0] in ("tkinter", "customtkinter")]
