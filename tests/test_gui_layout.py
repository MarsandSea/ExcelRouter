"""
界面代码按模式拆成几个 mixin 后的结构守卫（只读源码做 AST 检查，不 import tkinter，可无头跑）。

App 由 gui/app.py 和 gui/mode_*.py 的几个 mixin 拼成：同名方法写在两处时，
排在 MRO 前面的那个会静默盖掉另一个，界面行为跟着变却不报错——这里守着这条。

Copyright (c) 2026 AbeLin
MIT License
"""

import ast
import os
import re

GUI = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "gui")
MODES = ("mode_split", "mode_collect", "mode_desens", "mode_pdf")


def _methods(fname, cls_suffix):
    tree = ast.parse(open(os.path.join(GUI, fname + ".py"), encoding="utf-8").read())
    for n in tree.body:
        if isinstance(n, ast.ClassDef) and n.name.endswith(cls_suffix):
            return {m.name for m in n.body if isinstance(m, ast.FunctionDef)}
    raise AssertionError(f"{fname}.py 里找不到 *{cls_suffix} 类")


def test_no_method_defined_twice():
    seen = {"app": _methods("app", "App")}
    for f in MODES:
        seen[f] = _methods(f, "ModeMixin")
    names = list(seen)
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            both = seen[a] & seen[b]
            assert not both, f"{a} 和 {b} 都定义了 {sorted(both)}"


def test_app_version_stays_in_app_py():
    # build_linux.sh 用正则从 gui/app.py 读 APP_VERSION，别把它挪进 common.py
    src = open(os.path.join(GUI, "app.py"), encoding="utf-8").read()
    assert re.search(r'APP_VERSION\s*=\s*"([^"]+)"', src)


def test_mode_modules_do_not_import_app():
    # mixin 反过来 import gui.app 会成环；也不该在模块级 import core（GUI 启动不加载 core）
    for f in MODES + ("common",):
        tree = ast.parse(open(os.path.join(GUI, f + ".py"), encoding="utf-8").read())
        for n in tree.body:
            if isinstance(n, ast.ImportFrom):
                assert n.module not in ("gui.app",), f
                assert not (n.module or "").startswith("core"), f
            if isinstance(n, ast.Import):
                assert not any(a.name.startswith("core") for a in n.names), f
