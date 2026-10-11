"""
excelrouter skill · 汇总 / 查收前的「收文件」：找出文件夹里能合的表，把合不了的说清楚为什么。

Copyright (c) 2026 AbeLin
MIT License

收回来的表大多是从微信、企业微信、邮件里一份份另存下来的，文件夹里什么都有。
内核 core/collect.py 只认 .xlsx / .xls，遇到下面这些要么静默跳过、要么报一句看不懂的
「File is not a zip file」。这里先过一遍，**每个没合进去的文件都给出人话原因和办法**：

  · WPS 自己的 .et、CSV、.xlsb —— 内核读不了，请另存为 .xlsx
  · 很多业务系统「导出 Excel」导出的其实是网页（HTML）改了个 .xls 扩展名，Excel 能打开、程序读不了
  · 设了打开密码的 .xlsx（真实情况里不少）—— 文件头是 OLE 而不是 zip
  · 文件夹里混着上次拆分的结果（拆分结果 / 10111530结果）—— 合进来行数翻倍，先拦下
  · 拆分结果里同时有「汇总」和「到人」两份 —— 同一行出现两次，只合「汇总」那份
  · 同一份表交了几次（「城区一.xlsx」「城区一(1).xlsx」）—— 内容不同时都合进来，但点名提醒
  · .xlsm（带宏的表）—— 内核按值读得了，照常收进来

只做判断、不改任何文件。
"""

import os
import re

SUPPORTED = (".xlsx", ".xlsm", ".xls")
UNSUPPORTED = {
    ".et": "是 WPS 表格自己的格式（.et）",
    ".ett": "是 WPS 表格模板（.ett）",
    ".csv": "是 CSV 文本表格",
    ".xlsb": "是二进制工作簿（.xlsb）",
    ".numbers": "是苹果 Numbers 表格",
    ".ods": "是 OpenDocument 表格（.ods）",
}
_OLE = b"\xd0\xcf\x11\xe0"
_ZIP = b"PK"
_SPLIT_DIR = re.compile(r"^(拆分结果|.+_拆分结果|\d{8}结果(_\d+)?|分发结果)$")
_COPY_MARK = re.compile(r"(副本|复制|\(\d+\)|（\d+）|copy)", re.I)
RESULT_DIR = "汇总结果"
LIST_WORDS = ("应交名单", "发送清单", "分发清单")
OWN_OUTPUT = ("回收进度", "汇总_")


def _head(path, n=512):
    try:
        with open(path, "rb") as f:
            return f.read(n)
    except OSError:
        return b""


def sniff(path):
    """文件内容和扩展名对不上时返回原因（人话），对得上返回 ""。只看文件头几百字节。"""
    ext = os.path.splitext(path)[1].lower()
    head = _head(path)
    if not head:
        return "是空文件或读不了"
    if ext in (".xlsx", ".xlsm"):
        if head.startswith(_OLE):
            return "设了打开密码（或其实是老的 .xls 改了扩展名），请去掉密码、另存为 .xlsx 后再放回来"
        if not head.startswith(_ZIP):
            return "不是有效的 Excel 文件（可能下载不完整或已损坏），请重新另存一份"
    elif ext == ".xls":
        low = head.lstrip(b"\xef\xbb\xbf \r\n\t").lower()
        if low.startswith((b"<", b"mime-version")) or b"<html" in low or b"<table" in low:
            return "其实是网页格式（很多系统「导出 Excel」导出的就是这种），请用 Excel 或 WPS 打开后另存为 .xlsx"
        if head.startswith(_ZIP):
            return "其实是 .xlsx 格式却叫 .xls，把扩展名改成 .xlsx 即可"
        if not head.startswith(_OLE):
            return "不是有效的 Excel 文件，请用 Excel 或 WPS 打开后另存为 .xlsx"
    return ""


def order_key(rel):
    """同一文件夹里原件排在「副本」「(1)」前面：表头以原件为准，内容相同时留下的也是原件。"""
    d, n = os.path.split(rel)
    stem = os.path.splitext(n)[0]
    return d, _COPY_MARK.sub("", stem).strip(" -_"), bool(_COPY_MARK.search(stem)), n


def gather(input_path, out_dir, base_stem=None):
    """返回 dict：
      files     [(绝对路径, 相对路径)] 能交给内核的表（名单表不在内）
      lists     名单表（文件名含「应交名单」「发送清单」「分发清单」）的绝对路径
      problems  [人话] 没收进来的文件和原因
      warnings  [人话] 收进来了但要提醒的
      bad       [(绝对路径, 相对路径)] problems 里那些文件（查收时它们也算「交了」，只是读不了）
      stale     输入里混着的历史拆分结果目录（调用方决定拦不拦）
    base_stem：去掉「副本」「(2)」后缀的函数（core.collect._base_stem），用来点名重复交的表。
    """
    out = {"files": [], "lists": [], "problems": [], "warnings": [], "stale": [], "bad": []}
    src = os.path.abspath(input_path)
    if os.path.isfile(src):
        ext = os.path.splitext(src)[1].lower()
        name = os.path.basename(src)
        if ext in UNSUPPORTED:
            out["problems"].append(f"「{name}」{UNSUPPORTED[ext]}，读不了，请用 Excel 或 WPS 另存为 .xlsx")
            out["bad"].append((src, name))
        elif ext not in SUPPORTED:
            out["problems"].append(f"「{name}」不是 Excel 表格")
            out["bad"].append((src, name))     # 与 problems 一一对应
        elif sniff(src):
            out["problems"].append(f"「{name}」{sniff(src)}")
            out["bad"].append((src, name))
        else:
            out["files"].append((src, name))
        _lock_warning(os.path.dirname(src), [name], out)
        return out

    out_abs = os.path.normcase(os.path.abspath(out_dir)) if out_dir else ""
    loose, in_split, person_skipped, locked = [], [], 0, []
    for root, dirs, names in os.walk(src):
        keep = []
        for d in sorted(dirs):
            full = os.path.normcase(os.path.abspath(os.path.join(root, d)))
            if d == RESULT_DIR or (out_abs and (full == out_abs or full.startswith(out_abs + os.sep))):
                continue
            if d == "到人" and "汇总" in dirs:
                person_skipped += sum(1 for _r, _d, ns in os.walk(os.path.join(root, d))
                                      for n in ns if n.lower().endswith(SUPPORTED))
                continue
            keep.append(d)
        dirs[:] = keep
        rel_root = os.path.relpath(root, src)
        inside_split = any(_SPLIT_DIR.match(p) for p in rel_root.split(os.sep) if p != ".")
        for d in dirs:
            if _SPLIT_DIR.match(d) and not inside_split:
                out["stale"].append(os.path.join(root, d))
        for n in sorted(names):
            low = n.lower()
            ext = os.path.splitext(low)[1]
            path = os.path.join(root, n)
            rel = os.path.relpath(path, src)
            if n.startswith("~$"):
                locked.append(n[2:])
                continue
            if low.endswith("__tmp__.xlsx"):
                continue
            if ext in UNSUPPORTED:
                out["problems"].append(f"「{rel}」{UNSUPPORTED[ext]}，读不了，请用 Excel 或 WPS 另存为 .xlsx 后放回来")
                out["bad"].append((path, rel))
                continue
            if ext not in SUPPORTED:
                continue
            stem = os.path.splitext(n)[0]
            if stem.startswith(OWN_OUTPUT):
                continue
            why = sniff(path)
            if why:
                out["problems"].append(f"「{rel}」{why}")
                out["bad"].append((path, rel))
                continue
            if any(w in stem for w in LIST_WORDS):
                out["lists"].append(path)
                continue
            out["files"].append((path, rel))
            (in_split if inside_split else loose).append(rel)
    out["files"].sort(key=lambda f: order_key(f[1]))
    # 只有「拆分结果」本身被选来合（里面没有别的表）时不算混进来
    if not loose:
        out["stale"] = []
    if person_skipped:
        out["warnings"].append(f"拆分结果里「到人」文件夹的 {person_skipped} 个表和「汇总」里是同样的行，"
                               f"只合了「汇总」那份，免得每行出现两次")
    if base_stem:
        seen = {}
        for _p, rel in out["files"]:
            seen.setdefault((os.path.dirname(rel), base_stem(rel)), []).append(rel)
        multi = [v for v in seen.values() if len(v) > 1]
        if multi:
            shown = "；".join("、".join(f"「{r}」" for r in v) for v in multi[:3])
            out["warnings"].append(f"{len(multi)} 组文件像是同一份表交了几次：{shown}"
                                   + (" 等" if len(multi) > 3 else "")
                                   + "。内容完全一样的只合一份，不一样的都合进来了，请核对要不要删掉旧的那份")
    _lock_warning(src, locked, out, already=True)
    return out


def _lock_warning(folder, names, out, already=False):
    """Excel / WPS 打开着文件时会留下 ~$ 开头的锁文件：读到的是上次保存的内容。"""
    locked = names if already else [n for n in names if os.path.exists(os.path.join(folder, "~$" + n))
                                    or os.path.exists(os.path.join(folder, "~$" + n[2:]))]
    if locked:
        out["warnings"].append(f"{'、'.join(locked[:3])}{' 等' if len(locked) > 3 else ''}正被 Excel/WPS 打开着，"
                               f"读到的是上次保存的内容；有没保存的修改请先保存")
