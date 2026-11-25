"""两表核对的界面文案：关键列下拉的选项、结果卡片里的几行字。

Copyright (c) 2026 AbeLin · MIT License

和 platform_compat 一样**不 import tkinter / customtkinter**，单测可以无头跑。
输入是 core.compare.scan_pair 的返回值和 run_compare 收尾的 [SUMMARY]（JSON 解出来的 dict）。
"""

KEY_SEP = " + "        # 两列合起来当关键列时，下拉里这样写
SAMPLE_SEP = "、"
MORE = "……"


def key_label(cols):
    return KEY_SEP.join(cols)


def key_options(info, saved=None):
    """关键列下拉：(选项文字列表, {选项文字: [列名]}, 默认选中的文字)。

    推荐的排前面（写明能对上几行），其余两表都有的列跟在后面；saved 是上次用的关键列（列名列表），
    还在就默认选它，否则选第一个推荐的；一个推荐都没有时默认空串（让用户自己挑）。
    """
    labels, mapping = [], {}
    for i, k in enumerate(info.get("keys") or []):
        name = key_label(k["cols"])
        tag = "推荐" if i == 0 else "也可以"
        text = f"{name}（{tag}，能对上 {k['matched']} 行）"
        labels.append(text)
        mapping[text] = list(k["cols"])
    taken = {tuple(v) for v in mapping.values()}
    for c in info.get("columns") or []:
        if (c,) not in taken:
            labels.append(c)
            mapping[c] = [c]
    default = ""
    if saved:
        default = next((t for t, cols in mapping.items() if cols == list(saved)), "")
    if not default and info.get("keys"):
        default = labels[0]
    return labels, mapping, default


def _samples(items, total):
    if not items:
        return ""
    s = SAMPLE_SEP.join(str(x) for x in items)
    return f"：{s}{MORE if total > len(items) else ''}"


def result_lines(s):
    """结果卡片的几行：[(文字, 种类)]，种类是 "add" / "del" / "chg" / "warn" / "info"。"""
    pv = s.get("preview") or {}
    out = []
    n = s.get("only_b", 0)
    out.append((f"➕ 新增 {n} 行（新表有、原表没有）" + _samples(pv.get("only_b"), n), "add" if n else "info"))
    n = s.get("only_a", 0)
    out.append((f"➖ 没了 {n} 行（原表有、新表没有）" + _samples(pv.get("only_a"), n), "del" if n else "info"))
    n = s.get("changed_rows", 0)
    chg = f"✏️ 改过 {n} 行、{s.get('changed_cells', 0)} 个格子"
    if s.get("per_col"):
        chg += "；改得最多的列：" + SAMPLE_SEP.join(f"{k} {v} 行" for k, v in s["per_col"].items())
    out.append((chg, "chg" if n else "info"))
    shown = pv.get("changed") or []
    for key, cols in shown:
        out.append((f"    · {key}：改了 {SAMPLE_SEP.join(cols)}", "chg"))
    if n > len(shown):
        out.append((f"    {MORE}其余见核对结果里的「有变化」「变化行」", "info"))
    if s.get("unpaired"):
        out.append((f"⚠ {s['unpaired']} 行关键列重复或空着，没法配对，单独列在「关键列重复」里", "warn"))
    for side, key in (("原表", "only_a_cols"), ("新表", "only_b_cols")):
        cols = s.get(key) or []
        if cols:
            more = " 等" if len(cols) > 6 else ""
            out.append((f"ℹ 只有{side}有的列不比较：{SAMPLE_SEP.join(cols[:6])}{more}", "info"))
    return out


def banner_lines(s, output):
    """③ 下面完成摘要横幅：(文字, 有没有要留意的)。"""
    same = not (s.get("only_a") or s.get("only_b") or s.get("changed_rows"))
    head = (f"🔍 按「{key_label(s.get('key') or [])}」对上 {s.get('matched', 0)} 行 · "
            f"新增 {s.get('only_b', 0)} · 没了 {s.get('only_a', 0)} · 改过 {s.get('changed_rows', 0)} 行")
    lines = [head]
    if same and not s.get("unpaired"):
        lines.append("✅ 两张表的内容一样（按比较口径：100 和 100.00、全角半角、首尾空格都不算不同）")
    if s.get("unpaired"):
        lines.append(f"⚠ {s['unpaired']} 行关键列重复或空着，没有配对")
    lines.append(f"📁 {output}")
    return "\n".join(lines), bool(s.get("unpaired"))
