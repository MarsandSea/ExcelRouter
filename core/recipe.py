"""方案：把一次设置存成一个小 JSON 文件，下个月一键套用、发给同事也能用。

Copyright (c) 2026 AbeLin · MIT License

方案只存「怎么做」（拆分字段、填写列、应交名单……），不存「对哪些文件做」：路径、截止时间一律不存。
方案文件会在同事之间传，按外来数据处理：只认白名单里的键，类型以默认值为准，文字长度有上限，
绝不执行、绝不当路径打开。桌面版和技能共用这一份格式。
"""
import datetime
import json
import os
import unicodedata

from core.utils import soft_clean

RECIPE_VERSION = 1
RECIPE_KEY = "excelrouter_recipe"
RECIPE_DIR = "我的方案"
RECENT_MAX = 5
MAX_TEXT = 5000          # 一个文字设置最多这么多字（应交名单、填写列可能较长）
MAX_ITEMS = 1000         # 一个列表设置最多这么多项
MAX_ITEM = 200           # 列表里每项最多这么多字
MAX_NAME, MAX_NOTE = 60, 500

# 每个设置的默认值：类型以它为准（和 config/default_config.json、界面的 FALLBACK 一致）
DEFAULTS = {
    "split_column": "", "selected_values": [], "header_mode": "auto", "header_row": 1, "grid_keys": [],
    "id_keys": [], "to_person": False, "person_column": "", "person_file_filter": [], "make_zip": True,
    "merge_across_files": False, "preserve_format": True, "keep_formulas": False, "keep_hidden": False,
    "exact_match": True, "value_alias_map": {}, "skip_values": ["合计", "小计", "总计", "平均", ""],
    "fill_columns": "",
    "collect_check": False, "collect_expect": "", "collect_required": "", "collect_by": "", "collect_title": "",
    "pdf_grid_column": "", "pdf_password_column": "", "pdf_receiver_column": "", "pdf_encrypt": True,
    "pdf_watermark": True, "pdf_watermark_text": "{grid} {date}", "pdf_watermark_opacity": 0.15,
    "pdf_watermark_angle": 45,
}
RECIPE_KEYS = {
    "excel": ["split_column", "selected_values", "header_mode", "header_row", "grid_keys", "id_keys", "to_person",
              "person_column", "person_file_filter", "make_zip", "merge_across_files", "preserve_format",
              "keep_formulas", "keep_hidden", "exact_match", "value_alias_map", "skip_values", "fill_columns"],
    "collect": ["collect_check", "collect_expect", "collect_required", "collect_by", "collect_title",
                "header_mode", "header_row"],
    "pdf": ["pdf_grid_column", "pdf_password_column", "pdf_receiver_column", "pdf_encrypt", "pdf_watermark",
            "pdf_watermark_text", "pdf_watermark_opacity", "pdf_watermark_angle"],
}
MODE_NAMES = {"excel": "Excel 拆分", "collect": "Excel 汇总", "pdf": "文件加密分发"}
# 套用后清空的键：每次都不一样（截止时间）或者只属于上一次（下发记录）
CLEAR_ON_APPLY = {"excel": ("due",), "collect": ("collect_due", "collect_record"), "pdf": ()}
_RANGES = {"header_row": (1, 1000), "pdf_watermark_opacity": (0.0, 1.0), "pdf_watermark_angle": (-360, 360)}
_CHOICES = {"header_mode": ("auto", "row", "keyword")}
# 套用时要在表里找得到的列：(键, 说给人听的名字, 只在这个键为真时才查)
_COLUMN_KEYS = {"excel": [("split_column", "拆分字段", None), ("person_column", "到人字段", "to_person")],
                "collect": [("collect_by", "按哪一列分人", None)], "pdf": []}


class RecipeError(ValueError):
    """方案文件读不了、不是方案、模式不对：消息是可以直接给用户看的人话。"""


def _text(v, limit=MAX_TEXT):
    """文字设置：去掉控制字符，但保留换行和制表符（应交名单可以每行一个）。"""
    if v is None:
        return ""
    t = "".join(ch for ch in str(v) if ch in "\n\t" or not unicodedata.category(ch).startswith("C"))
    return t.strip()[:limit]


def _clean(key, value):
    """按默认值的类型校验一个设置。返回 (值, 问题说明)；类型对不上时值为 None。"""
    default = DEFAULTS[key]
    if isinstance(default, bool):
        return (value, "") if isinstance(value, bool) else (None, "应为是 / 否")
    if isinstance(default, (int, float)):
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            return None, "应为数字"
        lo, hi = _RANGES.get(key, (float("-inf"), float("inf")))
        if not lo <= value <= hi:
            return None, f"应在 {lo}～{hi} 之间"
        return (int(value) if isinstance(default, int) and float(value).is_integer() else value), ""
    if isinstance(default, str):
        if isinstance(value, (list, tuple)) and all(isinstance(x, str) for x in value):
            value = "、".join(value)              # 技能里名单常写成列表
        if not isinstance(value, str):
            return None, "应为文字"
        if key in _CHOICES and value not in _CHOICES[key]:
            return None, "取值不认识"
        return _text(value), ("太长，截断了" if len(value) > MAX_TEXT else "")
    if isinstance(default, list):
        if not isinstance(value, list) or not all(isinstance(x, (str, int, float)) for x in value):
            return None, "应为文字列表"
        items = [str(x)[:MAX_ITEM] for x in value[:MAX_ITEMS]]
        return items, ("太长，截断了" if len(value) > MAX_ITEMS else "")
    if isinstance(default, dict):           # value_alias_map：{规范值: [别名...]}
        if not isinstance(value, dict):
            return None, "应为 {规范值: [别名…]}"
        out = {}
        for k, v in list(value.items())[:MAX_ITEMS]:
            if isinstance(k, str) and isinstance(v, list) and all(isinstance(x, str) for x in v):
                out[k[:MAX_ITEM]] = [x[:MAX_ITEM] for x in v[:MAX_ITEMS]]
        return out, ("有几项格式不对，丢掉了" if len(out) < len(value) else "")
    return None, "类型不认识"


def make_recipe(cfg, mode, name, note="", columns=None, input_name="", app=""):
    """把当前设置存成方案 dict：只取这个模式白名单里的键；hints 只记列名和上次用的表名（不参与任何操作）。"""
    if mode not in RECIPE_KEYS:
        raise RecipeError(f"不认识的模式「{mode}」")
    settings = {}
    for k in RECIPE_KEYS[mode]:
        if k in cfg:
            v, _why = _clean(k, cfg[k])
            if v is not None:
                settings[k] = v
    hints = {}
    if input_name:
        hints["input_name"] = _text(os.path.basename(str(input_name)), MAX_ITEM)
    if columns:
        hints["columns"] = [str(c)[:MAX_ITEM] for c in list(columns)[:MAX_ITEMS] if str(c).strip()]
    return {RECIPE_KEY: RECIPE_VERSION, "name": _text(name, MAX_NAME) or MODE_NAMES[mode], "mode": mode,
            "app": str(app)[:20], "created": datetime.datetime.now().isoformat(timespec="seconds"),
            "note": _text(note, MAX_NOTE), "settings": settings, "hints": hints}


def save_recipe(recipe, path):
    """写成 UTF-8 JSON；目录不在就建。返回路径。"""
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(recipe, f, ensure_ascii=False, indent=1)
        f.write("\n")
    return path


def recipe_filename(name):
    """方案名 → 文件名（去掉文件名里不能用的字符）。"""
    from core.utils import safe_filename
    return (safe_filename(name) or "方案") + ".json"


def is_recipe(path):
    """这个 .json 是不是方案（拖放分流用：「下发记录.json」也是 .json，靠有没有 excelrouter_recipe 键区分）。"""
    try:
        if os.path.getsize(path) > 2_000_000:
            return False
        with open(path, encoding="utf-8-sig") as f:
            data = json.load(f)
        return isinstance(data, dict) and RECIPE_KEY in data
    except (OSError, ValueError):
        return False


def load_recipe(path, mode=None):
    """读方案并逐项校验，返回 (方案, 警告列表)。读不了、不是方案、模式不对（给了 mode 时）抛 RecipeError。"""
    try:
        if os.path.getsize(path) > 2_000_000:
            raise RecipeError("这个文件太大，不像是方案文件")
        with open(path, encoding="utf-8-sig") as f:
            data = json.load(f)
    except OSError as e:
        raise RecipeError(f"方案文件打不开：{e}") from e
    except ValueError as e:
        raise RecipeError("方案文件内容坏了（不是有效的 JSON），请重新存一份") from e
    if not isinstance(data, dict) or RECIPE_KEY not in data:
        raise RecipeError("这个 .json 不是 ExcelRouter 的方案文件")
    warnings = []
    ver = data.get(RECIPE_KEY)
    if not isinstance(ver, int) or ver > RECIPE_VERSION:
        warnings.append("这个方案是新版本做的，部分设置可能用不上")
    rmode = data.get("mode")
    if rmode not in RECIPE_KEYS:
        raise RecipeError(f"方案里的模式「{rmode}」不认识")
    if mode and rmode != mode:
        raise RecipeError(f"这是「{MODE_NAMES[rmode]}」的方案，不能用来做「{MODE_NAMES[mode]}」")
    raw = data.get("settings") if isinstance(data.get("settings"), dict) else {}
    settings, unknown = {}, []
    for k, v in raw.items():
        if k not in RECIPE_KEYS[rmode]:
            unknown.append(k)
            continue
        val, why = _clean(k, v)
        if val is None:
            warnings.append(f"设置「{k}」{why}，没有用")
            continue
        if why:
            warnings.append(f"设置「{k}」{why}")
        settings[k] = val
    if unknown:
        warnings.append(f"有 {len(unknown)} 项设置这个版本不支持，已忽略")
    hints = data.get("hints") if isinstance(data.get("hints"), dict) else {}
    cols = hints.get("columns")
    recipe = {RECIPE_KEY: RECIPE_VERSION, "name": _text(data.get("name"), MAX_NAME) or MODE_NAMES[rmode],
              "mode": rmode, "app": _text(data.get("app"), 20), "created": _text(data.get("created"), 40),
              "note": _text(data.get("note"), MAX_NOTE), "settings": settings,
              "hints": {"input_name": _text(hints.get("input_name"), MAX_ITEM),
                        "columns": [str(c)[:MAX_ITEM] for c in cols[:MAX_ITEMS]] if isinstance(cols, list) else []}}
    return recipe, warnings


def apply_recipe(cfg, recipe):
    """把方案套到配置上，返回 (新配置, 改了哪些键, 说明)。路径不动；截止时间、下发记录清空（每次都不一样）。"""
    mode = recipe["mode"]
    new = dict(cfg)
    changed = [k for k, v in recipe["settings"].items() if cfg.get(k) != v]
    new.update(recipe["settings"])
    new["ui_mode"] = mode
    for k in CLEAR_ON_APPLY[mode]:
        new[k] = ""
    notes = []
    if mode == "collect" and recipe["settings"].get("collect_expect"):
        from core.collect import split_names
        notes.append(f"方案里有应交名单 {len(split_names(recipe['settings']['collect_expect']))} 个")
    return new, changed, notes


def summary(recipe):
    """一句话说清方案做什么：给日志和技能的 message 用。"""
    s = recipe["settings"]
    bits = []
    for k, label in (("split_column", "拆分字段"), ("person_column", "到人字段"), ("fill_columns", "填写列"),
                     ("collect_by", "按列分人"), ("collect_required", "必须填的列"), ("pdf_grid_column", "网格列")):
        if s.get(k) and (k != "person_column" or s.get("to_person")):
            bits.append(f"{label} = {s[k]}")
    if s.get("collect_expect"):
        from core.collect import split_names
        bits.append(f"应交名单 {len(split_names(s['collect_expect']))} 个")
    return "；".join(bits) or "（用默认设置）"


def check_columns(recipe, columns):
    """套用前比一比：方案里要用的列这张表里有没有、表的结构和上次像不像。返回警告列表（空 = 没问题）。

    找不到的列只提醒，调用方不能自动换成推荐的列——静默换列会按错的列拆完才发现。
    """
    have = {soft_clean(c) for c in columns or [] if soft_clean(c)}
    s, warnings = recipe["settings"], []
    for key, label, need in _COLUMN_KEYS[recipe["mode"]]:
        want = soft_clean(s.get(key, ""))
        if want and (need is None or s.get(need)) and want not in have:
            warnings.append(f"方案里的{label}「{want}」这张表里没有")
    old = {soft_clean(c) for c in recipe.get("hints", {}).get("columns", []) if soft_clean(c)}
    if old and have and len(old & have) * 2 < len(old):
        warnings.append("这张表和方案上次用的表不太一样（列名对上的不到一半），请确认选对了表")
    return warnings
