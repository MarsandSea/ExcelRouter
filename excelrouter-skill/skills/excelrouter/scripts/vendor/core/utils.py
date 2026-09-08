"""
Excel 智能拆分工具 - 文本清理与通用工具
Copyright (c) 2026 AbeLin
MIT License
"""

import re
import unicodedata


def soft_clean(s):
    """温和清理：去掉控制字符和各种不可见空白，保留可见内容。

    None 与 pandas 的 NaN 一律视为空字符串，避免空单元格被当成有内容。
    """
    if s is None:
        return ""
    # pandas 空单元格是 float('nan')，nan != nan
    if isinstance(s, float) and s != s:
        return ""
    s = str(s)
    s = ''.join(ch for ch in s if not unicodedata.category(ch).startswith('C'))
    for ch in ('　', '\xa0', '​', '‌', '‍', '﻿'):
        s = s.replace(ch, '')
    return s.strip()


def hard_clean(s):
    """强力清理：只保留字母、数字、中文，用于表头/列名匹配。"""
    if s is None:
        return ""
    if isinstance(s, float) and s != s:
        return ""
    return re.sub(r'[^\w一-龥]', '', str(s))


def apply_alias(value, alias_map):
    """取值归并：把别名映射成规范值（旧 position_map 的通用化身）。

    alias_map 格式：{规范值: [别名列表]}。按 key 顺序优先匹配（用 hard_clean 做包含匹配）。
    未命中任何别名时，返回 soft_clean 后的原值。

    例：{"商企": ["商企经理", "商企"]} 会把"商企经理""商企"都归并成"商企"。
    alias_map 为空时等价于 soft_clean。
    """
    v = soft_clean(value)
    if not alias_map:
        return v
    vc = hard_clean(v)
    if not vc:
        return v
    for canonical, aliases in alias_map.items():
        for a in aliases:
            ac = hard_clean(a)
            if ac and ac in vc:
                return canonical
    return v


def is_skip_value(value, skip_set=None):
    """判断一个取值是否应被跳过（空值 / 合计 / 小计等非真实分组）。

    空值（清理后为空）永远跳过。其余按 hard_clean 后小写与 skip_set 比较。
    """
    vc = hard_clean(value).lower()
    if vc == "":
        return True
    if not skip_set:
        return False
    return vc in {hard_clean(s).lower() for s in skip_set if hard_clean(s)}


def safe_filename(name, fallback="未命名"):
    """把任意取值转成合法的 Windows 文件名/文件夹名片段。"""
    s = soft_clean(name)
    s = re.sub(r'[\\/:*?"<>|]', '_', s)   # Windows 非法字符
    s = s.strip('. ')                      # 结尾的点和空格在 Windows 上非法
    return s or fallback
