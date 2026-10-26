"""core/utils.py 的单元测试。"""
import math

from core.utils import soft_clean, hard_clean, apply_alias, is_skip_value, safe_filename


def test_soft_clean_basic():
    assert soft_clean("  张三 ") == "张三"
    assert soft_clean(None) == ""
    assert soft_clean(float("nan")) == ""          # pandas 空单元格
    assert soft_clean("北　京") == "北京"            # 全角空格被清掉
    assert soft_clean(123) == "123"


def test_hard_clean():
    assert hard_clean("【网格】A-1") == "网格A1"
    assert hard_clean(None) == ""
    assert hard_clean("商企 经理") == "商企经理"


def test_apply_alias():
    amap = {"商企": ["商企经理", "商企"], "行销": ["行销经理", "市场"]}
    assert apply_alias("商企经理", amap) == "商企"
    assert apply_alias("市场部", amap) == "行销"
    assert apply_alias("装维", amap) == "装维"        # 未命中 -> 原值
    assert apply_alias(" 销售部 ", {}) == "销售部"     # 空映射 = soft_clean


def test_is_skip_value():
    skip = ["合计", "小计", ""]
    assert is_skip_value("", skip) is True
    assert is_skip_value(None, skip) is True
    assert is_skip_value("合计", skip) is True
    assert is_skip_value("小 计", skip) is True        # hard_clean 后匹配
    assert is_skip_value("销售部", skip) is False
    assert is_skip_value("销售部", None) is False
    assert is_skip_value("", None) is True             # 空值永远跳过


def test_safe_filename_windows_device_names():
    """CON、NUL、COM1 这类设备名在 Windows 上建不出文件夹：前面加下划线（启明星辰 2026-10-26）。"""
    assert [safe_filename(n) for n in ("CON", "nul", "Com1", "LPT9.txt", "AUX")] == ["_CON", "_nul", "_Com1", "_LPT9.txt", "_AUX"]
    assert safe_filename("CONSOLE") == "CONSOLE" and safe_filename("东区COM1") == "东区COM1"


def test_safe_filename():
    assert safe_filename("销售/部:东?区") == "销售_部_东_区"
    assert safe_filename("  ") == "未命名"
    assert safe_filename("正常名称") == "正常名称"
