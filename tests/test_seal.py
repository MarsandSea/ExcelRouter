"""core/seal.py：Windows 数据保护接口封存短密钥。

Copyright (c) 2026 AbeLin · MIT License
"""
import sys

import pytest

from core import seal

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Windows 数据保护接口只在 Windows 上有")


def test_round_trip_and_not_plain():
    key = bytes(range(16))
    text = seal.seal(key)
    assert seal.unseal(text) == key and key.hex() not in text and seal.available()
    assert seal.seal(key) != text                     # 每次封出来都不一样（带随机盐）


def test_tampered_or_garbage_is_refused():
    text = seal.seal(b"0123456789abcdef")
    i = len(text) - 12                                # 改到受保护的内容里（开头的头信息不校验）
    broken = text[:i] + ("A" if text[i] != "A" else "B") + text[i + 1:]
    for bad in (broken, "不是base64", ""):
        with pytest.raises(ValueError):
            seal.unseal(bad)
