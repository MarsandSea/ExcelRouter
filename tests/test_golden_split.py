"""拆分对比测试：同一批样本，拆出来的文件清单、每个单元格的值 / 类型 / 格式、合并单元格、列宽、ZIP 成员
都必须和 tests/golden/expected/ 里的标准答案一致。

标准答案在 v2.10.1（pandas 读表路径）上生成，用来守住 v3.0 换成纯 openpyxl 读表后行为不变。
有意改变的行为不在这里，见各自的专门测试。重定标准答案：py tests/golden/make_golden.py

Copyright (c) 2026 AbeLin · MIT License
"""
import json
import os

import pytest

from golden import samples, snapshot

EXPECTED = os.path.join(os.path.dirname(os.path.abspath(__file__)), "golden", "expected")


@pytest.mark.parametrize("name", samples.NAMES)
def test_split_matches_golden(tmp_path, name):
    actual = json.loads(json.dumps(snapshot.run_scenario(name, str(tmp_path))))
    with open(os.path.join(EXPECTED, f"{name}.json"), encoding="utf-8") as f:
        expected = json.load(f)
    diffs = snapshot.diff(expected, actual)
    assert not diffs, f"场景 {name} 和标准答案不一致：\n" + "\n".join(diffs)
