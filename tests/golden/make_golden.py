"""生成拆分对比测试的标准答案：py tests/golden/make_golden.py

只在需要重定标准答案时跑（例如有意改变了拆分输出）。v3.0 的标准答案是在 v2.10.1（pandas 读表路径）上生成的。
每个场景跑两遍，两遍的快照必须一样，保证快照里没有时间戳之类每次都变的东西。
产出写在系统临时目录，跑完删掉；标准答案写进 tests/golden/expected/。

Copyright (c) 2026 AbeLin · MIT License
"""
import json
import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(os.path.dirname(HERE)))   # 仓库根：import core
sys.path.insert(0, os.path.dirname(HERE))                    # tests/：import golden

from golden import samples, snapshot  # noqa: E402


def main():
    exp_dir = os.path.join(HERE, "expected")
    os.makedirs(exp_dir, exist_ok=True)
    only = set(sys.argv[1:])
    for name in samples.NAMES:
        if only and name not in only:
            continue
        shots = []
        for _ in range(2):
            root = tempfile.mkdtemp(prefix="er-golden-")
            try:
                shots.append(snapshot.run_scenario(name, root))
            finally:
                shutil.rmtree(root, ignore_errors=True)
        diffs = snapshot.diff(shots[0], shots[1])
        if diffs:
            raise SystemExit(f"{name}: 两次运行的快照不一样：\n" + "\n".join(diffs))
        with open(os.path.join(exp_dir, f"{name}.json"), "w", encoding="utf-8") as f:
            json.dump(shots[0], f, ensure_ascii=False, indent=1, sort_keys=True)
            f.write("\n")
        print(f"{name}: {len(shots[0])} 个文件")


if __name__ == "__main__":
    main()
