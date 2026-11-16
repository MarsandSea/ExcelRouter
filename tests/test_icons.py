# Copyright (c) 2026 AbeLin · MIT License
"""图标一致性：仓库里的图标文件必须是 make_icons.py 现画出来的那一份。

守两件事：
  1. app.ico / app.png / favicon.ico 没被手改、也没忘了重新生成——
     改图标只能改 make_icons.py 再跑一遍，否则这里会红。
  2. 任务栏用的 16/20/24px 小帧和桌面大图标是同一个图形（有分叉连线）。
     v3.2.1 之前小帧只画「一个高块 → 三个小块」，任务栏和桌面看着像两个图标。
"""
import os

import pytest

Image = pytest.importorskip("PIL.Image")
make_icons = pytest.importorskip("make_icons")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _diff(a, b):
    """两张同尺寸 RGBA 图逐通道的平均差和最大差。"""
    a, b = a.convert("RGBA"), b.convert("RGBA")
    assert a.size == b.size
    da, db = a.tobytes(), b.tobytes()
    diffs = [abs(x - y) for x, y in zip(da, db)]
    return sum(diffs) / len(diffs), max(diffs)


def _frames(path):
    im = Image.open(path)
    return {s[0]: im.ico.getimage(s) for s in im.info["sizes"]}


@pytest.mark.parametrize("rel, sizes", [
    ("app.ico", make_icons.ICO_SIZES),
    (os.path.join("docs", "favicon.ico"), make_icons.FAVICON_SIZES),
])
def test_ico_matches_generator(rel, sizes):
    frames = _frames(os.path.join(ROOT, rel))
    assert sorted(frames) == sorted(sizes), f"{rel} 的尺寸帧不对，重新跑 py make_icons.py"
    for s in sizes:
        mean, _ = _diff(frames[s], make_icons.render(s))
        # 不同 Pillow 版本抗锯齿会差一点点，平均差留 2 的余量
        assert mean < 2, f"{rel} 的 {s}px 帧和 make_icons.py 画的不一样（平均差 {mean:.1f}），重新跑 py make_icons.py"


def test_app_png_matches_generator():
    mean, _ = _diff(Image.open(os.path.join(ROOT, "app.png")), make_icons.render(256))
    assert mean < 2, "app.png 和 make_icons.py 画的不一样，重新跑 py make_icons.py"


@pytest.mark.parametrize("size", [16, 20, 24])
def test_small_frames_keep_branch_lines(size):
    """小帧里源表和三个文件块之间要有白色连线（分叉竖线 + 中间横线）。"""
    im = _frames(os.path.join(ROOT, "app.ico"))[size].convert("RGB")
    k = size / 16

    def white(col, row):
        r, g, b = im.getpixel((int((col + 0.5) * k), int((row + 0.5) * k)))
        return min(r, g, b) > 200

    assert white(8, 4) and white(8, 10), f"{size}px 帧没有分叉竖线"
    assert white(7, 7), f"{size}px 帧没有源表到竖线的连线"
    assert not white(7, 4), f"{size}px 帧连线和源表粘成一块了"


def test_builds_use_the_generated_icons():
    """Windows 打包用 app.ico，Linux 打包和 .desktop 用 app.png，别换成别的文件。"""
    def read(rel):
        with open(os.path.join(ROOT, rel), encoding="utf-8", errors="replace") as f:
            return f.read()

    win = read(os.path.join(".github", "workflows", "release.yml"))
    assert win.count("--icon app.ico") == 2 and win.count('--add-data "app.ico;."') == 2
    bat = read("build.bat")
    assert bat.count("--icon app.ico") >= 2 and bat.count('--add-data "app.ico;."') >= 2
    assert '--add-data "app.png:."' in read("build_linux.sh")
    assert "_internal/app.png" in read(os.path.join("packaging", "install-desktop.sh"))
