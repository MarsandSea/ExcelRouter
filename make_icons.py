#!/usr/bin/env python3
"""
make_icons.py —— 生成 ExcelRouter 的多尺寸图标。

为什么需要它：
  主图标（docs/logo.svg，源表 → 三个文件）在 32px 以上很好看，但 16/24px 下
  三个目标块和连接线会糊成一团 —— 而浏览器标签页的 favicon、Windows 任务栏和
  资源管理器「小图标」视图恰恰就是这两个尺寸。把一张 256px 的图缩到底，是
  「.ico 里有 7 个尺寸帧但小尺寸照样糊」的根因。

  解法不是换成字母缩写（那样就丢掉了「一眼看出在干什么」这个唯一优势），
  而是给小尺寸画一版简化的：去掉纸面行线、三个目标块减成两个、线条加粗、
  间距拉开。保住「一分多」的意思，16px 也立得住。苹果/微软的系统图标都是
  这么做的，不是一张图缩到底。

  几何参数只在本文件写一次（SMALL_* 常量），同时产出矢量源和位图帧，
  避免 SVG 和 PNG 各改各的对不上。

用法：
  py make_icons.py            # 重新生成所有图标 + 预览图
  py make_icons.py --preview  # 只出预览图，不覆盖任何图标

产出：
  docs/logo-small.svg   简化版矢量源（16/24/32px 用）
  app.ico               Windows exe/窗口图标，7 个尺寸帧，小帧用简化版
  docs/favicon.ico      落地页 favicon，4 个尺寸帧，同上
  icon-preview.png      各尺寸实际效果对照（不入库，看完可删）

依赖：只要 Pillow。刻意不引 cairosvg —— 它在 Windows 上要装 cairo 原生库，
为了几个图标不值得，而这个形状用 ImageDraw 画出来和 SVG 一模一样。
"""
from __future__ import annotations

import argparse
import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))

# ---- 配色：与 docs/logo.svg 完全一致，缩放时不会看出是两张图 ----
BG_TOP = (0x35, 0xA5, 0x66)
BG_MID = (0x1E, 0x7F, 0x4B)
BG_BOT = (0x12, 0x60, 0x3A)
PAPER = (0xFF, 0xFF, 0xFF)

# ---- 简化版几何（256 坐标系，与主图标同一个画布尺寸）----
# 关键约束来自 16px：去掉 1px 安全边后只剩 14 个可用像素，而 256 坐标系里
# 16 个单位才等于 1 个像素。也就是说——
#   · 每个白块至少要 48 单位（3px）才看得出是个块
#   · 块与块之间至少留 32 单位（2px）才不会粘连
# 按这个尺子量一遍就会发现：「源表 + 分叉线 + 三个目标块」在 16px 下根本排不下，
# 无论线条画多粗。所以简化版**直接去掉连接线**，只靠「左边一个大块 → 右边两个
# 小块」的排布表达拆分。少即是多，这是 16px 唯一能成立的写法。
PLATE = (12, 12, 244, 244)          # 底板外框
PLATE_R = 52                        # 圆角
SRC = (36, 56, 112, 200)            # 源表：76 × 144（4.75px × 9px @16）
TGT_TOP = (144, 56, 224, 112)       # 目标块 1：80 × 56
TGT_BOT = (144, 144, 224, 200)      # 目标块 2：80 × 56
# 三处间距都刻意 ≥32 单位（2px @16）：源表↔目标列 32，两个目标块之间 32。

SS = 8                              # 超采样倍率：先画 8 倍大再缩，等效抗锯齿


def _plate(draw: ImageDraw.ImageDraw, scale: float) -> None:
    """画底板的纵向渐变（圆角由调用方用蒙版裁）。

    scale 必须是「目标画布 / 256」，不是超采样倍率 —— 两者只有在
    size == 256 时才相等，混用会把底板画到画布外面去。
    """
    x0, y0, x1, y1 = (round(v * scale) for v in PLATE)
    h = y1 - y0
    for i in range(h):
        t = i / max(h - 1, 1)
        if t < 0.55:
            k = t / 0.55
            c = tuple(round(a + (b - a) * k) for a, b in zip(BG_TOP, BG_MID))
        else:
            k = (t - 0.55) / 0.45
            c = tuple(round(a + (b - a) * k) for a, b in zip(BG_MID, BG_BOT))
        draw.line([(x0, y0 + i), (x1, y0 + i)], fill=c + (255,))


def render_small(size: int) -> Image.Image:
    """渲染简化版图标到指定像素尺寸。"""
    big = size * SS
    canvas = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    scale = big / 256

    # 底板（先画满渐变，再用圆角蒙版裁）
    plate = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    _plate(ImageDraw.Draw(plate), scale)
    mask = Image.new("L", (big, big), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [v * scale for v in PLATE], radius=PLATE_R * scale, fill=255
    )
    canvas.paste(plate, (0, 0), mask)

    d = ImageDraw.Draw(canvas)
    for box, r in ((SRC, 16), (TGT_TOP, 14), (TGT_BOT, 14)):
        d.rounded_rectangle([v * scale for v in box], radius=r * scale, fill=PAPER + (255,))

    return canvas.resize((size, size), Image.LANCZOS)


def render_full(size: int) -> Image.Image:
    """主图标（app.png 那张 256px 成品）缩到指定尺寸。"""
    src = Image.open(os.path.join(HERE, "app.png")).convert("RGBA")
    return src.resize((size, size), Image.LANCZOS)


SVG_TEMPLATE = '''<svg viewBox="0 0 256 256" width="256" height="256" xmlns="http://www.w3.org/2000/svg" role="img">
  <!-- Copyright (c) 2026 AbeLin · MIT License -->
  <title>ExcelRouter（小尺寸简化版）</title>
  <desc>16/24px 专用：只保留「一个源表 → 两个文件」三个块，去掉纸面行线与分叉连线
        —— 16px 下每个块至少要 3 像素、间距至少 2 像素，连线再怎么加粗也排不下。
        几何参数的真源是仓库根目录的 make_icons.py，改这份 SVG 不会影响
        app.ico / favicon.ico，要改请改脚本再重新生成。</desc>
  <defs>
    <linearGradient id="bg" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#35A566"/>
      <stop offset="0.55" stop-color="#1E7F4B"/>
      <stop offset="1" stop-color="#12603A"/>
    </linearGradient>
  </defs>
  <rect x="{px0}" y="{py0}" width="{pw}" height="{ph}" rx="{pr}" fill="url(#bg)"/>
  <rect x="{sx0}" y="{sy0}" width="{sw}" height="{sh}" rx="16" fill="#FFFFFF"/>
  <rect x="{tx0}" y="{tty0}" width="{tw}" height="{th}" rx="14" fill="#FFFFFF"/>
  <rect x="{tx0}" y="{tby0}" width="{tw}" height="{th}" rx="14" fill="#FFFFFF"/>
</svg>
'''


def write_svg(path: str) -> None:
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(SVG_TEMPLATE.format(
            px0=PLATE[0], py0=PLATE[1], pw=PLATE[2] - PLATE[0], ph=PLATE[3] - PLATE[1], pr=PLATE_R,
            sx0=SRC[0], sy0=SRC[1], sw=SRC[2] - SRC[0], sh=SRC[3] - SRC[1],
            tx0=TGT_TOP[0], tty0=TGT_TOP[1], tby0=TGT_BOT[1],
            tw=TGT_TOP[2] - TGT_TOP[0], th=TGT_TOP[3] - TGT_TOP[1],
        ))


# 分界线定在 24：16/24 用简化版（主图标在这两档必糊），32 及以上用主图标
# （那里连接线和纸面行线已经分得开，而简化版反而显得空）。看 icon-preview.png
# 第三行就是实际写进 ico 的组合。
SMALL_SIZES = (16, 24)
ICO_SIZES = (16, 24, 32, 48, 64, 128, 256)
FAVICON_SIZES = (16, 24, 32, 48)


def frames(sizes) -> list[Image.Image]:
    return [render_small(s) if s in SMALL_SIZES else render_full(s) for s in sizes]


def make_preview(path: str) -> None:
    """三行对照：简化版 / 主图标 / 实际写进 ico 的那一版。"""
    sizes = [16, 24, 32, 48, 64]
    zoom, pad, gap = 6, 20, 16
    rows = [
        ("simplified", [render_small(s) for s in sizes]),
        ("full art", [render_full(s) for s in sizes]),
        ("as shipped", [render_small(s) if s in SMALL_SIZES else render_full(s) for s in sizes]),
    ]
    rowh = max(sizes) * zoom
    w = pad * 2 + sum(s * zoom for s in sizes) + gap * (len(sizes) - 1)
    h = pad + len(rows) * (rowh + pad)
    canvas = Image.new("RGBA", (w, h), (245, 245, 247, 255))
    for r, (_label, imgs) in enumerate(rows):
        x, y = pad, pad + r * (rowh + pad)
        for s, im in zip(sizes, imgs):
            big = im.resize((s * zoom, s * zoom), Image.NEAREST)
            canvas.paste(big, (x, y + (rowh - s * zoom) // 2), big)
            x += s * zoom + gap
    canvas.save(path)


def main() -> int:
    ap = argparse.ArgumentParser(description="生成 ExcelRouter 多尺寸图标")
    ap.add_argument("--preview", action="store_true", help="只出预览图，不覆盖图标文件")
    args = ap.parse_args()

    preview = os.path.join(HERE, "icon-preview.png")
    make_preview(preview)
    print(f"预览：{preview}")
    if args.preview:
        return 0

    svg = os.path.join(HERE, "docs", "logo-small.svg")
    write_svg(svg)
    print(f"矢量源：{svg}")

    for out, sizes in ((os.path.join(HERE, "app.ico"), ICO_SIZES),
                       (os.path.join(HERE, "docs", "favicon.ico"), FAVICON_SIZES)):
        ims = frames(sizes)
        # Pillow 的 ICO 保存：以最大的那张为底，sizes 指定要写进去的所有帧。
        # 但它会自己缩放，缩出来的小帧又是糊的 —— 所以改用 append_images
        # 显式给每一帧，这是「ico 里有 7 个尺寸却照样糊」的修法。
        ims[-1].save(out, format="ICO", sizes=[(s, s) for s in sizes],
                     append_images=ims[:-1])
        print(f"{out}：{len(sizes)} 个尺寸帧 {sizes}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
