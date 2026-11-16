#!/usr/bin/env python3
# Copyright (c) 2026 AbeLin · MIT License
"""
make_icons.py —— 生成 ExcelRouter 的多尺寸图标。

图形只有一个意思：「一张表 → 三个文件」（一转三），所有尺寸都画这个，
任务栏小图标和桌面大图标看起来是同一个东西。

为什么全部用代码画、不再缩 app.png：
  旧版 32px 以上的帧是从 app.png（macOS 风格成品）缩出来的。那张图四周留了
  约 10% 透明边、还带一圈投影——缩到桌面的 48px 后投影糊成灰色脏边，绿块也
  比旁边的 Windows 图标小一圈。Windows 图标不带投影、底板铺满画布，这里照做。

两档画法（分界线在 24），**图形结构一样**：源表 → 分叉连线 → 三个文件块。
  · 32px 及以上（完整版）：源表带行线、分叉拐角是圆弧。
  · 16/20/24px（像素版）：同一个结构按 16px 的像素格重画——源表 4×10、
    连线 1px、三个文件块 4×3（块间 2px）。256 坐标系里 16 个单位 = 1 像素，
    坐标都取 16 的倍数，16px 帧像素对齐不糊边；去掉的只有行线和圆角。
    v3.2.1 之前这一档只画「一个高块 → 三个小块」、没有连线，任务栏（16/24px）
    和桌面（48px 起）看着像两个图标，作者两次指出过——**别再把连线去掉**，
    tests/test_icons.py 会检查小帧里有没有连线。

用法：
  py make_icons.py                  # 重新生成所有图标
  py make_icons.py --preview PATH   # 只出预览图到 PATH，不覆盖任何图标

产出：
  app.ico               Windows exe / 窗口图标，10 个尺寸帧
  app.png               Linux 窗口与开始菜单图标（256px）
  docs/favicon.ico      落地页 favicon
  docs/logo-small.svg   像素版矢量源（只是给人看的，真源是本文件）
  docs/logo.png / logo.svg 是网页上的大图，带投影的 macOS 风格，不由本脚本生成。

依赖：只要 Pillow。
"""
from __future__ import annotations

import argparse
import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))

# ---- 配色：与 docs/logo.svg 一致 ----
BG_TOP = (0x35, 0xA5, 0x66)
BG_MID = (0x1E, 0x7F, 0x4B)
BG_BOT = (0x12, 0x60, 0x3A)
PAPER = (0xFF, 0xFF, 0xFF)
INK = BG_MID                         # 纸面行线

SS = 8                               # 超采样倍率：先画 8 倍大再缩，等效抗锯齿

# ---- 完整版几何（256 坐标系）----
PLATE = (8, 8, 248, 248)             # 底板几乎铺满，只留 0.5px@16 的安全边
PLATE_R = 52
SRC = (36, 70, 106, 186)             # 源表
SRC_R = 13
SRC_LINES = (102, 128, 154)          # 源表行线的 y
TGT_X = (162, 216)                   # 三个文件块的 x 范围
TGT_YS = ((40, 90), (103, 153), (166, 216))   # 块间隔 13，32px 也分得开
TGT_R = 12
STROKE = 10                          # 连线粗细（32px 时约 1.25px）
TRUNK_X = 136                        # 分叉竖线的 x
BEND_R = 14                          # 分叉拐角半径
LINE_W = 8                           # 行线粗细

# ---- 像素版几何（16/20/24px；坐标是 16 的倍数，16px 帧像素对齐）----
SMALL_PLATE_R = 40                   # 圆角收小一点，右上 / 右下角的块才不会被切
SMALL_SRC = (32, 48, 96, 208)        # 源表：4 × 10 px@16（第 2～5 列）
SMALL_TGT_X = (160, 224)             # 文件块：4 × 3 px@16（第 10～13 列）
SMALL_TGT_YS = ((16, 64), (96, 144), (176, 224))   # 3px 高、间隔 2px
SMALL_LINKS = (                      # 连线，全是 1px 宽的矩形
    (96, 112, 128, 128),             #   源表 → 竖线（第 7 行）
    (128, 32, 144, 208),             #   分叉竖线（第 8 列，第 2～12 行）
    (144, 32, 160, 48),              #   → 上块
    (144, 112, 160, 128),            #   → 中块
    (144, 192, 160, 208),            #   → 下块
)


def _plate(size: int, radius: float) -> Image.Image:
    """渐变圆角底板（画布已是超采样尺寸）。"""
    s = size / 256
    x0, y0, x1, y1 = (round(v * s) for v in PLATE)
    grad = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(grad)
    h = y1 - y0
    for i in range(h):
        t = i / max(h - 1, 1)
        if t < 0.55:
            a, b, k = BG_TOP, BG_MID, t / 0.55
        else:
            a, b, k = BG_MID, BG_BOT, (t - 0.55) / 0.45
        c = tuple(round(p + (q - p) * k) for p, q in zip(a, b))
        d.line([(x0, y0 + i), (x1, y0 + i)], fill=c + (255,))
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle([x0, y0, x1, y1], radius=radius * s, fill=255)
    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(grad, (0, 0), mask)
    return out


def _hline(d: ImageDraw.ImageDraw, x0: float, x1: float, y: float, w: float, fill) -> None:
    """圆头横线。"""
    d.line([(x0, y), (x1, y)], fill=fill, width=round(w))
    r = w / 2
    for x in (x0, x1):
        d.ellipse([x - r, y - r, x + r, y + r], fill=fill)


def render_full(size: int) -> Image.Image:
    """完整版：源表 → 分叉 → 三个文件。"""
    big = size * SS
    s = big / 256
    im = _plate(big, PLATE_R)
    d = ImageDraw.Draw(im)
    white = PAPER + (255,)
    ink = INK + (255,)
    w = STROKE * s

    # 连线：源表右边 → 分叉竖线 → 三个文件块左边
    mid = (TGT_YS[1][0] + TGT_YS[1][1]) / 2
    top = (TGT_YS[0][0] + TGT_YS[0][1]) / 2
    bot = (TGT_YS[2][0] + TGT_YS[2][1]) / 2
    x_src, x_tgt = SRC[2] * s, TGT_X[0] * s
    _hline(d, x_src, x_tgt, mid * s, w, white)
    tx, br = TRUNK_X * s, BEND_R * s
    d.line([(tx, (top + BEND_R) * s), (tx, (bot - BEND_R) * s)], fill=white, width=round(w))
    for cy, start, end in ((top + BEND_R, 180, 270), (bot - BEND_R, 90, 180)):
        cx = tx + br
        c = cy * s
        d.arc([cx - br - w / 2, c - br - w / 2, cx + br + w / 2, c + br + w / 2],
              start, end, fill=white, width=round(w))
    for y in (top, bot):
        _hline(d, tx + br, x_tgt, y * s, w, white)

    # 源表 + 行线
    d.rounded_rectangle([v * s for v in SRC], radius=SRC_R * s, fill=white)
    for y in SRC_LINES:
        _hline(d, (SRC[0] + 14) * s, (SRC[2] - 14) * s, y * s, LINE_W * s, ink)

    # 三个文件块 + 各一条行线
    for y0, y1 in TGT_YS:
        d.rounded_rectangle([TGT_X[0] * s, y0 * s, TGT_X[1] * s, y1 * s], radius=TGT_R * s, fill=white)
        if size >= 64:      # 再小这条线就糊进白块里了
            _hline(d, (TGT_X[0] + 13) * s, (TGT_X[1] - 13) * s, (y0 + y1) / 2 * s, LINE_W * s, ink)

    return im.resize((size, size), Image.LANCZOS)


def render_small(size: int) -> Image.Image:
    """像素版：源表 → 分叉 → 三个文件，和完整版同一个结构，只去掉行线和圆弧。"""
    big = size * SS
    s = big / 256
    im = _plate(big, SMALL_PLATE_R)
    d = ImageDraw.Draw(im)
    white = PAPER + (255,)
    d.rounded_rectangle([v * s for v in SMALL_SRC], radius=8 * s, fill=white)
    for r in SMALL_LINKS:
        d.rectangle([v * s for v in r], fill=white)
    for y0, y1 in SMALL_TGT_YS:
        d.rounded_rectangle([SMALL_TGT_X[0] * s, y0 * s, SMALL_TGT_X[1] * s, y1 * s],
                            radius=8 * s, fill=white)
    return im.resize((size, size), Image.LANCZOS)


SMALL_SIZES = (16, 20, 24)
# 20 / 40 / 96 是 125%、150% 缩放下任务栏和桌面实际取的尺寸，
# 缺了 Windows 会拿相邻帧自己缩，又是糊的。
ICO_SIZES = (16, 20, 24, 32, 40, 48, 64, 96, 128, 256)
FAVICON_SIZES = (16, 24, 32, 48)


def render(size: int) -> Image.Image:
    return render_small(size) if size in SMALL_SIZES else render_full(size)


SVG_TEMPLATE = '''<svg viewBox="0 0 256 256" width="256" height="256" xmlns="http://www.w3.org/2000/svg" role="img">
  <!-- Copyright (c) 2026 AbeLin · MIT License -->
  <title>ExcelRouter（小尺寸像素版）</title>
  <desc>16/20/24px 专用：源表 → 分叉连线 → 三个文件，和大图标同一个结构，
        只去掉纸面行线和圆弧拐角，按 16px 像素格对齐。
        几何参数的真源是仓库根目录的 make_icons.py，改这份 SVG 不会影响
        app.ico / favicon.ico，要改请改脚本再重新生成。</desc>
  <defs>
    <linearGradient id="bg" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#35A566"/>
      <stop offset="0.55" stop-color="#1E7F4B"/>
      <stop offset="1" stop-color="#12603A"/>
    </linearGradient>
  </defs>
  <rect x="{px}" y="{py}" width="{pw}" height="{ph}" rx="{pr}" fill="url(#bg)"/>
  <rect x="{sx}" y="{sy}" width="{sw}" height="{sh}" rx="8" fill="#FFFFFF"/>
{links}
{targets}
</svg>
'''


def write_svg(path: str) -> None:
    tx0, tx1 = SMALL_TGT_X
    targets = "\n".join(
        f'  <rect x="{tx0}" y="{y0}" width="{tx1 - tx0}" height="{y1 - y0}" rx="8" fill="#FFFFFF"/>'
        for y0, y1 in SMALL_TGT_YS)
    links = "\n".join(
        f'  <rect x="{x0}" y="{y0}" width="{x1 - x0}" height="{y1 - y0}" fill="#FFFFFF"/>'
        for x0, y0, x1, y1 in SMALL_LINKS)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(SVG_TEMPLATE.format(
            px=PLATE[0], py=PLATE[1], pw=PLATE[2] - PLATE[0], ph=PLATE[3] - PLATE[1], pr=SMALL_PLATE_R,
            sx=SMALL_SRC[0], sy=SMALL_SRC[1],
            sw=SMALL_SRC[2] - SMALL_SRC[0], sh=SMALL_SRC[3] - SMALL_SRC[1],
            links=links, targets=targets,
        ))


def make_preview(path: str) -> None:
    """浅色 / 深色两种桌面背景上，各尺寸按 1:1 和放大 4 倍各摆一遍。"""
    sizes = [16, 20, 24, 32, 40, 48, 64, 96]
    zoom, pad, gap = 4, 24, 20
    big = render(256)
    w = pad * 2 + sum(s * zoom for s in sizes) + gap * (len(sizes) - 1) + gap + 256
    rowh = max(max(sizes) * zoom, 256)
    h = pad + 2 * (rowh + pad + 96 + pad)
    canvas = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    y = pad
    for bg in ((243, 243, 243, 255), (32, 32, 32, 255)):
        band = Image.new("RGBA", (w, rowh + 96 + pad * 2), bg)
        canvas.paste(band, (0, y - pad // 2))
        x = pad
        for s in sizes:
            im = render(s)
            canvas.paste(im, (x, y), im)                            # 1:1
            z = im.resize((s * zoom, s * zoom), Image.NEAREST)
            canvas.paste(z, (x, y + 96), z)                         # 放大看像素
            x += s * zoom + gap
        canvas.paste(big, (x, y + (rowh + 96 - 256) // 2), big)
        y += rowh + 96 + pad * 2
    canvas.save(path)


def main() -> int:
    ap = argparse.ArgumentParser(description="生成 ExcelRouter 多尺寸图标")
    ap.add_argument("--preview", metavar="PATH", help="只出预览图到 PATH，不覆盖图标文件")
    args = ap.parse_args()

    if args.preview:
        make_preview(args.preview)
        print(f"预览：{args.preview}")
        return 0

    svg = os.path.join(HERE, "docs", "logo-small.svg")
    write_svg(svg)
    print(f"矢量源：{svg}")

    png = os.path.join(HERE, "app.png")
    render(256).save(png, optimize=True)
    print(f"Linux 图标：{png}")

    for out, sizes in ((os.path.join(HERE, "app.ico"), ICO_SIZES),
                       (os.path.join(HERE, "docs", "favicon.ico"), FAVICON_SIZES)):
        ims = [render(s) for s in sizes]
        # Pillow 存 ICO 时会拿最大那张自己缩出小帧（缩出来是糊的），
        # 所以用 append_images 显式给每一帧。
        ims[-1].save(out, format="ICO", sizes=[(s, s) for s in sizes],
                     append_images=ims[:-1])
        print(f"{out}：{len(sizes)} 个尺寸帧 {sizes}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
