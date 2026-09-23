# -*- coding: utf-8 -*-
"""知乎 / CSDN 推广头像生成器。

产出（全部在 promo/avatar/png/）：
  *-1024.png / *-400.png   各方案母版，可直接上传
  预览-尺寸阶梯.png          192/128/64/48/32/24 px 缩略对比（白底 + 深底）
  预览-场景模拟.png          知乎回答作者行 / CSDN 博客列表里的真实观感

改设计后重跑：python gen_avatar.py
约定同 gen_cards.py：不用 emoji，图形全部用 Pillow 绘制。
"""
import math
import os

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
APP = os.path.join(ROOT, "app.png")
PNG = os.path.join(HERE, "png")
os.makedirs(PNG, exist_ok=True)

S = 1024          # 母版边长
SS = 4            # 超采样倍数
N = S * SS        # 绘制画布边长

# 品牌色：实测取自 app.png 的渐变两端
G_TOP = (88, 177, 127)
G_BOT = (22, 108, 64)
WHITE = (255, 255, 255, 255)
DARKBG = (22, 24, 23)
INK = (26, 31, 28)
GRAY = (139, 148, 143)

FB = "C:/Windows/Fonts/msyhbd.ttc"
FR = "C:/Windows/Fonts/msyh.ttc"
LATB = "C:/Windows/Fonts/arialbd.ttf"

SAFE = 0.90       # 安全圆余量：保证平台裁圆不切到图形


# ---------------------------------------------------------------- 底版
def plate():
    g = Image.new("RGB", (1, 512))
    d = ImageDraw.Draw(g)
    for y in range(512):
        t = y / 511
        d.point((0, y), fill=tuple(int(G_TOP[i] + (G_BOT[i] - G_TOP[i]) * t) for i in range(3)))
    base = g.resize((N, N), Image.BILINEAR).convert("RGBA")

    hl = Image.new("L", (256, 256), 0)
    dh = ImageDraw.Draw(hl)
    for r in range(150, 0, -1):
        dh.ellipse([128 - r, 128 - r, 128 + r, 128 + r],
                   fill=int(255 * 0.13 * (1 - r / 150) ** 1.6))
    hl = hl.resize((N, N), Image.BILINEAR)
    base.alpha_composite(Image.composite(Image.new("RGBA", (N, N), WHITE),
                                         Image.new("RGBA", (N, N), (0, 0, 0, 0)), hl))
    return base


# ------------------------------------------------- 设计坐标 → 画布坐标
def make_k(bbox, margin=SAFE):
    """按外接圆内切反算缩放，保证内容整体落在安全圆内"""
    x0, y0, x1, y1 = bbox
    return (N / 2 * margin) / math.hypot((x1 - x0) / 2, (y1 - y0) / 2)


def P(k, x, y):
    return (x * k + N / 2, y * k + N / 2)


def RR(d, k, x0, y0, x1, y1, r, fill=WHITE):
    d.rounded_rectangle([P(k, x0, y0), P(k, x1, y1)], radius=r * k, fill=fill)


def route(d, k, pts, lw, fill=WHITE):
    """折线走线，joint=curve 让拐角变圆。端点务必藏在色块内部，否则会露圆头"""
    d.line([P(k, x, y) for x, y in pts], fill=fill,
           width=max(1, int(round(lw * k))), joint="curve")


def glyph(ch, font_path, px=1600):
    f = ImageFont.truetype(font_path, px)
    tmp = Image.new("RGBA", (px * 3, px * 3), (0, 0, 0, 0))
    ImageDraw.Draw(tmp).text((px, px), ch, font=f, fill=WHITE)
    return tmp.crop(tmp.getbbox())


def paste_tile(img, tile, height, cx, cy):
    """把已裁紧的文字图块按指定高度贴到画布像素位置 (N/2+cx, N/2+cy)"""
    w, h = tile.size
    nw, nh = max(1, round(w * height / h)), max(1, round(height))
    img.alpha_composite(tile.resize((nw, nh), Image.LANCZOS),
                        (round(N / 2 + cx - nw / 2), round(N / 2 + cy - nh / 2)))


# ------------------------------------------------------------- 方案 02
def m_flat3(img):
    """一表拆三份：把原图标去壳扁平化，元素整体放大"""
    bbox = (-700, -565, 700, 565)
    k = make_k(bbox)
    d = ImageDraw.Draw(img)
    lw = 118
    RR(d, k, -700, -250, -210, 250, 78)                     # 源表
    for yc in (-460, 0, 460):                               # 三份结果
        RR(d, k, 210, yc - 105, 700, yc + 105, 52)
    route(d, k, [(-300, 0), (-30, 0), (-30, -460), (330, -460)], lw)
    route(d, k, [(-300, 0), (-30, 0), (-30, 460), (330, 460)], lw)
    route(d, k, [(-300, 0), (330, 0)], lw)


# ------------------------------------------------------------- 方案 03
def m_flat2(img):
    """只留两块，元素最大，缩到 32px 最清"""
    bbox = (-700, -590, 760, 590)
    k = make_k(bbox)
    d = ImageDraw.Draw(img)
    lw = 175
    RR(d, k, -700, -340, -190, 340, 95)
    for yc in (-430, 430):
        RR(d, k, 200, yc - 160, 760, yc + 160, 70)
    route(d, k, [(-300, 0), (-10, 0), (-10, -430), (330, -430)], lw)
    route(d, k, [(-300, 0), (-10, 0), (-10, 430), (330, 430)], lw)


# ------------------------------------------------------------- 方案 04
def m_e3(img):
    """字母标：E = Excel，右列三格 = 拆多份"""
    tile = glyph("E", LATB)
    ar = tile.size[0] / tile.size[1]
    EH = 1000.0                       # E 的高度（设计单位）
    EW = EH * ar
    bar_w, bar_h, bar_gap = 430, 235, 175
    BH = 3 * bar_h + 2 * bar_gap      # 三格总高
    L = -880.0                        # 内容左边界
    BX = L + EW + 300                 # 三格左边界
    R = BX + bar_w
    YH = max(EH, BH) / 2
    bbox = (L, -YH, R, YH)
    k = make_k(bbox)
    d = ImageDraw.Draw(img)
    paste_tile(img, tile, EH * k, (L + EW / 2) * k, 0)
    for yc in (-(bar_h + bar_gap), 0, bar_h + bar_gap):
        RR(d, k, BX, yc - bar_h / 2, BX + bar_w, yc + bar_h / 2, 52)


# ------------------------------------------------------------- 方案 05
def m_char(img):
    """单字标「拆」：中文语境辨识度最高"""
    k = make_k((-500, -500, 500, 500))
    paste_tile(img, glyph("拆", FB), 1000 * k, 0, 0)


CANDIDATES = [
    ("02-扁平三拆", m_flat3, "把原图标的「一表拆三份」去壳扁平化，元素整体放大"),
    ("03-极简二拆", m_flat2, "只留两块，元素最大，缩到 32px 最清楚"),
    ("04-字母E+三格", m_e3, "字母标路线：E=Excel，右列三格=拆多份"),
    ("05-汉字拆", m_char, "单字标，中文语境辨识度最高、最好记"),
]


def build():
    items = []
    for name, fn, _d in CANDIDATES:
        im = plate()
        fn(im)
        out = im.convert("RGB").resize((S, S), Image.LANCZOS)
        out.save(os.path.join(PNG, f"{name}-1024.png"))
        out.resize((400, 400), Image.LANCZOS).save(os.path.join(PNG, f"{name}-400.png"))
        items.append((name, out))

    # 对照组：原图标直接上传 / 裁紧边距后上传
    o = Image.open(APP).convert("RGBA").resize((S, S), Image.LANCZOS)
    on_white = Image.new("RGB", (S, S), (255, 255, 255))
    on_white.paste(o, (0, 0), o)
    on_white.save(os.path.join(PNG, "01-对照-原图标直接上传-1024.png"))
    t = o.crop(o.getbbox()).resize((S, S), Image.LANCZOS)
    on_green = Image.new("RGB", (S, S), G_BOT)
    on_green.paste(t, (0, 0), t)
    on_green.save(os.path.join(PNG, "01b-对照-原图标裁紧边距-1024.png"))

    items.insert(0, ("01-原图标直接传", on_white))
    items.insert(1, ("01b-原图标裁紧边距", on_green))
    return items


# ----------------------------------------------------------- 预览：尺寸阶梯
SIZES = [128, 96, 64, 48, 32, 24]
ROW_H = 172


def circ(im, d):
    big = im.resize((d * 4, d * 4), Image.LANCZOS)
    m = Image.new("L", (d * 4, d * 4), 0)
    ImageDraw.Draw(m).ellipse([0, 0, d * 4 - 1, d * 4 - 1], fill=255)
    return big.resize((d, d), Image.LANCZOS), m.resize((d, d), Image.LANCZOS)


def col_x():
    """自左向右排布的列起点，间距固定，避免标签重叠"""
    xs, cur = [], 400
    for sz in SIZES:
        xs.append(cur)
        cur += sz + 38
    return xs


def sheet_ladder(items):
    xs = col_x()
    row_w = xs[-1] + SIZES[-1] + 40
    n = len(items)
    H = 70 + n * ROW_H + 66 + n * ROW_H + 36
    img = Image.new("RGB", (row_w, H), (255, 255, 255))
    d = ImageDraw.Draw(img)
    fb = ImageFont.truetype(FB, 30)
    f22 = ImageFont.truetype(FR, 22)

    y = 22
    d.text((40, y), "白底 · 日间模式", font=fb, fill=INK)
    y += 46
    for x, sz in zip(xs, SIZES):
        d.text((x + sz / 2, y), f"{sz}", font=f22, fill=GRAY, anchor="ma")
    y += 30
    for name, im in items:
        d.text((40, y + 74), name, font=fb, fill=INK)
        for x, sz in zip(xs, SIZES):
            c, m = circ(im, sz)
            img.paste(c, (x, y + 82 - sz // 2), m)
        y += ROW_H

    d.line([(0, y - 14), (row_w, y - 14)], fill=(228, 232, 229), width=2)
    y += 10
    d.text((40, y), "深底 · 夜间模式", font=fb, fill=INK)
    y += 46
    for x, sz in zip(xs, SIZES):
        d.text((x + sz / 2, y), f"{sz}", font=f22, fill=GRAY, anchor="ma")
    y += 30
    for name, im in items:
        d.text((40, y + 74), name, font=fb, fill=INK)
        for x, sz in zip(xs, SIZES):
            c, m = circ(im, sz)
            tile = Image.new("RGB", (sz, sz), DARKBG)
            tile.paste(c, (0, 0), m)
            img.paste(tile, (x, y + 82 - sz // 2))
        y += ROW_H

    p = os.path.join(PNG, "预览-尺寸阶梯.png")
    img.save(p)
    print("saved", p, img.size)


# ----------------------------------------------------------- 预览：场景模拟
def sheet_scene(items):
    pick = dict(items)
    cur, new, alt = pick["01-原图标直接传"], pick["03-极简二拆"], pick["05-汉字拆"]
    W, H = 1180, 950
    img = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(img)
    fb = ImageFont.truetype(FB, 32)
    f30 = ImageFont.truetype(FB, 30)
    f26 = ImageFont.truetype(FR, 26)
    f22 = ImageFont.truetype(FR, 22)
    f18 = ImageFont.truetype(FR, 18)

    def avatar(src, sz, x, y):
        c, m = circ(src, sz)
        img.paste(c, (x, y), m)

    d.text((40, 22), "放进真实信息流看一遍", font=fb, fill=INK)
    d.text((40, 66), "平台把方图裁成圆 —— 这才是用户实际看到的尺寸", font=f18, fill=GRAY)

    rows = [("直接复用原图标", cur), ("换成头像稿（极简二拆）", new)]

    y = 118
    d.text((40, y), "知乎 · 回答作者行（头像 36px）", font=f26, fill=GRAY)
    y += 40
    for i, (label, src) in enumerate(rows):
        yy = y + i * 112
        d.rounded_rectangle([40, yy - 10, W - 40, yy + 84], radius=14,
                            fill=(250, 250, 249), outline=(232, 236, 233), width=2)
        avatar(src, 36, 66, yy + 14)
        d.text((122, yy + 6), "头看AI｜Excel批量拆分", font=f30, fill=INK)
        d.text((122, yy + 50), "AI 自动化实战 · 让 AI 干真实的活 ｜ 免费开源 · 数据不出本机",
               font=f18, fill=(122, 132, 126))
        d.text((W - 250, yy + 30), label, font=f18, fill=GRAY)

    y += 236
    d.text((40, y), "CSDN · 博客列表作者位（头像 40px）", font=f26, fill=GRAY)
    y += 40
    for i, (label, src) in enumerate(rows):
        yy = y + i * 100
        d.rounded_rectangle([40, yy - 10, W - 40, yy + 76], radius=12,
                            fill=(250, 250, 249), outline=(232, 236, 233), width=2)
        avatar(src, 40, 62, yy + 8)
        d.text((124, yy + 0), "头看AI", font=f26, fill=INK)
        d.text((124, yy + 34), "Excel 批量拆分怎么按部门拆成多个文件", font=f22, fill=(92, 102, 96))
        d.text((W - 250, yy + 24), label, font=f18, fill=GRAY)

    x = 50
    yb = 762
    d.line([(40, yb - 34), (W - 40, yb - 34)], fill=(228, 232, 229), width=2)
    for label, src in [("原图标直接上传", cur), ("极简二拆", new), ("汉字「拆」", alt)]:
        avatar(src, 132, x, yb)
        d.text((x, yb + 146), label, font=f18, fill=GRAY)
        x += 290

    p = os.path.join(PNG, "预览-场景模拟.png")
    img.save(p)
    print("saved", p, img.size)


if __name__ == "__main__":
    it = build()
    sheet_ladder(it)
    sheet_scene(it)
    print("\n完成")
