# -*- coding: utf-8 -*-
"""生成 ExcelRouter 小红书 9 张配图（1080x1440 PNG），纯 Pillow 离线渲染。
改文案后重跑：python gen_cards.py
注意：不用 emoji 字符（雅黑无彩色 emoji，会渲染成方框），图标全部用图形绘制。
"""
import os
from PIL import Image, ImageDraw, ImageFont

W, H = 1080, 1440
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "png")
os.makedirs(OUT, exist_ok=True)

FR = "C:/Windows/Fonts/msyh.ttc"
FB = "C:/Windows/Fonts/msyhbd.ttc"

INK   = (26, 31, 28)
INK2  = (74, 83, 78)
GRAY  = (139, 148, 143)
LINE  = (227, 232, 228)
GREEN = (30, 127, 75)
GL    = (232, 243, 237)
GB    = (191, 224, 205)
GMID  = (61, 148, 96)
AMBER = (242, 153, 74)
AL    = (255, 248, 239)
BG    = (250, 249, 247)
RED   = (192, 57, 43)
RL    = (253, 236, 235)
WHITE = (255, 255, 255)

def f(sz, bold=False):
    return ImageFont.truetype(FB if bold else FR, sz)

PADX, PADY = 84, 88
CW = W - PADX * 2

# ---------------- 基础组件 ----------------
def new(bg=None):
    img = Image.new("RGB", (W, H), bg or BG)
    return img, ImageDraw.Draw(img)

def grad(img, c1, c2):
    d = ImageDraw.Draw(img)
    for y in range(H):
        t = y / H
        d.line([(0, y), (W, y)],
               fill=tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3)))

def rr(d, box, r, fill=None, outline=None, wd=2):
    d.rounded_rectangle(box, radius=r, fill=fill, outline=outline, width=wd)

def wrap(s, font, maxw):
    out = []
    for seg in s.split("\n"):
        cur = ""
        for ch in seg:
            if font.getlength(cur + ch) > maxw and cur:
                out.append(cur); cur = ch
            else:
                cur += ch
        out.append(cur)
    return out

def para(d, x, y, s, font, fill, lh, maxw=CW):
    for ln in wrap(s, font, maxw):
        d.text((x, y), ln, font=font, fill=fill, anchor="la")
        y += lh
    return y

def center(d, cx, y, s, font, fill, lh, maxw=CW):
    for ln in wrap(s, font, maxw):
        d.text((cx, y), ln, font=font, fill=fill, anchor="ma")
        y += lh
    return y

def kicker(d, y, s, color=GREEN):
    d.text((PADX, y), s, font=f(32, True), fill=color, anchor="la")
    return y + 66

def title(d, y, s, sz=82, color=INK, lh=None):
    return para(d, PADX, y, s, f(sz, True), color, lh or int(sz * 1.28))

def foot(d, s):
    y = H - PADY - 40
    d.line([(PADX, y - 36), (W - PADX, y - 36)], fill=LINE, width=3)
    d.text((PADX, y), s, font=f(30), fill=GRAY, anchor="la")

def callout(d, y, s, sz=30):
    lh = int(sz * 1.6)
    lines = wrap(s, f(sz, True), CW - 104)
    hgt = lh * len(lines) + 48
    rr(d, [PADX, y, W - PADX, y + hgt], 14, fill=AL)
    d.rounded_rectangle([PADX, y, PADX + 10, y + hgt], radius=5, fill=AMBER)
    yy = y + 24
    for ln in lines:
        d.text((PADX + 40, yy), ln, font=f(sz, True), fill=(138, 90, 32), anchor="la")
        yy += lh
    return y + hgt

def arrow_dn(d, cx, y, color=GREEN, sz=44, wd=7):
    d.line([(cx, y), (cx, y + sz * 0.68)], fill=color, width=wd)
    d.polygon([(cx - 16, y + sz * 0.62), (cx + 16, y + sz * 0.62), (cx, y + sz)], fill=color)
    return y + sz

def tri_dn(d, cx, cy, color, w=16, h=12):
    d.polygon([(cx - w, cy - h), (cx + w, cy - h), (cx, cy + h)], fill=color)

def chip(d, x, y, s, fg=INK2, bg=(243, 245, 244), sz=30, padx=20, pady=11):
    font = f(sz)
    wdt = int(font.getlength(s)) + padx * 2
    hgt = sz + pady * 2
    rr(d, [x, y, x + wdt, y + hgt], 12, fill=bg)
    d.text((x + padx, y + hgt // 2), s, font=font, fill=fg, anchor="lm")
    return x + wdt + 14, y + hgt + 14, hgt

def panel(d, y, hgt):
    rr(d, [PADX, y, W - PADX, y + hgt], 28, fill=WHITE, outline=LINE, wd=2)
    return y + 44

def step_badge(d, y, n, label, sz=38):
    rr(d, [PADX + 44, y, PADX + 102, y + 58], 29, fill=GREEN)
    d.text((PADX + 73, y + 29), n, font=f(32, True), fill=WHITE, anchor="mm")
    d.text((PADX + 124, y + 29), label, font=f(sz, True), fill=INK, anchor="lm")
    return y + 104

# ---------------- 卡片 1 · 封面 ----------------
def card1():
    img, d = new()
    grad(img, (242, 248, 244), (253, 245, 236))
    y = PADY
    rr(d, [PADX, y, PADX + 58, y + 58], 16, fill=GREEN)
    d.text((PADX + 29, y + 29), "ER", font=f(30, True), fill=WHITE, anchor="mm")
    d.text((PADX + 78, y + 29), "ExcelRouter", font=f(38, True), fill=INK, anchor="lm")
    y = kicker(d, y + 116, "EXCEL 批量拆分")

    y = title(d, y + 6, "一张总表\n先按部门\n再拆到每个人", 88)
    y += 40

    def box(yy, s, hl=False):
        hgt = 96
        rr(d, [PADX, yy, W - PADX, yy + hgt], 20,
           fill=GL if hl else WHITE, outline=GB if hl else LINE, wd=2)
        d.text((PADX + 34, yy + hgt // 2), s,
               font=f(38, True), fill=GREEN if hl else INK, anchor="lm")
        return yy + hgt + 10

    y = box(y, "销售明细.xlsx　300 人 · 8 个部门")
    d.text((PADX + 6, y + 4), "↓ 一级 · 拆分字段", font=f(30, True), fill=GREEN, anchor="la")
    y += 52
    y = box(y, "销售部 / 财务部 / 市场部 …", hl=True)
    d.text((PADX + 6, y + 4), "↓ 二级 · 同时拆到人", font=f(30, True), fill=GREEN, anchor="la")
    y += 52
    y = box(y, "张三月度.xlsx · 李四月度.xlsx …", hl=True)

    ty = 1104
    for row in (["一次跑完", "格式原样保留", "数据不出本机"],
                ["免费 · 开源", "Windows"]):
        x = PADX
        for s in row:
            x, _, ch = chip(d, x, ty, s, sz=30, padx=26, pady=11)
        ty += 72

    foot(d, "一次点选，两套结果同时到手")
    img.save(f"{OUT}/01-封面.png")

# ---------------- 卡片 2 · 痛点 ----------------
def card2():
    img, d = new()
    y = kicker(d, PADY, "先说最痛的两件事")
    y = title(d, y + 6, "发工资条、考核表\n最怕的是这两件", 76)
    y += 54

    ch = 380
    rr(d, [PADX, y, PADX + 440, y + ch], 26, fill=RL)
    rr(d, [PADX + 470, y, W - PADX, y + ch], 26, fill=GL)
    for bx, col, t, n, sub in [
        (PADX, RED, "以前", "2 小时", "复制 → 另存 → 改名\n下周还得重来一遍"),
        (PADX + 470, GREEN, "现在", "30 秒", "选表 → 选字段 → 开始\n表存着，下周直接跑")]:
        cx = bx + 220
        d.text((cx, y + 48), t, font=f(34, True), fill=col, anchor="ma")
        d.text((cx, y + 122), n, font=f(96, True), fill=col, anchor="ma")
        center(d, cx, y + 246, sub, f(30), INK2, 46, maxw=380)

    y = callout(d, y + ch + 56,
                "复制漏一行，A 的工资就进了 B 的邮箱。\n这种事一旦发生，解释不清。")
    foot(d, "别再手动拆了，真的会出事")
    img.save(f"{OUT}/02-痛点.png")

# ---------------- 卡片 3 · 一二级拆分（核心） ----------------
def card3():
    img, d = new()
    y = kicker(d, PADY, "核心差异 · 一二级一起拆")
    y = title(d, y + 6, "别的工具停在「按部门拆」\n它还能往下再拆一层", 60)
    y += 36

    def layer(yy, tag, main, chips, hl=False):
        chip_fg = WHITE if hl else INK2
        chip_bg = GMID if hl else (243, 245, 244)
        yy2 = yy + 126
        for c in chips:
            pass
        rows, cur, curw = [], [], 0
        widths = [(c, int(f(30).getlength(c)) + 40 + 14) for c in chips]
        for c, cw in widths:
            if curw + cw > CW - 76:
                rows.append(cur); cur, curw = [], 0
            cur.append(c); curw += cw
        if cur: rows.append(cur)
        hgt = 126 + (len(rows) * 64 if rows else 0) + 26
        rr(d, [PADX, yy, W - PADX, yy + hgt], 26,
           fill=GREEN if hl else WHITE, outline=GREEN if hl else LINE, wd=2)
        d.text((PADX + 38, yy + 34), tag,
               font=f(28, True), fill=(169, 220, 190) if hl else GRAY, anchor="la")
        d.text((PADX + 38, yy + 74), main,
               font=f(40, True), fill=WHITE if hl else INK, anchor="la")
        yy2 = yy + 126
        for row in rows:
            x = PADX + 38
            for c in row:
                x, _, chh = chip(d, x, yy2, c, fg=chip_fg, bg=chip_bg)
            yy2 += 64
        return yy + hgt + 8

    y = layer(y, "源 表", "5 个月 × 55 人，3 行复杂表头", [])
    y = arrow_dn(d, PADX + 30, y + 10) + 14
    y = layer(y, "一 级 · 拆分字段", "按「区域」拆", ["华北", "华东", "华南", "西南", "…"])
    y = arrow_dn(d, PADX + 30, y + 10) + 14
    y = layer(y, "二 级 · 同时拆到人", "每个区域里，再按「姓名」拆",
              ["张三.xlsx", "李四.xlsx", "王五.xlsx", "…"], hl=True)

    callout(d, y + 36, "一次运行，汇总 + 每人一份同时产出，不用跑两遍。")
    foot(d, "还可以只对文件名含「工资」的表拆到人")
    img.save(f"{OUT}/03-一二级拆分.png")

# ---------------- 卡片 4 · 操作① ----------------
def card4():
    img, d = new()
    y = kicker(d, PADY, "实操 · 第 1 步")
    y = title(d, y + 6, "① 选表格", 82)
    y = panel(d, y + 40, 460)
    y = step_badge(d, y, "1", "选表格")

    for i, (s, on) in enumerate([("选一个 Excel 文件", False), ("选整个文件夹（批量拆）", True)]):
        bx = PADX + 44 + i * 452
        rr(d, [bx, y, bx + 428, y + 130], 18,
           fill=GL if on else (251, 252, 250),
           outline=GREEN if on else (195, 205, 199), wd=2)
        d.text((bx + 214, y + 65), s,
               font=f(33, True), fill=GREEN if on else INK2, anchor="mm")
    y += 168

    para(d, PADX + 44, y,
         "· 文件或文件夹可以直接拖进窗口，也能粘贴路径\n"
         "· 选整个文件夹才会出现「打包 ZIP / 拆到人」的选项",
         f(30), GRAY, 46, maxw=CW - 88)

    y = PADY + 66 + 6 + 105 + 40 + 460 + 48
    for t, s in [("整文件夹丢进去", "十几张表一次处理，不用一张张点"),
                 ("自动识别表头", "3 行表头、合并单元格也能认出真正的表头行")]:
        rr(d, [PADX, y, W - PADX, y + 168], 22, fill=WHITE, outline=LINE, wd=2)
        rr(d, [PADX + 44, y + 54, PADX + 84, y + 94], 8, fill=GREEN)
        d.text((PADX + 116, y + 46), t, font=f(40, True), fill=INK, anchor="la")
        d.text((PADX + 116, y + 106), s, font=f(30), fill=INK2, anchor="la")
        y += 190

    foot(d, "界面从上到下就三步，不用看教程")
    img.save(f"{OUT}/04-操作1.png")

# ---------------- 卡片 5 · 操作② 一级 ----------------
def card5():
    img, d = new()
    y = kicker(d, PADY, "实操 · 第 2 步")
    y = title(d, y + 6, "② 选字段", 82)
    d.text((PADX + 340, y - 78), "（一级）", font=f(38), fill=INK2, anchor="la")
    y = panel(d, y + 40, 560)
    y = step_badge(d, y, "2", "选字段")

    d.text((PADX + 44, y + 48), "拆分字段", font=f(34, True), fill=INK2, anchor="lm")
    rr(d, [PADX + 254, y, W - PADX - 44, y + 96], 14, fill=GL, outline=GB, wd=2)
    d.text((PADX + 292, y + 48), "区域", font=f(38, True), fill=GREEN, anchor="lm")
    tri_dn(d, W - PADX - 84, y + 48, (139, 148, 143))
    y += 138

    rr(d, [PADX + 44, y, PADX + 344, y + 84], 14, fill=GL)
    d.text((PADX + 84, y + 42), "将拆成 5 组", font=f(34, True), fill=GREEN, anchor="lm")
    y += 126

    para(d, PADX + 44, y,
         "· 点开始前就知道会拆出什么，不满意就换一列\n"
         "· 只想拆其中几个分组 →「查看/勾选分组」勾掉不要的\n"
         "· 下拉是空的？点「重新识别」",
         f(30), GRAY, 46, maxw=CW - 88)

    callout(d, y + 190, "如果预览显示几十上百组，多半选到工号/姓名这类列了。")
    foot(d, "先预览，再开跑 —— 不盲拆")
    img.save(f"{OUT}/05-操作2一级.png")

# ---------------- 卡片 6 · 操作② 二级（最重要） ----------------
def card6():
    img, d = new()
    y = kicker(d, PADY, "核心差异 · 第 2 步的隐藏关卡")
    y = title(d, y + 6, "勾上「同时拆到人」\n才算是真用上它了", 60)
    y = panel(d, y + 36, 718)
    y = step_badge(d, y, "2", "选字段 · 批量选项") - 8

    def chk(yy, t, s, on):
        rr(d, [PADX + 44, yy, PADX + 92, yy + 48], 10,
           fill=GREEN if on else WHITE, outline=GREEN if on else (195, 205, 199), wd=3)
        if on:
            d.line([(PADX + 56, yy + 25), (PADX + 66, yy + 35), (PADX + 82, yy + 15)],
                   fill=WHITE, width=5, joint="curve")
        d.text((PADX + 120, yy), t, font=f(36, True),
               fill=GREEN if on else INK, anchor="la")
        d.text((PADX + 120, yy + 56), s, font=f(27), fill=GRAY, anchor="la")
        return yy + 112

    y = chk(y, "每个分组打包成 ZIP", "方便把「华北」整包直接转发", False)
    y = chk(y, "同时拆到人", "在分组之外，再把每个人单独出一个文件", True)
    y += 12

    rr(d, [PADX + 108, y, W - PADX - 44, y + 190], 16, fill=(246, 248, 247))
    for i, (lb, val) in enumerate([("按哪个字段区分人", "姓名"), ("只处理文件名含", "工资")]):
        yy = y + 22 + i * 86
        d.text((PADX + 148, yy + 28), lb, font=f(28), fill=INK2, anchor="lm")
        rr(d, [PADX + 470, yy, PADX + 770, yy + 62], 10, fill=GL, outline=GB, wd=2)
        d.text((PADX + 620, yy + 31), val, font=f(32, True), fill=GREEN, anchor="mm")
    y += 216

    para(d, PADX + 44, y,
         "· 留空 ＝ 全部表都拆到人\n"
         "· 填「工资」＝ 只有工资表拆到人，其他表照常按部门拆",
         f(29), GRAY, 44, maxw=CW - 88)

    callout(d, y + 118, "同一批表，有的拆到部门、有的拆到人 —— 一次搞定。")
    foot(d, "这一步，是它和普通拆分工具的分水岭")
    img.save(f"{OUT}/06-操作2二级.png")

# ---------------- 卡片 7 · 结果 ----------------
def card7():
    img, d = new()
    y = kicker(d, PADY, "跑完长这样")
    y = title(d, y + 6, "300 多个文件\n自己躺在文件夹里", 68)
    y = panel(d, y + 44, 790)

    ly = y + 8
    d.text((PADX + 34, ly), "拆分结果 · 按区域分组，每组含汇总 + 到人",
           font=f(26, True), fill=GRAY, anchor="la")
    ly += 56
    d.line([(PADX, ly), (W - PADX, ly)], fill=LINE, width=2)
    ly += 6

    rows = [
        ("华北", "", 1), ("华北_汇总.xlsx", "312 KB", 2), ("到人", "11 项", 1),
        ("张三月度.xlsx", "28 KB", 3), ("李四月度.xlsx", "28 KB", 3),
        ("王五月度.xlsx", "27 KB", 3), ("华东", "", 1),
        ("华东_汇总.xlsx", "298 KB", 2), ("华东.zip", "1.4 MB", 2),
        ("华南 / 西南 / …", "3 项", 1),
    ]
    for nm, sz, lvl in rows:
        x = PADX + 34 + (0 if lvl == 1 else 56)
        iy = ly + 30
        if lvl <= 1 or nm == "到人":
            rr(d, [x, iy - 18, x + 44, iy + 16], 8, fill=GREEN)
            rr(d, [x, iy - 12, x + 44, iy - 3], 4, fill=WHITE)
        elif nm.endswith(".zip"):
            rr(d, [x + 6, iy - 17, x + 38, iy + 17], 6, fill=(160, 170, 165))
            d.text((x + 22, iy + 1), "Z", font=f(20, True), fill=WHITE, anchor="mm")
        else:
            rr(d, [x + 6, iy - 17, x + 38, iy + 17], 6, fill=GREEN)
            d.text((x + 22, iy + 1), "X", font=f(20, True), fill=WHITE, anchor="mm")
        col = GREEN if lvl == 3 else INK
        d.text((x + 60, iy), nm, font=f(32, True) if lvl <= 2 else f(32), fill=col, anchor="lm")
        if sz:
            d.text((W - PADX - 34, iy), sz, font=f(26), fill=(164, 172, 167), anchor="rm")
        d.line([(PADX + 20, ly + 64), (W - PADX - 20, ly + 64)], fill=(240, 242, 241), width=1)
        ly += 66

    foot(d, "汇总 + 到人 + ZIP，一次全给")
    img.save(f"{OUT}/07-结果.png")

# ---------------- 卡片 8 · 卖点 ----------------
def card8():
    img, d = new()
    y = kicker(d, PADY, "除了拆得快")
    y = title(d, y + 6, "拆完不像被机器\n压成一坨的数据", 60)
    y += 28

    items = [
        ("格式原样保留", "复杂表头、合并单元格、颜色、列宽全保留，\n跟原表一模一样，不用二次排版"),
        ("跨文件自动合并", "5 个月的表，同一个人的数据自动并到一起"),
        ("数据不出本机", "不联网、不上传，断网也能跑。\n工资、客户资料我是不敢往网站传的"),
        ("免费 · 开源（MIT）", "无广告、无注册、无功能限制，想看源码就看"),
        ("顺带还能发 PDF", "同一份 PDF，给每个网格生成\n专属密码 + 专属水印的副本"),
    ]
    for t, s in items:
        n = len(s.split("\n"))
        hgt = 76 + 38 * n + 28
        rr(d, [PADX, y, W - PADX, y + hgt], 22, fill=WHITE, outline=LINE, wd=2)
        rr(d, [PADX + 46, y + 36, PADX + 86, y + 76], 20, fill=GL)
        rr(d, [PADX + 56, y + 46, PADX + 76, y + 66], 10, fill=GREEN)
        d.text((PADX + 122, y + 30), t, font=f(36, True), fill=INK, anchor="la")
        para(d, PADX + 122, y + 82, s, f(27), INK2, 38, maxw=CW - 200)
        y += hgt + 14

    foot(d, "Windows 桌面版 · 免安装 · 零编程")
    img.save(f"{OUT}/08-卖点.png")

# ---------------- 卡片 9 · 结尾 ----------------
def card9():
    img, d = new()
    grad(img, (242, 248, 244), (250, 249, 247))
    y = kicker(d, PADY, "如果你也每周在做这些")

    items = [
        "工资条 / 考核表 → 一人一份，谁也看不到别人的",
        "销售结算明细 → 按区域拆，发给各负责人",
        "门店 / 网格台账 → 按编号拆 + 打包群发",
        "学校 / 培训机构 → 按班级拆，再拆到每个学生",
    ]
    rr(d, [PADX, y, W - PADX, y + 470], 32, fill=WHITE, outline=LINE, wd=3)
    yy = y + 56
    for s in items:
        rr(d, [PADX + 56, yy + 14, PADX + 76, yy + 34], 10, fill=GREEN)
        d.text((PADX + 104, yy), s, font=f(32), fill=INK, anchor="la")
        yy += 104

    y = y + 470 + 96
    y = arrow_dn(d, W // 2, y, GREEN, 96, 10) + 20
    y = center(d, W // 2, y, "想要的自取", f(80, True), GREEN, 110)
    y = center(d, W // 2, y + 20, "看评论区置顶", f(56, True), INK, 80)
    center(d, W // 2, y + 26, "免费 · 开源 · Windows 直接能跑", f(34), INK2, 50)

    foot(d, "以前两小时的活，现在泡杯茶的功夫")
    img.save(f"{OUT}/09-结尾.png")

if __name__ == "__main__":
    for fn in (card1, card2, card3, card4, card5, card6, card7, card8, card9):
        fn()
        print("ok", fn.__name__)
    print("输出目录:", OUT)
