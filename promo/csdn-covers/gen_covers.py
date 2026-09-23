#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_covers.py —— 生成 CSDN 文章封面图（第 2~6 篇，共 5 张）。

尺寸：1140 × 640（16:9，CSDN 博客封面推荐比例）
风格：统一深色技术风 —— 左侧品牌色竖条 + 暗底细网格 + 大号标题 + 右上淡数字
      所有文章共用同一套版式，只换文案，保证风格一致。

用法：
  python gen_covers.py            # 生成全部 5 张 → png/
  python gen_covers.py --list     # 只打印文案清单，不出图

改文案：直接改下面 ARTICLES 里的字段，重跑一遍即可。
注意：第 4 篇是「不涉及产品」的纯技术篇，页脚**不带产品名**，别顺手改成统一文案。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

# ----------------------------------------------------------------- 画布与配色
W, H = 1140, 640
BG = (14, 20, 32)          # 暗底
GRID = (23, 32, 47)        # 网格线（比底色略亮）
ACCENT = (47, 204, 113)    # 品牌绿（Excel 系）
INK = (241, 246, 255)      # 主文字
INK_L2 = (205, 217, 236)   # 标题第二行（比次文字亮，保证缩略图可读）
INK2 = (150, 164, 186)     # 次文字
DIVIDER = (34, 46, 66)

FONT_BD = "C:/Windows/Fonts/msyhbd.ttc"   # 微软雅黑 Bold
FONT_RG = "C:/Windows/Fonts/msyh.ttc"     # 微软雅黑 Regular

PADX = 84
OUT_DIR = Path(__file__).resolve().parent / "png"

# ----------------------------------------------------------------- 每篇文案
ARTICLES = [
    dict(
        no="02",
        tag="Excel · 办公自动化",
        l1="工资表批量拆分成单人文件",
        l2="邮件合并 / 脚本 / 工具三条路",
        sub="三种方案实测对比 · 保留格式 · 数据不出本机",
        foot="ExcelRouter · 开源 MIT",
        name="02-工资表拆分成单人文件",
    ),
    dict(
        no="03",
        tag="AI · 数据安全",
        l1="用 AI 处理敏感数据前",
        l2="先分清这三层的边界",
        sub="模型层 / 执行层 / 存储层，风险到底在哪",
        foot="ExcelRouter · 本地运行",
        name="03-AI处理敏感数据的三层边界",
    ),
    dict(
        no="04",
        tag="Python · 代码质量",
        l1="AI 生成代码的 4 个高危点",
        l2="路径 · 删除 · 异常 · 循环",
        sub="每一条都附修复代码，来自真实踩坑",
        foot="Python · 代码质量",          # ← 纯技术篇，不带产品名
        name="04-AI生成代码的4个高危点",
    ),
    dict(
        no="05",
        tag="Python · Excel",
        l1="Python 拆分 Excel 保留原格式",
        l2="表头与合并单元格怎么处理",
        sub="openpyxl 逐单元格复制样式 · 完整代码",
        foot="ExcelRouter · 保留格式",
        name="05-Python拆分Excel保留原格式",
    ),
    dict(
        no="06",
        tag="AI · 技能封装",
        l1="把内部工具封装成 AI 技能",
        l2="一次让 AI 真干活的实践",
        sub="从「生成内容」到「执行任务」的第三步",
        foot="ExcelRouter · 双入口同内核",
        name="06-把内部工具封装成AI技能",
    ),
]


def f(size: int, bold: bool = True) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_BD if bold else FONT_RG, size)


def fit(draw: ImageDraw.ImageDraw, text: str, maxw: float,
        base: int, bold: bool = True) -> ImageFont.FreeTypeFont:
    """在不超过 maxw 的前提下，尽量用 base 字号；放不下才逐级缩小。"""
    size = base
    while size > 16:
        ft = f(size, bold)
        if draw.textlength(text, font=ft) <= maxw:
            return ft
        size -= 2
    return f(16, bold)


def draw_grid(d: ImageDraw.ImageDraw) -> None:
    step = 57
    for x in range(0, W, step):
        d.line([(x, 0), (x, H)], fill=GRID, width=1)
    for y in range(0, H, step):
        d.line([(0, y), (W, y)], fill=GRID, width=1)


def render(a: dict) -> Path:
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)

    draw_grid(d)

    # 左侧品牌色竖条
    d.rectangle([0, 0, 11, H], fill=ACCENT)

    maxw = W - PADX * 2

    # ---- 分类标签（描边胶囊）
    tag_f = f(26)
    tw = d.textlength(a["tag"], font=tag_f)
    pill = [PADX, 72, PADX + tw + 44, 122]
    d.rounded_rectangle(pill, radius=25, outline=ACCENT, width=2)
    d.text((pill[0] + 22, 97), a["tag"], font=tag_f, fill=ACCENT, anchor="lm")

    # ---- 主标题两行（字号固定，保证 5 张视觉一致）
    l1_f = fit(d, a["l1"], maxw, 72)
    l2_f = fit(d, a["l2"], maxw, 54)
    d.text((PADX, 166), a["l1"], font=l1_f, fill=INK, anchor="la")
    d.text((PADX, 262), a["l2"], font=l2_f, fill=INK_L2, anchor="la")

    # ---- 副标题
    sub_f = fit(d, a["sub"], maxw, 30, bold=False)
    d.text((PADX, 378), a["sub"], font=sub_f, fill=INK2, anchor="la")

    # ---- 页脚分隔线
    d.line([(PADX, 500), (W - PADX, 500)], fill=DIVIDER, width=2)

    # 左下：产品/主题标识（带品牌色小方块）
    d.rectangle([PADX, 536, PADX + 14, 550], fill=ACCENT)
    d.text((PADX + 28, 543), a["foot"], font=f(26, False), fill=INK2, anchor="lm")

    # 右下：篇序
    d.text((W - PADX, 543), f"{a['no']} / 06", font=f(26, True),
           fill=(90, 104, 126), anchor="rm")

    # ---- 右上角淡数字（用 RGBA 叠加，避免压暗网格）
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.text((W - 58, 96), a["no"], font=f(280, True),
            fill=(255, 255, 255, 17), anchor="ra")
    img = Image.alpha_composite(img.convert("RGBA"), ov).convert("RGB")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"{a['name']}.png"
    img.save(path, "PNG")
    return path


def main() -> int:
    p = argparse.ArgumentParser(description="生成 CSDN 封面图")
    p.add_argument("--list", action="store_true", help="只打印文案，不出图")
    args = p.parse_args()

    if args.list:
        for a in ARTICLES:
            print(f"[{a['no']}] {a['l1']} / {a['l2']}\n      {a['sub']}\n      {a['foot']}")
        return 0

    for a in ARTICLES:
        print(f"✓ {render(a)}")
    print(f"\n共 {len(ARTICLES)} 张，尺寸 {W}×{H}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
