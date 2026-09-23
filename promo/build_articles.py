# -*- coding: utf-8 -*-
# Copyright (c) 2026 AbeLin · MIT License
"""把 promo/ 里的成品推广稿生成为 docs/articles/ 下的静态页面（GitHub Pages 用）。

为什么只取 CSDN 版：6 个选题各有知乎版与 CSDN 版两稿，内容高度重合。同一个站内
挂两份近似正文属于**站内重复内容**，搜索引擎会让它们互相压制。CSDN 版更长、带代码与
对比表，因此选它作为自有站上的**唯一规范版本**；知乎版仍留在 promo/ 供去知乎回答时用
（平台发布时按 SOP 改开头第一句，与本站这份区分开）。

**只发文章正文**：promo/ 里的推广总纲、内容排期、题库属于对内投放策略（含账号名、
选题纪律、下载量基线），公开出去会直接毁掉文章的可信度，不进本站。

用法：py promo/build_articles.py
"""
import html
import io
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "promo", "成品-CSDN6篇-直接复制.md")
OUT = os.path.join(ROOT, "docs", "articles")

GH = "https://github.com/MarsandSea/ExcelRouter"
GT = "https://gitee.com/Marsandsea/Excelrouter"
SITE = "https://marsandsea.github.io/ExcelRouter"

# 每篇的 slug 与列表页摘要（人工指定，比自动截断可读）
META = {
    1: ("excel-split-by-column",
        "按列拆分的三条路：Power Query、Python 脚本、专用工具，含耗时与格式保留对比。"),
    2: ("payroll-split-per-person",
        "工资表拆成一人一份：邮件合并、脚本生成、本地批量工具，三种做法的边界在哪。"),
    3: ("ai-data-boundary",
        "把表格交给 AI 处理前，先分清模型层、执行层、存储层——怎么判断它到底联不联网。"),
    4: ("ai-code-risks",
        "AI 生成代码的 4 个高危点：路径写死、无预览的破坏性操作、静默吞异常、无边界循环。"),
    5: ("python-split-excel-keep-format",
        "用 openpyxl 拆分 Excel 还保住原格式：表头、合并单元格、行号错位的最小可用实现。"),
    6: ("tool-as-ai-skill",
        "把一个内部工具封装成 AI 可调用的技能，让助手真的去干活而不只是给建议。"),
}

CSS = """:root{
  --bg:#ffffff; --panel:#f7f8fa; --border:#e4e7eb;
  --ink:#15181c; --ink2:#454c54; --ink3:#79828b;
  --brand:#0E7FD1; --brand-d:#0b64a6;
  --codebg:#f9fafb;
}
*{box-sizing:border-box}
body{margin:0;padding:0;background:var(--bg);color:var(--ink);
  font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Hiragino Sans GB","Microsoft YaHei",sans-serif;
  font-size:16.5px;line-height:1.85;-webkit-font-smoothing:antialiased}
.wrap{max-width:760px;margin:0 auto;padding:0 20px}
a{color:var(--brand);text-decoration:none}
a:hover{text-decoration:underline}
.top{border-bottom:1px solid var(--border);padding:14px 0;font-size:14px;color:var(--ink3)}
.top .wrap{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.top b{color:var(--ink);font-weight:650}
article{padding:34px 0 10px}
h1{font-size:29px;line-height:1.4;margin:0 0 10px;letter-spacing:-.02em;font-weight:750}
.byline{font-size:13.5px;color:var(--ink3);margin:0 0 30px;padding-bottom:18px;border-bottom:1px solid var(--border)}
h2{font-size:21px;margin:40px 0 12px;letter-spacing:-.01em}
h3{font-size:17.5px;margin:28px 0 8px}
p{margin:0 0 16px;color:var(--ink2)}
strong{color:var(--ink);font-weight:650}
ul,ol{margin:0 0 18px;padding-left:24px;color:var(--ink2)}
li{margin:6px 0}
code{background:var(--panel);border:1px solid var(--border);border-radius:4px;padding:1px 6px;
  font-size:14px;font-family:ui-monospace,SFMono-Regular,Consolas,monospace;color:#b03060}
pre{background:var(--codebg);border:1px solid var(--border);border-left:3px solid var(--brand);
  border-radius:7px;padding:15px 17px;overflow-x:auto;font-size:13.5px;line-height:1.65;margin:0 0 18px;
  font-family:ui-monospace,SFMono-Regular,Consolas,monospace;color:#24292f}
pre code{background:none;border:none;padding:0;color:inherit;font-size:inherit}
blockquote{margin:0 0 18px;padding:12px 18px;border-left:3px solid var(--border);
  background:var(--panel);border-radius:0 7px 7px 0;color:var(--ink2);font-size:15px}
blockquote p:last-child{margin:0}
table{width:100%;border-collapse:collapse;margin:0 0 20px;font-size:14.5px}
th{text-align:left;background:var(--panel);padding:9px 12px;border:1px solid var(--border);
  font-weight:650;font-size:13.5px;color:var(--ink2)}
td{padding:9px 12px;border:1px solid var(--border);vertical-align:top;color:var(--ink2)}
.cta{background:var(--panel);border:1px solid var(--border);border-radius:10px;
  padding:20px 22px;margin:42px 0 0}
.cta h3{margin:0 0 8px;font-size:17px}
.cta p{margin:0 0 14px;font-size:15px}
.btn{display:inline-block;padding:11px 22px;border-radius:8px;font-weight:650;font-size:15px;
  background:var(--brand);color:#fff;border:1px solid var(--brand);margin-right:10px}
.btn:hover{background:var(--brand-d);text-decoration:none}
.btn.ghost{background:#fff;color:var(--ink);border-color:var(--border)}
.cards{list-style:none;padding:0;margin:0 0 20px}
.cards li{border:1px solid var(--border);border-radius:9px;padding:16px 18px;margin-bottom:12px;background:#fff}
.cards a{font-size:17.5px;font-weight:650;display:block;margin-bottom:5px}
.cards p{margin:0;font-size:14.5px}
footer{margin-top:56px;padding:24px 0 52px;border-top:1px solid var(--border);
  font-size:13.5px;color:var(--ink3);text-align:center}
footer a{color:var(--ink2)}
@media(max-width:700px){h1{font-size:24px}body{font-size:16px}}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --bg:#14171a; --panel:#1c2126; --border:#2b3238; --codebg:#1a1f24;
    --ink:#eef1f4; --ink2:#b6bfc7; --ink3:#8b959e;
    --brand:#38BDF8; --brand-d:#7dd3fc;
  }
  :root:not([data-theme="light"]) pre{color:#d6dde4}
  :root:not([data-theme="light"]) .btn{color:#06202e}
  :root:not([data-theme="light"]) .btn.ghost,
  :root:not([data-theme="light"]) .cards li{background:var(--panel);color:var(--ink)}
}
"""


def esc(s):
    return html.escape(s, quote=False)


def inline(s):
    """行内标记：先转义，再还原 **粗体** / `代码` / 裸链接。"""
    s = esc(s)
    s = re.sub(r"`([^`]+)`", r"<code>\1</code>", s)
    s = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"(?<!\")(https?://[^\s<>）)，。]+)", r'<a href="\1">\1</a>', s)
    return s


def md2html(md):
    """够用的 Markdown 子集：标题/围栏代码/表格/有序无序列表/引用/段落。"""
    out, lines, i = [], md.split("\n"), 0
    while i < len(lines):
        ln = lines[i]

        if ln.startswith("```"):                      # 围栏代码块
            i += 1
            buf = []
            while i < len(lines) and not lines[i].startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1
            out.append("<pre><code>%s</code></pre>" % esc("\n".join(buf)))
            continue

        m = re.match(r"^(#{2,4})\s+(.*)$", ln)        # 标题（源文件最高只到 ##）
        if m:
            lvl = min(len(m.group(1)), 3)
            out.append("<h%d>%s</h%d>" % (lvl, inline(m.group(2)), lvl))
            i += 1
            continue

        if ln.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s:|-]+\|$", lines[i + 1]):
            head = [c.strip() for c in ln.strip("|").split("|")]
            i += 2
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                rows.append([c.strip() for c in lines[i].strip("|").split("|")])
                i += 1
            out.append("<table><tr>%s</tr>%s</table>" % (
                "".join("<th>%s</th>" % inline(c) for c in head),
                "".join("<tr>%s</tr>" % "".join("<td>%s</td>" % inline(c) for c in r) for r in rows)))
            continue

        if re.match(r"^\s*[-*]\s+", ln):              # 无序列表
            items = []
            while i < len(lines) and re.match(r"^\s*[-*]\s+", lines[i]):
                items.append(re.sub(r"^\s*[-*]\s+", "", lines[i]))
                i += 1
            out.append("<ul>%s</ul>" % "".join("<li>%s</li>" % inline(x) for x in items))
            continue

        if re.match(r"^\s*\d+[.)]\s+", ln):           # 有序列表
            items = []
            while i < len(lines) and re.match(r"^\s*\d+[.)]\s+", lines[i]):
                items.append(re.sub(r"^\s*\d+[.)]\s+", "", lines[i]))
                i += 1
            out.append("<ol>%s</ol>" % "".join("<li>%s</li>" % inline(x) for x in items))
            continue

        if ln.startswith(">"):                        # 引用
            items = []
            while i < len(lines) and lines[i].startswith(">"):
                items.append(lines[i].lstrip(">").strip())
                i += 1
            out.append("<blockquote>%s</blockquote>"
                       % "".join("<p>%s</p>" % inline(x) for x in items if x))
            continue

        if ln.strip() == "":
            i += 1
            continue

        para = []                                     # 段落（连续行以 <br> 连接）
        while i < len(lines) and lines[i].strip() and not re.match(
                r"^(#{2,4}\s|```|\||>|\s*[-*]\s|\s*\d+[.)]\s)", lines[i]):
            para.append(lines[i].strip())
            i += 1
        out.append("<p>%s</p>" % "<br>".join(inline(x) for x in para))
    return "\n".join(out)


def parse(src):
    """切出 6 篇：标题 + ▼正文开始/▲正文结束 之间的正文。"""
    txt = io.open(src, encoding="utf-8").read()
    arts = []
    for m in re.finditer(
            r"(?ms)^# 第 (\d) 篇.*?^## 【标题】\s*\n+(.+?)\n.*?^▼ 正文开始\s*\n(.*?)^▲ 正文结束",
            txt):
        arts.append((int(m.group(1)), m.group(2).strip(), m.group(3).strip()))
    return arts


PAGE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title} — ExcelRouter</title>
<meta name="description" content="{desc}">
<link rel="canonical" href="{site}/articles/{slug}.html">
<meta property="og:type" content="article">
<meta property="og:title" content="{title}">
<meta property="og:description" content="{desc}">
<meta property="og:url" content="{site}/articles/{slug}.html">
<link rel="icon" href="../logo.png">
<link rel="stylesheet" href="article.css">
<script type="application/ld+json">
{{"@context":"https://schema.org","@type":"TechArticle",
 "headline":{title_json},"description":{desc_json},
 "mainEntityOfPage":"{site}/articles/{slug}.html",
 "author":{{"@type":"Person","name":"AbeLin"}},
 "publisher":{{"@type":"Person","name":"AbeLin"}},
 "inLanguage":"zh-CN"}}
</script>
</head>
<body>

<nav class="top"><div class="wrap">
  <a href="../">← ExcelRouter</a><span>·</span><a href="./"><b>实用文章</b></a>
</div></nav>

<div class="wrap">
<article>
  <h1>{title}</h1>
  <p class="byline">作者 AbeLin · ExcelRouter 是一个免费开源（MIT）的 Excel 批量拆分工具</p>
{body}

  <div class="cta">
    <h3>文中提到的工具：ExcelRouter</h3>
    <p>把一个或一整批 Excel 按部门、区域、工号等字段拆成多个文件，<strong>完整保留复杂表头与格式</strong>，
       可再按人二级拆分并打包 ZIP。Windows + 银河麒麟 / 统信 UOS，免费开源（MIT），
       <strong>全程本机处理，不联网、不上传</strong>。</p>
    <a class="btn" href="{gt}/releases">下载（国内 · Gitee）</a>
    <a class="btn ghost" href="{gh}">GitHub 源码 ⭐</a>
  </div>
</article>
</div>

<footer>
  <p><a href="../">ExcelRouter 首页</a> · <a href="./">更多文章</a> ·
     <a href="{gh}">GitHub</a> · <a href="{gt}">Gitee</a></p>
  <p>作者 AbeLin · MIT 开源 · 数据仅在本机处理</p>
</footer>

</body>
</html>
"""

INDEX = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Excel 拆分实用文章 — ExcelRouter</title>
<meta name="description" content="围绕 Excel 批量拆分的实用文章：按列/按部门拆分的几种做法对比、工资条一人一份怎么批量做、拆分后怎么保住原格式、把表格交给 AI 处理的数据边界。">
<link rel="canonical" href="{site}/articles/">
<link rel="icon" href="../logo.png">
<link rel="stylesheet" href="article.css">
</head>
<body>

<nav class="top"><div class="wrap">
  <a href="../">← ExcelRouter</a><span>·</span><b>实用文章</b>
</div></nav>

<div class="wrap">
<article>
  <h1>实用文章</h1>
  <p class="byline">Excel 批量拆分的几种做法、踩过的坑，以及把表格交给 AI 处理时的边界</p>

  <ul class="cards">
{items}
  </ul>

  <div class="cta">
    <h3>ExcelRouter</h3>
    <p>整个文件夹一键拆完：按部门、区域、工号等字段自动拆分，保留原格式，打包分发。
       免费开源（MIT），数据仅在本机处理。</p>
    <a class="btn" href="../">看看它能做什么</a>
    <a class="btn ghost" href="{gt}/releases">直接下载</a>
  </div>
</article>
</div>

<footer>
  <p><a href="../">ExcelRouter 首页</a> · <a href="{gh}">GitHub</a> · <a href="{gt}">Gitee</a></p>
  <p>作者 AbeLin · MIT 开源 · 数据仅在本机处理</p>
</footer>

</body>
</html>
"""


def main():
    import json
    if not os.path.isdir(OUT):
        os.makedirs(OUT)
    io.open(os.path.join(OUT, "article.css"), "w", encoding="utf-8", newline="\n").write(CSS)

    arts = parse(SRC)
    if len(arts) != 6:
        raise SystemExit("解析到 %d 篇，期望 6 篇——源文件结构可能变了" % len(arts))

    items = []
    for n, title, body in arts:
        slug, desc = META[n]
        page = PAGE.format(
            title=esc(title), desc=esc(desc), slug=slug, site=SITE, gh=GH, gt=GT,
            title_json=json.dumps(title, ensure_ascii=False),
            desc_json=json.dumps(desc, ensure_ascii=False),
            body=md2html(body))
        io.open(os.path.join(OUT, slug + ".html"), "w", encoding="utf-8", newline="\n").write(page)
        items.append('    <li><a href="%s.html">%s</a><p>%s</p></li>' % (slug, esc(title), esc(desc)))
        print("生成  articles/%s.html   %s" % (slug, title))

    io.open(os.path.join(OUT, "index.html"), "w", encoding="utf-8", newline="\n").write(
        INDEX.format(items="\n".join(items), site=SITE, gh=GH, gt=GT))
    print("生成  articles/index.html（%d 篇）" % len(arts))


if __name__ == "__main__":
    main()
