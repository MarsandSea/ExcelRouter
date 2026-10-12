#!/usr/bin/env python
"""
生成 / 回填 Gitee 发行版的名称和正文：「下载直链 + 产品介绍 + 下载指引 + 本版更新」。

为什么需要它：百度搜 excelrouter 排第一的是 Gitee 发行版页，摘要取自 release 正文；
国内同事也主要从 Gitee 下载——正文里的链接必须指向 Gitee（GitHub 国内常打不开）。
正文 = 模板 .github/gitee_release_body.md，其中 {notes} 换成
docs/release-notes-v{ver}.md 里「## 正文」以下的内容（没有这个文件就留空）。
release.yml 打 tag 时用 --write-body 生成正文建发行版；本脚本也用来改历史版本。

用法（令牌 = GitHub secret GITEE 那个 Gitee 私人令牌）：
    set GITEE_TOKEN=xxxx
    py update_gitee_releases.py                  # 只更新最新版
    py update_gitee_releases.py --tag v3.0.0     # 只更新这一版
    py update_gitee_releases.py --all            # 所有版本
    py update_gitee_releases.py --dry-run        # 只打印将要写入的内容
    py update_gitee_releases.py --write-body 3.0.0 body.md   # 只生成正文写进文件（CI 用）
"""
import argparse
import json
import os
import pathlib
import re
import sys
import urllib.parse
import urllib.request

API = "https://gitee.com/api/v5/repos/Marsandsea/Excelrouter/releases"
ROOT = pathlib.Path(__file__).parent
TEMPLATE = ROOT / ".github" / "gitee_release_body.md"
# release-linux.yml 发麒麟包时把这一行换成麒麟下载指引；重刷正文时要把它保留下来
LINUX_PLACEHOLDER = "- 银河麒麟 / 统信 UOS 版按需构建"
LINUX_LINE = "- **银河麒麟 / 统信 UOS**"


def release_notes(ver):
    """docs/release-notes-v{ver}.md 里「## 正文」以下的内容；去掉第一行的粗体大标题。"""
    p = ROOT / "docs" / f"release-notes-v{ver}.md"
    if not p.exists():
        return ""
    text = p.read_text(encoding="utf-8").replace("\r\n", "\n")
    if "## 正文" not in text:
        return ""
    lines = text.split("## 正文", 1)[1].strip().split("\n")
    if lines and lines[0].startswith("**ExcelRouter"):
        lines = lines[1:]
    # 模板里各节是 ###，更新说明的小节降一级
    body = "\n".join("#" + ln if ln.startswith("### ") else ln for ln in lines).strip()
    return f"### 本版更新\n\n{body}" if body else ""


def build_body(ver, old_body=""):
    """按模板生成正文；old_body 里已有麒麟下载指引时原样带上。"""
    text = TEMPLATE.read_text(encoding="utf-8").replace("\r\n", "\n")
    text = text.replace("{ver}", ver).replace("{notes}", release_notes(ver))
    old = (old_body or "").replace("\r\n", "\n")
    linux = [ln for ln in old.split("\n") if ln.startswith(LINUX_LINE)]
    if linux:
        text = "\n".join(linux[0] if ln.startswith(LINUX_PLACEHOLDER) else ln
                         for ln in text.split("\n"))
    return re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"


def release_name(ver, old_name=""):
    if "银河麒麟" in (old_name or ""):
        return old_name
    return f"v{ver} · Excel 智能拆分汇总工具 Windows 免安装版"


def request(url, data=None, method="GET"):
    body = urllib.parse.urlencode(data).encode() if data else None
    req = urllib.request.Request(url, data=body, method=method)
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def main():
    ap = argparse.ArgumentParser(description="生成 / 回填 Gitee release 名称和正文")
    ap.add_argument("--all", action="store_true", help="更新全部版本（默认只更新最新版）")
    ap.add_argument("--tag", help="只更新这一个版本，如 v3.0.0")
    ap.add_argument("--dry-run", action="store_true", help="只打印，不写入")
    ap.add_argument("--write-body", nargs=2, metavar=("VER", "PATH"), help="只生成正文写进文件")
    args = ap.parse_args()

    if args.write_body:
        ver, path = args.write_body
        pathlib.Path(path).write_text(build_body(ver.lstrip("v")), encoding="utf-8")
        return

    token = os.environ.get("GITEE_TOKEN", "")
    if not token and not args.dry_run:
        sys.exit("请先设置环境变量 GITEE_TOKEN")

    releases = request(f"{API}?per_page=100&direction=desc")
    if args.tag:
        releases = [r for r in releases if r["tag_name"] == args.tag]
        if not releases:
            sys.exit(f"Gitee 上没有 {args.tag} 的发行版")
    elif not args.all:
        releases = releases[:1]

    for rel in releases:
        tag = rel["tag_name"]
        ver = tag.lstrip("v")
        payload = {
            "tag_name": tag,
            "name": release_name(ver, rel.get("name")),
            "body": build_body(ver, rel.get("body")),
        }
        if args.dry_run:
            print(f"--- {tag} (id={rel['id']}) ---\n{payload['name']}\n\n{payload['body']}")
            continue
        payload["access_token"] = token
        request(f"{API}/{rel['id']}", payload, method="PATCH")
        print(f"已更新 {tag}")


if __name__ == "__main__":
    main()
