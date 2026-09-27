#!/usr/bin/env python
"""
把 Gitee 上已发布的 release 的名称/正文回填成「产品介绍 + 下载指引」。

为什么需要它：百度搜 excelrouter 排第一的是 Gitee 发行版页，摘要取自 release
正文——以前 CI 只写了一句「同步自 GitHub Release」，搜索结果里看不出这是什么工具。
新发的版本由 release.yml 用同一份模板 .github/gitee_release_body.md 自动生成；
这个脚本只用来补历史版本（或改完模板后重刷）。

用法（令牌 = GitHub secret GITEE 那个 Gitee 私人令牌）：
    set GITEE_TOKEN=xxxx
    py update_gitee_releases.py              # 只更新最新版
    py update_gitee_releases.py --all        # 所有版本
    py update_gitee_releases.py --dry-run    # 只打印将要写入的内容
"""
import argparse
import json
import os
import pathlib
import sys
import urllib.parse
import urllib.request

API = "https://gitee.com/api/v5/repos/Marsandsea/Excelrouter/releases"
TEMPLATE = pathlib.Path(__file__).parent / ".github" / "gitee_release_body.md"


def request(url, data=None, method="GET"):
    body = urllib.parse.urlencode(data).encode() if data else None
    req = urllib.request.Request(url, data=body, method=method)
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def main():
    ap = argparse.ArgumentParser(description="回填 Gitee release 名称/正文")
    ap.add_argument("--all", action="store_true", help="更新全部版本（默认只更新最新版）")
    ap.add_argument("--dry-run", action="store_true", help="只打印，不写入")
    args = ap.parse_args()

    token = os.environ.get("GITEE_TOKEN", "")
    if not token and not args.dry_run:
        sys.exit("请先设置环境变量 GITEE_TOKEN")

    template = TEMPLATE.read_text(encoding="utf-8")
    releases = request(f"{API}?per_page=100&direction=desc")
    if not args.all:
        releases = releases[:1]

    for rel in releases:
        tag = rel["tag_name"]
        ver = tag.lstrip("v")
        payload = {
            "tag_name": tag,
            "name": f"v{ver} · Excel 智能拆分工具 Windows 免安装版",
            "body": template.replace("{ver}", ver),
        }
        if args.dry_run:
            print(f"--- {tag} (id={rel['id']}) ---\n{payload['name']}\n\n{payload['body']}")
            continue
        payload["access_token"] = token
        request(f"{API}/{rel['id']}", payload, method="PATCH")
        print(f"已更新 {tag}")


if __name__ == "__main__":
    main()
