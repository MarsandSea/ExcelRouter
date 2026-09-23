#!/usr/bin/env python3
"""
check_skillhub.py —— 核实 ExcelRouter 在 SkillHub 上的真实发布状态与搜索排名。

为什么需要它：
  SkillHub 的读接口走 CDN 缓存，发布后裸查会一直返回旧版本，
  让人误以为"这次没推成功"（2026-09-10 就这么误判过一次）。
  本脚本默认带 `?t=<unix时间戳>` 破缓存，拿到的才是真数据。

用法：
  python check_skillhub.py                     # 版本状态 + 默认关键词排名
  python check_skillhub.py --expect 2.7.4      # 断言线上版本，不一致则退出码 1（可进 CI）
  python check_skillhub.py -k excel拆分 -k 表格拆分
  python check_skillhub.py --versions          # 额外列出全部历史版本

只用标准库，无需 token。
"""
from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime

HOST = "https://api.skillhub.cn"   # 注意：主站 skillhub.cn/api/... 返回的是前端 HTML，不能用
SLUG = "excelrouter"
DEFAULT_KEYWORDS = ["excel拆分", "表格拆分", "PDF加密分发"]
TIMEOUT = 20


def api(endpoint: str, **params):
    """带破缓存的 GET。params 里塞 t=<时间戳> 是刻意为之，见模块 docstring。

    形参叫 endpoint 而不是 path：/file 端点的查询参数本身就叫 path，
    叫 path 会和它撞成 "got multiple values for argument"。"""
    params["t"] = str(int(time.time()))
    url = f"{HOST}{endpoint}?{urllib.parse.urlencode(params)}"
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
        body = resp.read().decode("utf-8")
    try:
        return json.loads(body)
    except json.JSONDecodeError:
        return body   # /file 端点返回的是文件原文，不是 JSON


def fmt_ts(ms: int | None) -> str:
    if not ms:
        return "-"
    return datetime.fromtimestamp(ms / 1000).astimezone().strftime("%Y-%m-%d %H:%M:%S")


def show_status() -> str:
    """打印线上最新版本 + displayName，返回版本号。"""
    d = api(f"/api/v1/skills/{SLUG}")
    lv = d.get("latestVersion") or {}
    sk = d.get("skill") or {}
    stats = sk.get("stats") or {}

    ver = lv.get("version", "?")
    print("=" * 62)
    print("SkillHub 线上状态（已破缓存）")
    print("=" * 62)
    print(f"  最新版本   : {ver}")
    print(f"  发布时间   : {fmt_ts(lv.get('createdAt'))}")
    print(f"  显示名     : {sk.get('displayName', '-')}")
    print(f"  下载 / 版本: {stats.get('downloads', '-')} / {stats.get('versions', '-')}")
    print(f"  安装 / 收藏: {stats.get('installs', '-')} / {stats.get('stars', '-')}")
    # 认领：claim_state 为 unclaimed 但 claimable=False 时，是「无需认领」而不是
    # 「忘了认领」—— 技能本来就发在自己账号下，认领是给从 GitHub 自动抓取、
    # 没有主人的技能用的。这条误报过一次，别再照着它去找认领入口。
    claim = sk.get("claim_state", "-")
    if claim == "unclaimed" and sk.get("claimable") is False:
        claim = "无需认领（已在自己账号下）"
    print(f"  评论 / 认领: {stats.get('comments', '-')} / {claim}")
    print(f"  命名空间   : {(d.get('namespace') or {}).get('canonicalName', '-')}")

    # ★ 详情页「概述」那一栏渲染的是**包里的 README.md**，不是 skill.overviewMd。
    #   overviewMd 全站所有技能（含百万下载的）都是空的、发布流程也不写它 ——
    #   盯着它只会得出「文案没上去」的错误结论。2026-09-23 实测确认：
    #   本技能 README.md 404 时，详情页「概述」显示「无法加载文档 / Failed to fetch」。
    try:
        rm = api(f"/api/v1/skills/{SLUG}/file", path="README.md")
        n = len(rm) if isinstance(rm, str) else len(json.dumps(rm))
        print(f"  概述文档   : README.md 正常，{n} 字符")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            print("  概述文档   : （README.md 不在包里！详情页「概述」会显示"
                  "「无法加载文档」——它渲染的就是这个文件，不是 overviewMd）")
        else:
            print(f"  概述文档   : 查询失败 HTTP {e.code}")
    print(f"  分类       : {sk.get('category', '-')} / "
          f"{'、'.join(c.get('name', '') for c in (sk.get('subCategories') or [])) or '-'}")

    # 名字开头是否命中关键词 —— 这是排名高低的决定因素
    name = (sk.get("displayName") or "").lower()
    hits = [k for k in DEFAULT_KEYWORDS if k.lower() in name]
    lead = [k for k in hits if name.startswith(k.lower())]
    print(f"  名字命中   : {hits or '（无）'}")
    print(f"  开头命中   : {lead or '（无）'}  ← 只有开头命中才吃高权重")
    return ver


def show_versions(limit: int = 10) -> None:
    d = api(f"/api/v1/skills/{SLUG}/versions")
    vs = (d.get("versions") or [])[:limit]
    print()
    print(f"历史版本（最近 {len(vs)} 个，可反推是哪次 CI run 发的）")
    for v in vs:
        print(f"  - {v.get('version','?'):8s} {fmt_ts(v.get('createdAt'))}  {(v.get('changelog') or '')[:40]}")


def search_rank(keyword: str) -> tuple[int | None, float]:
    """返回 (本技能在搜索结果中的名次, 自己的得分)。名次 None 表示未出现。"""
    d = api("/api/v1/search", q=keyword)
    results = d.get("results") or []
    for i, r in enumerate(results, 1):
        if r.get("slug") == SLUG:
            return i, float(r.get("score") or 0)
    return None, 0.0


def show_ranks(keywords: list[str], top: int = 3) -> None:
    print()
    print("=" * 62)
    print("搜索排名")
    print("=" * 62)
    for kw in keywords:
        rank, score = search_rank(kw)
        d = api("/api/v1/search", q=kw)
        results = (d.get("results") or [])[:top]
        line = f"  「{kw}」→ " + (f"第 {rank} 名（score {score:.4f}）" if rank else "未进榜")
        print(line)
        for i, r in enumerate(results, 1):
            who = (r.get("display_name") or r.get("displayName") or "")[:32]
            star = "  <<< 本技能" if r.get("slug") == SLUG else ""
            print(f"       {i}. {who:32s} {float(r.get('score') or 0):.4f}{star}")


def main() -> int:
    p = argparse.ArgumentParser(description="核实 SkillHub 发布状态与搜索排名")
    p.add_argument("-k", "--keyword", action="append", dest="keywords",
                   help="要查排名的关键词，可重复；默认 excel拆分/表格拆分/PDF加密分发")
    p.add_argument("--expect", help="期望的线上版本号，不一致则退出码 1")
    p.add_argument("--versions", action="store_true", help="额外列出历史版本")
    p.add_argument("--no-rank", action="store_true", help="跳过搜索排名检查")
    args = p.parse_args()

    try:
        ver = show_status()
        if args.versions:
            show_versions()
        if not args.no_rank:
            show_ranks(args.keywords or DEFAULT_KEYWORDS)
    except Exception as e:  # noqa: BLE001 —— 探针脚本，任何网络/解析失败都只需报出来
        print(f"查询失败：{type(e).__name__}: {e}", file=sys.stderr)
        return 2

    print()
    if args.expect:
        if ver == args.expect:
            print(f"✅ 线上版本 = {args.expect}，发布已生效")
            return 0
        print(f"❌ 线上版本 = {ver}，期望 {args.expect} —— 发布尚未生效（或真的失败了）")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
