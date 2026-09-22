#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
log_skillhub_stats.py —— 把 SkillHub 的下载量等指标按次追加到 CSV，用来画增长曲线。

为什么需要它：
  SkillHub 接口只给"当前累计值"，**不提供时间序列**（没有日下载曲线、没有调用次数）。
  想要曲线，唯一办法是自己定期跑一次、把数字记下来。跑得越勤，曲线越细。

用法：
  python log_skillhub_stats.py            # 追加一条记录（同一天可跑多次）
  python log_skillhub_stats.py --show     # 打印 CSV 全文 + 首末对比（不写文件）
  python log_skillhub_stats.py --rank 0   # 跳过排名查询（更快，只记总量）

输出：stats/skillhub_history.csv
  date,time,downloads,installs,stars,comments,versions,latest_version,rank_excel拆分,rank_表格拆分
"""
from __future__ import annotations

import argparse
import csv
import json
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

HOST = "https://api.skillhub.cn"
SLUG = "excelrouter"
RANK_KEYWORDS = ["excel拆分", "表格拆分"]
CSV_PATH = Path(__file__).resolve().parent / "stats" / "skillhub_history.csv"
FIELDS = ["date", "time", "downloads", "installs", "stars", "comments",
          "versions", "latest_version"] + [f"rank_{k}" for k in RANK_KEYWORDS]


def api(path: str, **params):
    params["t"] = str(int(time.time()))          # 破 CDN 缓存，见 check_skillhub.py
    url = f"{HOST}{path}?{urllib.parse.urlencode(params)}"
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))


def fetch_rank(kw: str):
    try:
        d = api("/api/v1/search", q=kw)
        for i, r in enumerate(d.get("results") or [], 1):
            if r.get("slug") == SLUG:
                return i
    except Exception:  # noqa: BLE001
        pass
    return ""          # 空 = 未进榜或查询失败


def collect(with_rank: bool = True) -> dict:
    d = api(f"/api/v1/skills/{SLUG}")
    stats = ((d.get("skill") or {}).get("stats")) or {}
    now = datetime.now()
    row = {
        "date": now.strftime("%Y-%m-%d"),
        "time": now.strftime("%H:%M"),
        "downloads": stats.get("downloads", ""),
        "installs": stats.get("installs", ""),
        "stars": stats.get("stars", ""),
        "comments": stats.get("comments", ""),
        "versions": stats.get("versions", ""),
        "latest_version": ((d.get("latestVersion") or {}).get("version")) or "",
    }
    for kw in RANK_KEYWORDS:
        row[f"rank_{kw}"] = fetch_rank(kw) if with_rank else ""
    return row


def append(row: dict) -> None:
    CSV_PATH.parent.mkdir(parents=True, exist_ok=True)
    new = not CSV_PATH.exists()
    with CSV_PATH.open("a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        if new:
            w.writeheader()
        w.writerow(row)
    print(f"已追加 → {CSV_PATH}")
    print("  " + "  ".join(f"{k}={row[k]}" for k in FIELDS))


def show() -> None:
    if not CSV_PATH.exists():
        print(f"还没有记录文件：{CSV_PATH}\n先跑一次 `python log_skillhub_stats.py` 吧。")
        return
    rows = list(csv.DictReader(CSV_PATH.open(encoding="utf-8-sig")))
    print(f"{CSV_PATH}  （共 {len(rows)} 条）\n")
    for r in rows:
        print("  " + "  ".join(f"{r.get(k, '')}" for k in FIELDS))
    if len(rows) >= 2:
        a, b = rows[0], rows[-1]
        try:
            dl = int(b["downloads"]) - int(a["downloads"])
            days = (datetime.strptime(b["date"], "%Y-%m-%d")
                    - datetime.strptime(a["date"], "%Y-%m-%d")).days or 1
            print(f"\n首末对比：{a['downloads']} → {b['downloads']}  "
                  f"共 +{dl}，跨度 {days} 天，日均 {dl / days:.1f}")
        except (ValueError, KeyError):
            pass


def main() -> int:
    p = argparse.ArgumentParser(description="记录 SkillHub 指标，攒日曲线")
    p.add_argument("--show", action="store_true", help="只看历史记录，不写入")
    p.add_argument("--rank", type=int, default=1,
                   help="是否查询关键词排名（1=查，0=跳过；默认 1）")
    args = p.parse_args()

    if args.show:
        show()
        return 0
    try:
        row = collect(with_rank=bool(args.rank))
    except Exception as e:  # noqa: BLE001
        print(f"抓取失败：{type(e).__name__}: {e}", file=sys.stderr)
        return 2
    append(row)
    return 0


if __name__ == "__main__":
    sys.exit(main())
