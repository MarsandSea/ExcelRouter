#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_clawhub.py —— 核实 ExcelRouter 在 ClawHub 上的发布/审核状态。

背景（2026-09-14 首次打通 ClawHub 时摸清）：
  - ClawHub 的版本线与 SkillHub **互相独立**：没传 --version 时首发 1.0.0，
    之后内容变化自动递增 patch。别拿它去对 SkillHub 的 2.7.x。
  - 「发布成功」不等于「已公开」：提交后进 pending.publication 等安全扫描，
    这期间搜索搜不到、详情接口回 Skill not found。扫描是服务端异步的，
    本地干等没有意义，用本脚本隔一段时间探一次即可。

用法：
  python check_clawhub.py                # 查一次状态
  python check_clawhub.py --watch        # 每 60s 轮询，直到扫描出结果
  python check_clawhub.py --api          # 额外用公开搜索接口探测是否已可被搜到

需要 clawhub CLI。首次使用：
  npm i clawhub        # 入口是 node_modules/clawhub/bin/clawdhub.js（注意 clawDhub）
  node node_modules/clawhub/bin/clawdhub.js login --token <clh_...>

也可用环境变量 CLAWHUB_CLI 指定 CLI 入口的绝对路径。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request

SLUG = "excelrouter"
PENDING_MARK = "pending"          # 状态里含这个字样就是「还在审核」
TIMEOUT = 60
API_TIMEOUT = 20


def find_cli() -> list[str] | None:
    """定位 clawhub CLI。Windows 下 node_modules/.bin/clawhub 是 shell 脚本，
    直接 subprocess 会报 WinError 193 —— 必须走真实的 bin/clawdhub.js。"""
    env = os.environ.get("CLAWHUB_CLI")
    if env and os.path.exists(env):
        return ["node", env]

    # 从 cwd 向上逐级找 node_modules（兼容在仓库任意子目录下运行）
    cur = os.path.abspath(os.getcwd())
    while True:
        cand = os.path.join(cur, "node_modules", "clawhub", "bin", "clawdhub.js")
        if os.path.exists(cand):
            return ["node", cand]
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent

    if shutil.which("clawhub"):
        return ["clawhub"]
    return None


def verify(cli: list[str]) -> tuple[str, str]:
    """返回 (状态原文, 'pending' | 'ok' | 'error')。"""
    try:
        r = subprocess.run(cli + ["skill", "verify", SLUG],
                           capture_output=True, text=True, timeout=TIMEOUT)
    except subprocess.TimeoutExpired:
        return "verify 超时", "error"
    except OSError as e:
        return f"无法执行 CLI：{e}", "error"

    lines = [ln for ln in (r.stdout + r.stderr).strip().splitlines() if ln.strip()]
    msg = lines[0] if lines else "(无输出)"
    if "pending" in msg.lower():
        return msg, "pending"
    if r.returncode == 0:
        return "\n".join(lines[:30]), "ok"
    return "\n".join(lines[:30]), "error"


def public_search() -> str:
    """用公开搜索接口探测技能是否已经能被搜到（扫描通过后才会出现）。"""
    url = "https://clawhub.ai/api/v1/skills?q=" + SLUG
    try:
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=API_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception as e:  # noqa: BLE001
        return f"探测失败：{type(e).__name__}: {e}"
    items = data.get("items") or []
    for it in items:
        if it.get("slug") == SLUG:
            return f"✅ 已公开：{it.get('displayName')}（搜索结果第 {items.index(it)+1} 位）"
    return f"⏳ 尚未出现在搜索结果中（返回 {len(items)} 条，均非本技能）"


def main() -> int:
    p = argparse.ArgumentParser(description="核实 ClawHub 发布/审核状态")
    p.add_argument("--watch", action="store_true", help="轮询直到审核出结果")
    p.add_argument("--interval", type=int, default=60, help="轮询间隔秒数，默认 60")
    p.add_argument("--api", action="store_true", help="额外探测公开搜索接口")
    args = p.parse_args()

    cli = find_cli()
    if not cli:
        print("找不到 clawhub CLI。先执行：\n"
              "  npm i clawhub\n"
              "  node node_modules/clawhub/bin/clawdhub.js login --token <clh_...>\n"
              "或设环境变量 CLAWHUB_CLI 指向 bin/clawdhub.js 的绝对路径。", file=sys.stderr)
        return 2

    print(f"CLI: {' '.join(cli)}")
    print("=" * 62)

    attempt = 0
    while True:
        msg, state = verify(cli)
        stamp = time.strftime("%H:%M:%S")
        icon = {"pending": "⏳", "ok": "✅", "error": "❌"}[state]
        print(f"[{stamp}] {icon} {msg}")
        if state == "ok":
            print("\n--- 完整输出 ---")
            print(msg)
            return 0
        if state == "error" and "hidden by moderation" not in msg:
            return 1

        if args.api:
            print(f"         {public_search()}")

        if not args.watch:
            print("\n提示：加 --watch 可轮询等待扫描完成。")
            return 0
        attempt += 1
        if attempt >= 30:
            print("\n轮询 30 次仍未出结果，先停。可稍后手动再查一次。")
            return 0
        time.sleep(max(args.interval, 10))


if __name__ == "__main__":
    sys.exit(main())
