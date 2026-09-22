#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
git push 走不通时的退路：用 GitHub Git Data API 推送（只走 api.github.com）。

为什么需要它：本环境对 github.com 的访问时通时断（git over https 经常
Connection reset / Empty reply from server），但 api.github.com 一直可达。
改动文件不多、且都是文本时，用这个脚本比等网络恢复更省事。

用法：
    python push_via_api.py <文件1> [文件2 ...] -m "提交说明"
    python push_via_api.py <文件...> -m "说明" --dry-run   # 只看会推什么

限制：
  - 只处理文本文件（二进制请用 git push）
  - 不处理删除/重命名（只做"新增或覆盖"）
  - 推完本地 git 会与远端 sha 不一致，需要 git fetch && git reset --soft origin/main 对齐
"""
import argparse
import json
import pathlib
import subprocess
import sys

REPO = "MarsandSea/excel-router"
BRANCH = "main"


def gh(*args, inp=None):
    r = subprocess.run(["gh", "api", *args], capture_output=True, text=True, input=inp)
    if r.returncode != 0:
        sys.exit("gh api 失败: " + (r.stderr or r.stdout)[:500])
    return json.loads(r.stdout) if r.stdout.strip() else {}


def main():
    ap = argparse.ArgumentParser(description="用 GitHub API 推送（git push 不通时的退路）")
    ap.add_argument("files", nargs="*", help="要推送的文件（相对仓库根的路径）；只删除时可为空")
    ap.add_argument("-m", "--message", required=True, help="提交说明")
    ap.add_argument("--dry-run", action="store_true", help="只显示将要推送的内容")
    ap.add_argument("--branch", default=BRANCH)
    ap.add_argument("--delete", nargs="*", default=[],
                    help="要从仓库删除的路径（Git API 里把 blob sha 置空即为删除）")
    args = ap.parse_args()

    missing = [f for f in args.files if not pathlib.Path(f).is_file()]
    if missing:
        sys.exit("文件不存在: " + ", ".join(missing))

    head = gh(f"repos/{REPO}/git/ref/heads/{args.branch}")["object"]["sha"]
    base_tree = gh(f"repos/{REPO}/git/commits/{head}")["tree"]["sha"]
    print(f"远端当前 HEAD: {head[:8]}")
    print(f"将推送 {len(args.files)} 个文件:")
    for f in args.files:
        print("   -", f)

    if args.dry_run:
        print("\n(dry-run) 未实际推送")
        return

    tree = []
    for f in args.files:
        # 统一转 LF：仓库里存 LF，本地工作区可能是 CRLF，不转会显示整文件改动
        content = pathlib.Path(f).read_text(encoding="utf-8").replace("\r\n", "\n")
        blob = gh(
            f"repos/{REPO}/git/blobs", "-X", "POST", "--input", "-",
            inp=json.dumps({"content": content, "encoding": "utf-8"}),
        )
        tree.append({"path": f, "mode": "100644", "type": "blob", "sha": blob["sha"]})

    for f in args.delete:
        # sha 为 null 即表示删除该路径
        tree.append({"path": f, "mode": "100644", "type": "blob", "sha": None})
        print("   删除:", f)

    new_tree = gh(
        f"repos/{REPO}/git/trees", "-X", "POST", "--input", "-",
        inp=json.dumps({"base_tree": base_tree, "tree": tree}),
    )
    new_commit = gh(
        f"repos/{REPO}/git/commits", "-X", "POST", "--input", "-",
        inp=json.dumps({"message": args.message, "tree": new_tree["sha"], "parents": [head]}),
    )
    gh(
        f"repos/{REPO}/git/refs/heads/{args.branch}", "-X", "PATCH", "--input", "-",
        inp=json.dumps({"sha": new_commit["sha"]}),
    )
    print(f"\n已推送 {new_commit['sha'][:8]}")
    print(f"https://github.com/{REPO}/commit/{new_commit['sha'][:8]}")
    print("\n提示：本地 git 记得对齐 -> git fetch origin && git reset --soft origin/main")


if __name__ == "__main__":
    main()
