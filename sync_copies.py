# -*- coding: utf-8 -*-
"""
excelrouter skill · 本地副本统一同步脚本

背景：同一个 skill 散落在多个平台各自的扫描目录里，以前靠手工 cp，
      既容易漏同步、也容易把平台的元文件（_meta.json / _icon.png）冲掉。
      这个脚本把「真源 → 各副本」的一次性同步 + 校验固化下来。

唯一真源：与本脚本同级的 excelrouter-skill\\skills\\excelrouter\\（git 仓库）

注意：SRC 由脚本自身位置推导，**不要改回写死的绝对路径**——
2026-09-05 工作区从 G:\\WorkBuddy\\excel-router 迁到 G:\\WorkSpace\\excel-router 后，
写死路径会让脚本误判「真源目录不存在」而整个失效。

用法：
    python sync_copies.py              # 同步全部副本
    python sync_copies.py --dry-run    # 只看会同步什么、各副本状态，不写盘
    python sync_copies.py --check      # 只校验各副本与真源是否一致（不同步）
    python sync_copies.py --only claude  # 只同步某一份（名字见下面 COPIES 的 key）

约定：
  - 只做「覆盖式复制」，绝不删除副本里的任何文件——平台的元文件
    （_meta.json / _icon.png / _skillhub_meta.json）不在真源里，删了平台会认不出这个 skill。
  - 真源里的测试残留（*.txt / *.log 之类）不会被自动清理，同步前请保持真源干净。
"""
import argparse
import filecmp
import hashlib
import os
import shutil
import sys
from pathlib import Path

# ---------------------------------------------------------------- 真源
# 由脚本自身位置推导，跟着仓库走，不写死绝对路径（见文件头说明）
SRC = str(Path(__file__).resolve().parent / "excelrouter-skill" / "skills" / "excelrouter")

# ---------------------------------------------------------------- 副本清单
# key      : --only 时用的短名
# name     : 人类可读说明（谁在加载它、能不能删）
# path     : 副本所在目录
# required : True 表示平台强依赖、删了就加载不到
COPIES = [
    {
        "key": "workbuddy",
        "name": "WorkBuddy（SkillHub 安装位，WorkBuddy 实际加载这一份）",
        "path": r"C:\Users\Administrator\.workbuddy\skills\excelrouter__skillhub",
        "required": True,
    },
    {
        "key": "claude",
        "name": "Claude Code（~/.claude/skills，用户要求保留）",
        "path": r"C:\Users\Administrator\.claude\skills\excelrouter",
        "required": True,
    },
    {
        "key": "wps",
        "name": "WPS 灵犀（serverdir/user_skills）",
        "path": r"C:\Users\Administrator\AppData\Roaming\WPS 灵犀\serverdir\user_skills\excelrouter-skill-pkg",
        "required": False,
    },
]

# 平台自有文件：副本里有、真源里没有，同步时必须原样保留
PLATFORM_FILES = {"_meta.json", "_icon.png", "_skillhub_meta.json"}

# 不同步的垃圾文件（真源里若出现会被跳过，也算一种兜底）
SKIP_NAMES = {".DS_Store", "Thumbs.db"}
SKIP_SUFFIX = {".pyc"}


def _sha(path):
    """内容哈希，**忽略行尾差异**（CRLF / LF 视为同一内容）。

    为什么必须归一化：本机 git 是 core.autocrlf=true —— 受版本控制的文件（如
    scripts/*.py）检出到工作区是 CRLF，而不受版本控制 / 被忽略的文件工作区是 LF。
    逐字节比会把「只差行尾」的文件报成「内容不一致」（2026-09-15 踩坑：
    vendor/core/*.py 每个文件都恰好大出「行数」个字节，看着像被改过，其实是 CR+行数）。
    """
    h = hashlib.sha256()
    with open(path, "rb") as f:
        data = f.read().replace(b"\r\n", b"\n")
    h.update(data)
    return h.hexdigest()


def _same(src_file, dst_file):
    """内容是否相同（忽略行尾差异，且要求同样字节数才算，避免误合并）。"""
    if os.path.getsize(src_file) == os.path.getsize(dst_file):
        if filecmp.cmp(src_file, dst_file, shallow=False):
            return True
    return _sha(src_file) == _sha(dst_file)


def _src_files():
    """返回真源里需要同步的相对路径列表。"""
    out = []
    for root, dirs, files in os.walk(SRC):
        dirs[:] = [d for d in dirs if d not in ("__pycache__", ".git", ".pytest_cache", ".ruff_cache")]
        for fn in files:
            if fn in SKIP_NAMES or os.path.splitext(fn)[1] in SKIP_SUFFIX:
                continue
            if fn.startswith("."):
                continue
            out.append(os.path.relpath(os.path.join(root, fn), SRC))
    return sorted(out)


def _sync_one(copy, rel_files, dry_run):
    """把 rel_files 同步到 copy['path']。返回 (待更新数, 已最新数)。"""
    dst_root = copy["path"]
    updated, same = 0, 0
    for rel in rel_files:
        src_file = os.path.join(SRC, rel)
        dst_file = os.path.join(dst_root, rel)
        os.makedirs(os.path.dirname(dst_file), exist_ok=True)
        # 已存在且内容一致 → 跳过（忽略行尾差异，避免把纯 CRLF/LF 差异当改动刷一遍）
        if os.path.exists(dst_file) and _same(src_file, dst_file):
            same += 1
            continue
        updated += 1
        if not dry_run:
            shutil.copy2(src_file, dst_file)
    return updated, same


def _check_one(copy, rel_files):
    """校验副本与真源是否一致。返回 (缺失, 不一致) 两个列表。"""
    missing, differs = [], []
    for rel in rel_files:
        dst_file = os.path.join(copy["path"], rel)
        if not os.path.exists(dst_file):
            missing.append(rel)
            continue
        if _sha(os.path.join(SRC, rel)) != _sha(dst_file):
            differs.append(rel)
    return missing, differs


def main():
    ap = argparse.ArgumentParser(description="excelrouter skill 本地副本统一同步")
    ap.add_argument("--dry-run", action="store_true", help="只报告，不写盘")
    ap.add_argument("--check", action="store_true", help="只校验一致性，不同步")
    ap.add_argument("--only", default="", help="只处理某一份副本（workbuddy / claude / wps）")
    args = ap.parse_args()

    if not os.path.isdir(SRC):
        print(f"真源目录不存在：{SRC}", file=sys.stderr)
        return 1

    rel_files = _src_files()
    targets = COPIES
    if args.only:
        targets = [c for c in COPIES if c["key"] == args.only]
        if not targets:
            print(f"没有 key 为 {args.only!r} 的副本，可选：{[c['key'] for c in COPIES]}")
            return 1

    print(f"真源：{SRC}")
    print(f"待同步文件：{len(rel_files)} 个")
    print("-" * 62)

    rc = 0
    for c in targets:
        path = c["path"]
        print(f"\n[{c['key']}] {c['name']}")
        print(f"  路径：{path}")
        if not os.path.isdir(path):
            print("  ⚠️  目录不存在（跳过）" + ("【平台强依赖，建议检查】" if c["required"] else ""))
            rc = rc or (1 if c["required"] else 0)
            continue

        if args.check:
            missing, differs = _check_one(c, rel_files)
            if not missing and not differs:
                print(f"  ✅ 与真源完全一致（{len(rel_files)} 个文件）")
            else:
                print(f"  ⚠️  缺失 {len(missing)} 个，内容不一致 {len(differs)} 个")
                for m in missing[:5]:
                    print(f"      缺：{m}")
                for d in differs[:5]:
                    print(f"      异：{d}")
                if len(missing) + len(differs) > 10:
                    print(f"      ……（共 {len(missing) + len(differs)} 处）")
                rc = 1
            continue

        updated, same = _sync_one(c, rel_files, args.dry_run)
        tag = "（dry-run）" if args.dry_run else ""
        print(f"  {'将更新' if args.dry_run else '已更新'} {updated} 个，已最新 {same} 个 {tag}")
        # 提醒平台元文件仍完好
        kept = [f for f in PLATFORM_FILES if os.path.exists(os.path.join(path, f))]
        if kept:
            print(f"  平台元文件保留：{', '.join(kept)}")

    print("\n" + "-" * 62)
    if not args.check and not args.dry_run:
        print("同步完成。如需上架，再跑：python package_excelrouter.py")
    return rc


if __name__ == "__main__":
    sys.exit(main())
