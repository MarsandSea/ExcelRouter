# 打包 excelrouter skill 上架 SkillHub
# 源：git 仓库里的 excelrouter-skill/skills/excelrouter（唯一真源）；产物：本文件同目录下的 zip
#
# 路径一律按「本脚本所在位置」推导，不写死绝对路径——真源迁移过一次，
# 写死的路径会让整个打包静默失效（改成指向不存在的目录却仍能跑出空包）。
import os
import re
import zipfile
from pathlib import Path

BASE = Path(__file__).resolve().parent
SKILL_DIR = BASE / "excelrouter-skill" / "skills" / "excelrouter"

# 版本号从 manifest.yaml 读，避免与 manifest 双源漂移
_manifest = (SKILL_DIR / "manifest.yaml").read_text(encoding="utf-8")
_m = re.search(r'^version:\s*"?([^"\n]+?)"?\s*$', _manifest, re.M)
if not _m:
    raise SystemExit("manifest.yaml 里没找到 version: 行，无法决定版本号")
VERSION = _m.group(1).strip()
OUTPUT_ZIP = BASE / f"excelrouter-skill-v{VERSION}.zip"

# SkillHub 不接受隐藏文件（. 开头），package 时必须过滤掉（.upstream-tag 等）
exclude_dirs = {"__pycache__", ".git", ".vscode", ".idea", "node_modules", ".pytest_cache", ".ruff_cache"}
exclude_files = {".DS_Store", "Thumbs.db"}


def should_include(name, is_dir=False):
    if name.startswith("."):
        return False
    if is_dir and name in exclude_dirs:
        return False
    if not is_dir and name in exclude_files:
        return False
    return True


count = 0
with zipfile.ZipFile(OUTPUT_ZIP, "w", zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(SKILL_DIR):
        dirs[:] = [d for d in dirs if should_include(d, is_dir=True)]
        for f in files:
            if not should_include(f):
                print("  跳过(隐藏):", os.path.join(root, f))
                continue
            fp = os.path.join(root, f)
            arcname = os.path.relpath(fp, SKILL_DIR)
            # 统一成 LF 行尾再入包。
            # 为什么必须做：本机 git 是 core.autocrlf=true，受版本控制的文件（scripts/*.py 等）
            # 检出到工作区是 CRLF —— 直接打包会把 CRLF 带进 zip。一旦这个包被放到 Linux/macOS
            # 或容器里（手机上跑 skill / 云端执行），脚本还带 CR 就会炸（`#!/usr/bin/env python3\r`
            # 这种是经典故障）。统一 LF 后跨平台都能跑。
            data = Path(fp).read_bytes().replace(b"\r\n", b"\n")
            zf.writestr(arcname, data)
            count += 1

print(f"Packed: {OUTPUT_ZIP}")
print(f"File count: {count}")
print("Line endings: normalized to LF")
