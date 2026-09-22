#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""excelrouter · 多仓 / 多副本同步体检（一条命令跑完全部同步链路）

为什么需要它（2026-09-15 复核后新增）：
    同一个 skill 现在同时活在 5 条链路上 ——
      ① 仓库 core/            : 桌面版真源
      ② skill/scripts/vendor/core/ : 上架用的 vendored 副本（跟 tag 拷过来的）
      ③ 三个平台扫描目录       : WorkBuddy / Claude Code / WPS 灵犀
      ④ 两个远端              : GitHub（真源）+ Gitee（纯镜像）
      ⑤ 两个市场渠道           : SkillHub + ClawHub
    链路一多，「看着同步了、其实没同步」的坑就多，靠人眼盯不住。
    这里把**可机器判定**的那几项固化下来，发版前跑一次。

用法：
    python check_sync.py           # 打印体检报告；有硬错误时退出码 1
    python check_sync.py --quiet   # 只打印 ⚠️/❌ 项

判读：
    ❌ = 硬错误（内容真的不一致 / 会被平台拒收），修完再发版
    ⚠️  = 需要人判断（可能是设计使然，也可能是坑，报告里都写了原因）
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SKILL_DIR = ROOT / "excelrouter-skill" / "skills" / "excelrouter"
VENDOR_CORE = SKILL_DIR / "scripts" / "vendor" / "core"
CORE = ROOT / "core"
UPSTREAM_TAG_FILE = ROOT / "excelrouter-skill" / ".upstream-tag"
MANIFEST = SKILL_DIR / "manifest.yaml"
PLUGIN_JSON = ROOT / "excelrouter-skill" / ".claude-plugin" / "plugin.json"
SKILL_MD = SKILL_DIR / "SKILL.md"

# 历史上散落出去的旧 skill 副本：都存在过，且都是**远古格式**（frontmatter 里连
# version/slug 都没有）。任何工具按「目录名像 excelrouter 的 skill」去打包上架，
# 就会拿这些内容去覆盖线上 —— 老仓的 cron 就这么干过一次。
LEGACY_SKILL_DIRS = [
    (r"G:\WorkSpace\Excelrouter-skill\skills\excelrouter",
     "旧独立仓工作副本（老真源路径，与真源同在前缀 G:\\WorkSpace 下，极易选错）"),
    (r"G:\WorkBuddy\excel-router\legacy\excelrouter-repo\skills\excelrouter",
     "工作区 legacy 归档里的旧副本"),
    (r"C:\Users\Administrator\excelrouter-skill-tmp\skills\excelrouter",
     "Windows 用户目录下的 tmp 副本"),
    (r"C:\Users\Administrator\WorkBuddy\2026-08-17-14-57-57\excelrouter-repo\skills\excelrouter",
     "更早的一次 WorkBuddy 会话产物副本"),
]

OK, WARN, FAIL = "ok", "warn", "fail"
ICON = {OK: "✅", WARN: "⚠️ ", FAIL: "❌"}


# ---------------------------------------------------------------- 工具
def norm_bytes(path: Path) -> bytes:
    """按行尾无关的方式读文件：CRLF / LF 视为同一内容。

    为什么必须归一化：本机 git 是 core.autocrlf=true，仓库里 core/*.py 是 LF、
    vendor/core/*.py 工作区是 CRLF，直接逐字节比会把 430 行的文件整个报成差异
    （2026-09-15 真的被骗过一次）。
    """
    return path.read_bytes().replace(b"\r\n", b"\n")


def sha_norm(path: Path) -> str:
    return hashlib.sha256(norm_bytes(path)).hexdigest()


def git(*args: str) -> tuple[int, str]:
    try:
        p = subprocess.run(["git", *args], cwd=str(ROOT), capture_output=True,
                           text=True, encoding="utf-8", errors="replace")
        return p.returncode, (p.stdout or "").strip()
    except FileNotFoundError:
        return 127, ""


def git_raw(*args: str) -> tuple[int, str]:
    """与 git() 相同，但**不 strip** —— 取文件内容时绝不能吃掉尾部换行，
    否则会把「末行少了 \\n」误判成内容不一致（2026-09-15 自己踩到过）。"""
    try:
        p = subprocess.run(["git", *args], cwd=str(ROOT), capture_output=True,
                           text=True, encoding="utf-8", errors="replace")
        return p.returncode, p.stdout or ""
    except FileNotFoundError:
        return 127, ""


def frontmatter(text: str) -> dict[str, str]:
    """逐行解析 frontmatter（与 skillhub CLI 同款粗略行为，够用于取字段）。"""
    lines = text.split("\n")
    if not lines or lines[0].strip() != "---":
        return {}
    out: dict[str, str] = {}
    for raw in lines[1:]:
        if raw.strip() == "---":
            break
        if not raw.strip() or raw.lstrip().startswith("#") or ":" not in raw:
            continue
        k, _, v = raw.partition(":")
        out[k.strip()] = v.strip().strip('"').strip("'")
    return out


def py_files(d: Path) -> list[str]:
    return sorted(p.name for p in d.glob("*.py") if p.name != "__init__.py")


# ---------------------------------------------------------------- 各项检查
def check_vendor_vs_core() -> tuple[str, list[str]]:
    """① vendored 代码是否与仓库 core/ 一致（内容层面）。"""
    if not VENDOR_CORE.is_dir():
        return FAIL, [f"vendor 目录不存在：{VENDOR_CORE}"]
    a, b = py_files(CORE), py_files(VENDOR_CORE)
    msgs: list[str] = []
    bad = False
    only_core, only_vendor = set(a) - set(b), set(b) - set(a)
    if only_core or only_vendor:
        bad = True
        msgs.append(f"文件名集合不同 —— 只在 core/: {sorted(only_core)}；只在 vendor: {sorted(only_vendor)}")
    diff = [f for f in sorted(set(a) & set(b)) if sha_norm(CORE / f) != sha_norm(VENDOR_CORE / f)]
    if diff:
        bad = True
        msgs.append("内容不一致：" + ", ".join(diff))
    if not bad:
        msgs.append(f"{len(a)} 个模块逐行一致（已忽略行尾差异）")
    return (FAIL if bad else OK), msgs


def check_vendor_provenance() -> tuple[str, list[str]]:
    """② vendor/core 是否等于 .upstream-tag 指向的那个 tag 的 core/（可溯源）。"""
    if not UPSTREAM_TAG_FILE.is_file():
        return WARN, [f"没有 {UPSTREAM_TAG_FILE.name}（无法溯源 vendor 代码来自哪个 tag）"]
    tag = UPSTREAM_TAG_FILE.read_text(encoding="utf-8").strip()
    rc, _ = git("rev-parse", "--verify", f"refs/tags/{tag}")
    if rc != 0:
        return WARN, [f"{UPSTREAM_TAG_FILE.name} 写的 {tag} 在本仓找不到（本地可能没 fetch 到该 tag）"]
    diff = []
    for f in py_files(VENDOR_CORE):
        rc, blob = git_raw("show", f"{tag}:core/{f}")
        if rc != 0:
            diff.append(f"{f}（该 tag 里没有这个文件）")
            continue
        if blob.encode("utf-8") != norm_bytes(VENDOR_CORE / f):
            diff.append(f)
    if diff:
        return FAIL, [f"vendor/core 与 tag {tag} 不一致：" + ", ".join(diff)]
    return OK, [f"vendor/core == {tag} 的 core/（.upstream-tag 名副其实）"]


def check_versions() -> tuple[str, list[str]]:
    """③ SKILL.md / manifest.yaml / plugin.json 三处版本号是否一致。"""
    msgs = []
    v_skill = frontmatter(SKILL_MD.read_text(encoding="utf-8")).get("version", "")
    m = re.search(r'^version:\s*"?([^"\n]+)"?', MANIFEST.read_text(encoding="utf-8"), re.M)
    v_manifest = (m.group(1).strip() if m else "")
    try:
        v_plugin = str(json.loads(PLUGIN_JSON.read_text(encoding="utf-8")).get("version", ""))
    except Exception as exc:  # noqa: BLE001
        return FAIL, [f"plugin.json 读不动：{exc}"]
    msgs.append(f"SKILL.md={v_skill or '(空)'}  manifest={v_manifest or '(空)'}  plugin.json={v_plugin or '(空)'}")
    vals = {v_skill, v_manifest, v_plugin}
    if len(vals) == 1 and v_skill:
        return OK, msgs
    msgs.append("三处不一致：manifest / plugin.json 由 CI 的 Align version fields 在发布时对齐；"
                "若刚改完 SKILL.md 还没发版，这里落后属正常，**别手工填**。")
    return WARN, msgs


def check_package_hygiene() -> tuple[str, list[str]]:
    """④ 发布包卫生：skillhub CLI 对整包含 `.` 开头文件会整包拒收。"""
    bad = []
    for p in SKILL_DIR.rglob("*"):
        if p.name.startswith(".") and "__pycache__" not in p.parts:
            bad.append(str(p.relative_to(SKILL_DIR)))
    if bad:
        return FAIL, ["发布目录里有 `.` 开头文件（skillhub CLI 会整包拒收）：" + ", ".join(bad)]
    return OK, ["发布目录内没有 `.` 开头文件"]


def check_local_copies() -> tuple[str, list[str]]:
    """⑤ 三个平台扫描目录是否与真源逐字节一致（复用 sync_copies.py 的副本地图）。"""
    sys.path.insert(0, str(ROOT))
    try:
        import sync_copies  # type: ignore
    except Exception as exc:  # noqa: BLE001
        return WARN, [f"import sync_copies 失败，跳过副本校验：{exc}"]
    src, rels = Path(sync_copies.SRC), sync_copies._src_files()
    msgs, status = [], OK
    for c in sync_copies.COPIES:
        dst = Path(c["path"])
        if not dst.is_dir():
            msgs.append(f"[{c['key']}] 目录不存在{'（平台强依赖）' if c['required'] else ''}：{dst}")
            status = FAIL if c["required"] else max(status, WARN)
            continue
        missing = [r for r in rels if not (dst / r).exists()]
        differs = [r for r in rels if (dst / r).exists() and sha_norm(src / r) != sha_norm(dst / r)]
        if missing or differs:
            msgs.append(f"[{c['key']}] 缺失 {len(missing)} / 不一致 {len(differs)}："
                        f"{(missing + differs)[:3]}")
            status = FAIL
        else:
            msgs.append(f"[{c['key']}] 与真源一致（{len(rels)} 个文件）")
    return status, msgs


def check_legacy_copies() -> tuple[str, list[str]]:
    """⑥ 盘上是否还留着旧 skill 副本（被误打包上架的风险）。"""
    hits = [(p, why) for p, why in LEGACY_SKILL_DIRS if Path(p).is_dir()]
    if not hits:
        return OK, ["未发现遗留旧副本"]
    msgs = [f"{p}  —— {why}" for p, why in hits]
    msgs.append("这些副本是远古格式（frontmatter 连 version/slug 都没有）。"
                "别从这些路径打包/发布；建议移到备份后删除，或改名加 _OLD 前缀。")
    return WARN, msgs


def check_repo_hygiene() -> tuple[str, list[str]]:
    """⑦ 仓库根是否混入本地工作目录（会被误提交 / 干扰 git status）。"""
    msgs, status = [], OK
    for name in (".workbuddy",):
        p = ROOT / name
        if not p.is_dir():
            continue
        rc, _ = git("check-ignore", "-q", name)
        if rc == 0:
            msgs.append(f"{name}/ 存在但已被 .gitignore 忽略")
        else:
            status = WARN
            msgs.append(f"{name}/ 存在且**没被忽略**（git status 会一直显示 `?? {name}/`，"
                        f"`git add -A` 会把本地工作记录提交到公开仓库）。建议写进 .gitignore。")
    if not msgs:
        msgs.append("仓库根没有混入本地工作目录")
    return status, msgs


CHECKS = [
    ("vendor/core ↔ 仓库 core/", check_vendor_vs_core),
    ("vendor/core 溯源（.upstream-tag）", check_vendor_provenance),
    ("三处版本号一致性", check_versions),
    ("发布包卫生（`.` 开头文件）", check_package_hygiene),
    ("本地三副本 ↔ 真源", check_local_copies),
    ("遗留旧副本扫描", check_legacy_copies),
    ("仓库根卫生", check_repo_hygiene),
]


def main() -> int:
    ap = argparse.ArgumentParser(description="excelrouter 多仓/多副本同步体检")
    ap.add_argument("--quiet", action="store_true", help="只打印 ⚠️/❌ 项")
    args = ap.parse_args()

    print(f"体检对象：{ROOT}")
    print("=" * 66)
    worst = OK
    for title, fn in CHECKS:
        try:
            status, msgs = fn()
        except Exception as exc:  # noqa: BLE001
            status, msgs = FAIL, [f"检查自身报错：{type(exc).__name__}: {exc}"]
        if args.quiet and status == OK:
            worst = max(worst, status, key=lambda s: [OK, WARN, FAIL].index(s))
            continue
        print(f"\n{ICON[status]} {title}")
        for m in msgs:
            print(f"    {m}")
        worst = max(worst, status, key=lambda s: [OK, WARN, FAIL].index(s))

    print("\n" + "=" * 66)
    if worst == OK:
        print("全部通过：5 条同步链路一致。")
    elif worst == WARN:
        print("有 ⚠️ 项需要人判断（多为设计使然，见上面说明）。")
    else:
        print("有 ❌ 硬错误，修完再发版。")
    return 1 if worst == FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
