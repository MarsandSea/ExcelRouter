#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""SKILL.md frontmatter 预检 —— 发布前拦掉「静默降级」的元数据问题。

为什么要这个脚本（2026-09-15 踩坑）：
  SkillHub CLI 的 frontmatter 解析是**手写的逐行解析**（只看 ``key: value``，不做 YAML 解析），
  它能容忍非法 YAML；但 **WorkBuddy 客户端 / 其他工具用真 YAML 解析**，一遇非法就抛异常。
  客户端的 catch 分支会把技能描述**降级成安装目录名** —— 于是技能卡片的描述
  显示成 ``excelrouter__skillhub`` 这种奇怪的东西（用户实际看到的）。

  触发条件极隐蔽：description 里出现一个 ASCII「冒号+空格」即可，
  例如「English triggers: split excel by column」—— 发布链路全程绿灯，坏在展示端。

本脚本检查（任一不过即退出码 1）：
  1. frontmatter 用 ``---`` 正确包裹；
  2. 值里没有 ASCII「冒号+空格」（YAML 非法，客户端会降级）；
  3. skillhub CLI 必填字段齐全：slug / version / displayName；
  4. 有 yaml 库时，额外做一次真正的 YAML 解析（与客户端一致）。

用法：
    python check_skillmd.py <skill 目录>
    python check_skillmd.py <skill 目录> --quiet

CI 里在发布前调用（见 .github/workflows/publish-skill.yml）。
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

REQUIRED_KEYS = ("slug", "version", "displayName")
_SLUG_RE = re.compile(r"^[a-z0-9][a-z0-9-]{1,127}$")
_SEMVER_RE = re.compile(r"^\d+\.\d+(\.\d+)?([-+].+)?$")
# ASCII 冒号 + 空格：YAML 里出现在 plain scalar 中会被判成新的映射键 → 解析失败
_COLON_SPACE_RE = re.compile(r":\s")


def split_frontmatter(text: str) -> str | None:
    """按 ``---`` 切出 frontmatter 文本（与客户端/CLI 同款行为）。"""
    if not text.startswith("---"):
        return None
    lines = text.split("\n")
    if lines[0].strip() != "---":
        return None
    for i in range(1, len(lines)):
        if lines[i].strip() == "---":
            return "\n".join(lines[1:i])
    return None


def line_parse(fm: str) -> dict[str, str]:
    """复刻 skillhub CLI 的逐行解析，确认 CLI 侧能读到什么。"""
    out: dict[str, str] = {}
    for raw in fm.split("\n"):
        line = raw.rstrip()
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        if ":" not in line:
            continue
        key, _, val = line.partition(":")
        key, val = key.strip(), val.strip()
        if key:
            out[key] = val.strip("'\"") if len(val) >= 2 and val[0] == val[-1] and val[0] in "'\"" else val
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="SKILL.md frontmatter 预检")
    ap.add_argument("skill_dir", help="skill 目录（含 SKILL.md）")
    ap.add_argument("--quiet", action="store_true", help="只报错误")
    args = ap.parse_args()

    skill_md = Path(args.skill_dir) / "SKILL.md"
    if not skill_md.is_file():
        print(f"::error::找不到 {skill_md}")
        return 1

    text = skill_md.read_text(encoding="utf-8")
    fm = split_frontmatter(text)
    if fm is None:
        print(f"::error::{skill_md} 的 frontmatter 格式异常（首行需为 ---，且有结束 ---）")
        return 1

    errors: list[str] = []
    meta = line_parse(fm)

    # ① 冒号+空格：最隐蔽、后果最重（客户端描述降级成目录名）
    #    注意：若值整体用双引号包起来（推荐写法），引号内的 ": " 是**合法**的
    #    （2026-09-17 起 SKILL.md 用这种写法），这时不再报错。
    for i, raw in enumerate(fm.split("\n"), 1):
        line = raw.rstrip()
        if not line.strip() or line.lstrip().startswith("#") or ":" not in line:
            continue
        key, _, val = line.partition(":")
        val_s = val.strip()
        if len(val_s) >= 2 and val_s[0] == '"' and val_s[-1] == '"':
            continue  # 已被引号包裹 → YAML 层面无歧义
        if _COLON_SPACE_RE.search(val):
            hit = _COLON_SPACE_RE.search(val)
            snippet = val[max(0, hit.start() - 30):hit.start() + 30].strip()
            errors.append(
                f"frontmatter 第 {i} 行 `{key.strip()}` 的值里含 ASCII「冒号+空格」：…{snippet}…\n"
                f"        这会让 WorkBuddy 客户端 YAML 解析失败，技能卡描述会降级成目录名。\n"
                f"        改法：整体用双引号包起来，或换成中文冒号「：」/ 破折号「—」。"
            )

    # ①b 描述类字段长度上限（mobilework 端硬限制 1024；桌面端无限制所以以前没暴露）
    #     2026-09-17 导入报「技能描述超过1024字符上限」才发现的。
    MAX_DESC = 1024
    for key in ("description", "description_zh", "summary"):
        v = meta.get(key) or ""
        # 引号算不算长度取决于平台实现，这里按「值本身」算，并留 5 字符余量提醒
        if len(v) > MAX_DESC:
            errors.append(
                f"`{key}` 长度 {len(v)} 超过 mobilework 的 {MAX_DESC} 字符上限"
                f"（超 {len(v) - MAX_DESC}）——导入会被拒绝。请精简。"
            )
        elif len(v) > MAX_DESC - 10:
            print(f"   ⚠️  {key} 长度 {len(v)}，非常接近 {MAX_DESC} 上限，建议再压一点")

    # ② CLI 必填字段
    for key in REQUIRED_KEYS:
        if not (meta.get(key) or "").strip():
            errors.append(f"frontmatter 缺 `{key}`：skillhub CLI 会直接拒绝发布")
    slug = (meta.get("slug") or "").strip()
    if slug and not _SLUG_RE.match(slug):
        errors.append(f"slug 不合法（需 kebab-case，2-128 字符）：{slug!r}")
    version = (meta.get("version") or "").strip()
    if version and not _SEMVER_RE.match(version):
        errors.append(f"version 不是合法 SemVer：{version!r}")

    # ③ 真 YAML 解析（客户端同款行为；没装 yaml 库就跳过，不当成失败）
    yaml_note = ""
    try:
        import yaml  # type: ignore

        try:
            parsed = yaml.safe_load(fm) or {}
            missing = [k for k in ("summary", "description") if not parsed.get(k)]
            if missing:
                yaml_note = f"（提示：{', '.join(missing)} 为空，卡片可能没描述可显示）"
        except Exception as exc:  # noqa: BLE001
            errors.append(
                f"YAML 解析失败（客户端会因此把描述降级成目录名）：{type(exc).__name__}: "
                f"{str(exc).splitlines()[0]}"
            )
    except ImportError:
        yaml_note = "（本机无 yaml 库，已跳过 YAML 解析校验）"

    if errors:
        print("❌ SKILL.md frontmatter 预检未通过：\n")
        for e in errors:
            print(f"  - {e}")
        print()
        return 1

    if not args.quiet:
        print("✅ SKILL.md frontmatter 预检通过")
        print(f"   slug        : {slug}")
        print(f"   version     : {version}")
        print(f"   displayName : {meta.get('displayName')}")
        for k in ("summary", "description_zh"):
            v = (meta.get(k) or "").strip()
            print(f"   {k:12s}: {(v[:70] + '…') if len(v) > 70 else v}{'' if v else '（空）'}")
        print(f"   description : {len(meta.get('description') or '')} 字符 {yaml_note}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
