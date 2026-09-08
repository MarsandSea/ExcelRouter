# excelrouter skill · 本地副本地图

> 一张表看清：哪份是真源、哪份被谁加载、哪份是垃圾。
> 日常只需要两条命令：**改真源 → `python sync_copies.py`**，其余目录不用再手工 cp。

## 一、真源（唯一）

| 路径 | 说明 |
|---|---|
| `G:\WorkSpace\excel-router\excelrouter-skill\skills\excelrouter\` | **唯一真源**。2026-09-08 起 skill 与桌面版 `core/` **同一个 git 仓库**（`MarsandSea/excel-router`），所有改动只改这里。原独立仓库 `MarsandSea/excelrouter-skill` 已归档 |

## 二、在用的副本（由脚本同步，不要手工改）

| 短名 | 路径 | 谁在加载 | 能删吗 |
|---|---|---|---|
| `workbuddy` | `C:\Users\Administrator\.workbuddy\skills\excelrouter__skillhub` | **WorkBuddy 实际加载这一份**（SkillHub 安装位） | ❌ 不能删，删了 WorkBuddy 就找不到这个 skill |
| `claude` | `C:\Users\Administrator\.claude\skills\excelrouter` | Claude Code | ❌ 保留（你指定保留） |
| `wps` | `C:\Users\Administrator\AppData\Roaming\WPS 灵犀\serverdir\user_skills\excelrouter-skill-pkg` | WPS 灵犀 | ⚠️ 平台扫描位，删了 WPS 侧会失效 |

> ⚠️ `workbuddy` 那份目录里有平台自有的 `_meta.json` / `_icon.png` / `_skillhub_meta.json`，
> 它们**不在真源里**。同步脚本只做覆盖式复制、不删任何文件，所以这几个文件不会被冲掉。
> 别用"先删再拷"的方式手工同步，会把它们删掉导致平台认不出 skill。

## 三、统一同步脚本

```bash
python sync_copies.py             # 同步全部副本
python sync_copies.py --dry-run   # 只看会同步什么
python sync_copies.py --check     # 校验各副本与真源是否一致（不同步）
python sync_copies.py --only claude   # 只同步某一分
```

当前校验结果（2026-09-04）：三份副本均与真源一致，19 个文件。

打上架包（同步之后再做）：

```bash
python package_excelrouter.py     # 产物 excelrouter-skill-v2.7.1.zip
```

## 四、散落/历史目录（清理与否由你决定，**我未删除**）

清理前已做完整备份，可随时恢复，见下节。

| 目录 | 文件数 / 体积 | 状态 | 建议 |
|---|---|---|---|
| `G:\WorkBuddy\excel-router\legacy` | 132 / 877 KB | 已单独备份 | 可删（v2.6.x zip、交接文档、logo 均在备份里） |
| `G:\WorkBuddy\excel-router\_smoke` | 199 / 1.7 MB | 已备份 | 可删（冒烟测试残留 out1-out7 / r1-r7 / verify） |
| `C:\Users\Administrator\excelrouter-skill-tmp` | 66 / 464 KB | 已备份 | 可删（临时目录，名字带 tmp） |
| `C:\Users\Administrator\WorkBuddy\2026-08-17-14-57-57\excelrouter-repo` | 151 / 694 KB | 未备份（无需） | 可删（HEAD `40b2dae` 与工作区仓库完全一致，无未迁移历史） |

要手动清理就把上面的路径直接删掉即可。

## 五、备份文件（可恢复）

| 文件 | 内容 | 体积 |
|---|---|---|
| `legacy-archive-20260904.zip` | **legacy 专属备份**，132 文件（含 3 个 v2.6.x zip、EXCELROUTER_HANDOFF.md、logo） | 479 KB |
| `_archive-before-cleanup-20260904.zip` | 合并备份：legacy + _smoke + skill-tmp，282 文件 | 1347 KB |

两份都已通过完整性校验：CRC 全条目可读，抽样恢复后与原件**逐字节一致**。

恢复示例：

```bash
# 只恢复 legacy
python -c "import zipfile;zipfile.ZipFile('legacy-archive-20260904.zip').extractall('.')"
```
