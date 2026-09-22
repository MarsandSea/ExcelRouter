# excelrouter skill · 本地副本地图

> 一张表看清：哪份是真源、哪份被谁加载、哪份是垃圾。
> 日常只需要两条命令：**改真源 → `python sync_copies.py`**，其余目录不用再手工 cp。

## 一、真源（唯一）

| 路径 | 说明 |
|---|---|
| `G:\WorkSpace\excel-router\excelrouter-skill\skills\excelrouter\` | **唯一真源**。2026-09-08 起 skill 与桌面版 `core/` **同一个 git 仓库**（`MarsandSea/ExcelRouter`），所有改动只改这里。原独立仓库 `MarsandSea/excelrouter-skill` 已归档 |

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

> 发版前建议再跑一次 `python check_sync.py`：它把**五条同步链路**（仓库 core ↔ vendor/core、
> vendor 溯源、三处版本号、发布包卫生、三副本、遗留旧副本、仓库根卫生）一次体检完，
> 有硬错误时退出码 1。见本文件最后一节。

当前校验结果（2026-09-15）：三份副本均与真源一致，**20 个文件**。

打上架包（同步之后再做）：

```bash
python package_excelrouter.py     # 产物 excelrouter-skill-v2.7.1.zip
```

## 四、散落/历史目录（清理与否由你决定，**我未删除**）

清理前已做完整备份，可随时恢复，见下节。

> ⚠️ **2026-09-15 复核：这些目录目前都还在**（下表体积为当日实测）。
> 它们里面**装着远古版本的 skill**（`skills/excelrouter/` 的 frontmatter 连 `version` / `slug` 都没有）。
> 危害不是占地方，而是**被误当作真源打包上架** —— 老仓的 cron 就这么干过一次，
> 用旧内容覆盖了 SkillHub 上的新版本。所以：**永远不要从这些路径打包或发布**。
> `check_sync.py` 会把它们作为「遗留旧副本」列出来提醒。

| 目录 | 文件数 / 体积 | 状态 | 建议 |
|---|---|---|---|
| `G:\WorkSpace\Excelrouter-skill`（**未列入旧表，09-15 新发现**） | 含 `skills/excelrouter/` | **未备份**，且是**旧真源路径** | ⚠️ 优先处理：与真源同在前缀 `G:\WorkSpace` 下、只差大小写与连字符，人/脚本都极易选错 |
| `G:\WorkBuddy\excel-router\legacy` | 132 / 617 KB | 已单独备份 | 可删（v2.6.x zip、交接文档、logo 均在备份里） |
| `G:\WorkBuddy\excel-router\_smoke` | 199 / 999 KB | 已备份 | 可删（冒烟测试残留 out1-out7 / r1-r7 / verify） |
| `C:\Users\Administrator\excelrouter-skill-tmp` | 66 / 341 KB | 已备份 | 可删（临时目录，名字带 tmp） |
| `C:\Users\Administrator\WorkBuddy\2026-08-17-14-57-57\excelrouter-repo` | 151 / 453 KB | 未备份（无需） | 可删（HEAD `40b2dae` 与工作区仓库完全一致，无未迁移历史） |

要手动清理就把上面的路径直接删掉即可 —— **删除前请先确认要删的是旧副本，不是真源**
（真源是 `G:\WorkSpace\excel-router\excelrouter-skill\skills\excelrouter`，只有这一个）。

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

## 六、同步体检：`check_sync.py`

skill 现在活在**五条链路**上，任意一条断了都不报错、只是悄悄不一致（最难查）。
`sync_copies.py` 只管第 3 条（平台扫描目录），其余四条由 `check_sync.py` 一起体检：

| # | 链路 | 体检内容 | 硬错误示例 |
|---|---|---|---|
| 1 | 仓库 `core/` → skill `scripts/vendor/core/` | 逐模块比内容（**忽略行尾**） | 有人只改了 vendor、没改 core |
| 2 | `vendor/core` → `.upstream-tag` | vendor 是否真等于该 tag 的 `core/` | 手动发布时顺手同步了 core |
| 3 | 平台扫描目录（3 份） | 复用 `sync_copies` 的副本地图 | 副本落后一版 |
| 4 | 发布包 | 目录里是否有 `.` 开头文件 | skillhub CLI 整包拒收 |
| 5 | 盘上的旧副本 / 仓库根 | 遗留远古 skill、`.workbuddy/` 是否被忽略 | 从旧路径打包上架 |

```bash
python check_sync.py            # 全部体检，有 ❌ 时退出码 1（可进 CI）
python check_sync.py --quiet    # 只打印 ⚠️/❌
```

> 行尾差异（CRLF/LF）**不算不一致**：本机 `core.autocrlf=true`，`core/*.py` 提交库里是 LF、
> `vendor/core/*.py` 工作区是 CRLF，逐字节比会把整个文件报成差异。这也是 2026-09-15 踩过的坑。
