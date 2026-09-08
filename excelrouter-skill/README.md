<div align="center">

<img src="assets/logo-er.png" alt="ExcelRouter Logo" width="120">

# 📊 ExcelRouter · 表格拆分 & PDF 加密分发

**跟 AI 说一句话，拆表、加密 PDF 全自动完成** — 不用装软件，不用懂编程

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)
[![SkillHub](https://img.shields.io/badge/SkillHub-已上架-1a73e8)](https://skillhub.cn/skills/excelrouter)

*"帮我把这张工资表按部门拆开，每个部门一个文件"——剩下的交给 AI。*
*数据全程在你电脑上处理，不联网、不上传。*

作者：**AbeLin** · MIT 开源协议 · ⭐ 觉得好用请点个 Star

</div>

---

## 🎯 这能帮你做什么

你是不是经常遇到这些活儿？

- 📋 **一张几百人的工资表 / 考勤表**，要按部门拆成单独文件，发给各部门负责人
- 🔒 **一批 PDF 合同 / 报告**，要分别加密、打上专属水印，再分发给不同的人
- 👤 **同一个人的数据散落在好几个文件里**，要合并成一份个人档案

以前要么手动复制粘贴到手酸，要么找会写代码的人帮忙。**ExcelRouter 让你直接对 AI 说一句话，它就帮你搞定。**

---

## 👥 谁适合用

| 角色 | 典型场景 |
|---|---|
| 🧑‍💼 **HR / 行政** | 工资表按部门拆、考勤表按人拆、通知按区域拆 |
| 💰 **财务** | 凭证按月份拆、报表按部门拆、PDF 合同加密分发 |
| 📊 **运营 / 销售** | 客户名单按区域拆、业绩表按团队拆、报告 PDF 加密分发 |
| 👩‍🏫 **老师 / 班主任** | 成绩表按班级拆、按学生拆到每人一份 |

**不会编程？没关系。** 你只需要会打字、会描述需求，AI 来干脏活累活。

---

## ✨ 三大能力

### 1️⃣ Excel 按字段拆分

按任意一列把大表拆成多个小文件：部门、区域、工号、班级、姓名……随便选。

- ✅ **保留原格式**：字体、颜色、列宽、合并单元格，拆完跟原表一样
- ✅ **拆到个人**：先按部门拆，再按人拆到每人一个文件
- ✅ **跨文件合并**：同一个人在多个文件里的数据，自动合并成一份总表
- ✅ **智能识别**：AI 会先看一眼有哪些字段、取值，跟你确认没拆错再动手

### 2️⃣ PDF 批量加密分发

给一批 PDF 分别加上密码、打上专属水印，再生成一份密码清单方便你逐人通知。

- ✅ **银行级加密**（AES-256），没有密码打不开
- ✅ **专属水印**：自动写上接收人姓名 / 日期，防截图外传
- ✅ **密码清单**：输出一份「分发清单.xlsx」，记录每个人对应的密码

> ⚠️ **重要提醒**：密码清单含明文密码，只给你自己留底，
> **千万不要和加密后的 PDF 一起发出去**！

### 3️⃣ 数据全程本机处理

所有文件都在**你自己的电脑上**处理完成：

- 🚫 不联网、不上传任何数据
- 🚫 不留痕、不发送到任何服务器
- ✅ 适合处理工资、合同、客户信息等敏感数据
- ✅ 代码完全开源，可逐行审查

---

## 📸 效果预览

> 📷 截图准备中，即将补充到 `docs/screenshots/`
>
> 1. 一句话拆分前 → 拆分后按部门生成的文件列表
> 2. 打开拆分结果，展示格式（字体/颜色/列宽）原样保留
> 3. PDF 加密分发前后对比 + 密码分发清单

---

## 🚀 怎么装

### 方式一：WorkBuddy 一键安装（最简单，推荐）

已上架 SkillHub（[skillhub.cn/skills/excelrouter](https://skillhub.cn/skills/excelrouter)），在 WorkBuddy 里直接搜索安装：

> **WorkBuddy → 左侧「技能 / 专家」→ 搜索「ExcelRouter」→ 安装**

就这一步，装完就能用。

### 方式二：Claude Code 安装

在 Claude Code 里依次输入两行：

```
/plugin marketplace add MarsandSea/excelrouter-skill
/plugin install excelrouter
```

### 方式三：手动安装（给技术小伙伴）

<details>
<summary>点开查看命令行安装方式</summary>

```bash
git clone https://github.com/MarsandSea/excelrouter-skill.git
cp -r excelrouter-skill/skills/excelrouter ~/.claude/skills/      # Claude Code
# 或
cp -r excelrouter-skill/skills/excelrouter ~/.workbuddy/skills/   # WorkBuddy
```

第一次使用需要装一下 Python 依赖（装一次就行）：

```bash
pip install -r skills/excelrouter/requirements.txt
```

> 只用 Excel 拆分的话，装 `pandas` / `openpyxl` / `xlrd` 就够；
> 要用 PDF 加密分发，再装上 `pypdf` / `fpdf2` / `cryptography`
> （缺 `cryptography` 会自动降级，功能仍可用，只是加密强度低一些）。

</details>

---

## 💬 怎么用

装好之后，直接像聊天一样对 AI 说话就行。

### 拆 Excel

把文件路径发给 AI，说清楚按什么拆：

```
帮我把这份工资表按部门拆开，每个部门一个文件
```

```
这一批月度报表，按区域拆，顺便按姓名也拆到人
```

```
只要东区和西区的数据，其它区域不用管
```

AI 会**先自动看一眼**有哪些字段、每个字段有哪些取值，**跟你确认没拆错**
（比如问你"销售部"和"销售部门"算不算同一个部门），确认后才动手拆，
完成后告诉你文件生成在哪个文件夹。

### PDF 加密分发

先准备一张「网格→密码」Excel 清单（一行一个人，至少要有"密码"这一列），然后说：

```
把这几个 PDF 按这张密码表批量加密，每个人一个专属密码，再加个水印
```

AI 会确认哪列是密码、哪列是接收人，然后输出加密后的 PDF + 带水印 + 密码清单。

### 更多说法（随口说就行）

| 你想干嘛 | 就这么说 |
|---|---|
| 同一个人在好几个文件里 | "把同一个人的数据合并成一张总表" |
| 取值写法不统一 | "有的写'销售部'有的写'销售部门'，是同一个部门" |
| 表头不在第一行 | "表头在第三行，别认错了" |
| 保留计算公式 | "接收人要能看到金额是怎么算出来的" |
| 先看看不真拆 | "先看看会拆出哪些文件，不要真的写盘" |
| 自定义水印 | "水印写'内部资料·请勿外传'" |

---

## 🛡️ 我的数据安全吗

**放心，非常安全。**

- 🔒 所有文件都在你自己的电脑上处理，**全程不联网**
- 🔒 密码只用来加密 PDF，不会上传、不会记录到日志
- 🔒 代码完全开源（MIT 协议），任何人都能审查每一行代码
- 🔒 用的都是知名开源库，没有来路不明的依赖

---

## ❓ 常见问题

**要钱吗？**
免费，MIT 开源，随便用。

**我不会编程能用吗？**
能。你只需要会描述需求，比如"按部门拆开"——AI 来干技术活。

**支持哪些 Excel 格式？**
`.xlsx` 完美支持（保留所有格式）；`.xls`（旧格式）也能用，但样式可能丢失，建议先另存为 `.xlsx`。

**数据区有合并单元格怎么办？**
表头区的合并单元格会保留；数据区的不支持，建议源表数据区尽量不要合并单元格。

**水印里的中文变成问号？**
Windows 一般会自动找到中文字体；如果没找到，可以指定一个中文字体文件（`.ttf` 格式）。

**和 ExcelRouter 桌面版什么关系？**
本仓是 [excel-router](https://github.com/MarsandSea/excel-router) 桌面工具核心能力的 AI 版移植——同一套引擎，但不用装软件、不用开界面，对 AI 说一句话就行。想要图形界面给同事用，去下载桌面版 exe。

---

## 📣 觉得好用？

- ⭐ 给本仓库点个 Star，让更多人看到
- 🐛 遇到问题提 Issue（附上报错信息 + 最小复现文件）
- 🤝 欢迎 PR
- 💬 也欢迎分享给身边经常被表格折磨的同事

---

<details>
<summary><b>🔧 给开发者（目录结构 / 同步机制 / 开发指南）</b></summary>

### 目录结构

```
excelrouter-skill/
├── .claude-plugin/           # 插件市场清单（marketplace.json + plugin.json）
├── skills/excelrouter/
│   ├── SKILL.md              # skill 主文件，触发词/工作流/常见陷阱
│   ├── manifest.yaml         # WorkBuddy SkillHub 市场元信息
│   ├── requirements.txt
│   ├── scripts/
│   │   ├── _common.py        # 公共工具（sys.path 引导、默认配置、JSON 输出约定）
│   │   ├── er_inspect.py     # 看表：字段名/取值/表头行
│   │   ├── er_split.py       # Excel 拆分
│   │   ├── er_pdf_dist.py    # PDF 按网格加密分发
│   │   └── vendor/core/      # ← 上游 core/ 的逐字副本，自动同步，勿手改
│   └── references/           # 参数手册 / 场景配方 / 排查手册
├── tests/test_cli.py         # 三个脚本的冒烟测试
└── .github/workflows/
    ├── sync-upstream.yml     # 拉取式自动同步
    └── ci.yml                # push/PR 跑测试
```

### 版本如何跟上游同步

`skills/excelrouter/scripts/vendor/core/` 是
[excel-router](https://github.com/MarsandSea/excel-router) 仓库 `core/` 目录的逐字副本，
由 [`.github/workflows/sync-upstream.yml`](.github/workflows/sync-upstream.yml) 自动维护：

- 每天定时检查上游最新的 `v*` 发布 tag，有新版本就拉取、跑测试、测试通过才提交并打同名 tag。
- 也可以去本仓 Actions 页手动点 **Run workflow** 立即同步。
- 只跟发布 tag 走，不跟上游 `main` 分支的半成品。

当前对应的上游版本、commit、同步时间见
[`skills/excelrouter/scripts/vendor/UPSTREAM.md`](skills/excelrouter/scripts/vendor/UPSTREAM.md)。
**`vendor/` 目录下的文件不要手改**，下次同步会直接覆盖。

### 开发

```bash
pip install -r skills/excelrouter/requirements.txt
pip install pytest ruff
pytest -q
ruff check .
```

### 协议

MIT License，与上游 [excel-router](https://github.com/MarsandSea/excel-router) 一致。
`scripts/vendor/core/` 下每个文件保留原始版权头，请勿删除。

</details>
