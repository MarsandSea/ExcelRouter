<div align="center">

# ExcelRouter · Excel 智能拆分工具

**整个文件夹一键拆完：按部门、区域、工号等字段自动拆分，打包分发**

保留原格式 · 跨文件自动合并 · 单个文件也能拆 · 可同时拆到每个人

<br>

[![立即下载](https://img.shields.io/badge/%E2%AC%87%20%E7%AB%8B%E5%8D%B3%E4%B8%8B%E8%BD%BD-%E5%85%8D%E8%B4%B9%20%C2%B7%20%E5%85%8D%E5%AE%89%E8%A3%85%20%C2%B7%20Windows-1E7F4B?style=for-the-badge)](../../releases)

**[📖 3 分钟上手指引](docs/使用指引.md)** · **[❓ 常见问题](docs/FAQ.md)** · [下载哪个文件？](#-直接下载使用无需安装-python--download)

🔒 免费开源（MIT）· 数据仅在本机处理，不上传任何服务器 · 零编程，面向普通办公人员

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

*Batch-split a whole folder of Excel files by any field (department / region / ID), keep original
formatting, merge across files, optional per-person output — free, open-source, no coding required.*

作者 / Author：**AbeLin** · 觉得好用请点 ⭐ Star！

</div>

---

## 📌 一句话介绍 / What is this

**ExcelRouter 是一个 Windows 平台的开源 Excel 批量拆分工具**（MIT 协议，永久免费，无广告、
无注册、无功能限制）。它把**一个或一整批** Excel 表格，按你指定的**任意一个字段**
（部门 / 区域 / 门店 / 网格 / 工号 / 班级……）的取值自动拆成多个文件，**完整保留原有的复杂表头
与单元格格式**，可选再按第二个字段拆到每个人，最后按组打包成 ZIP 直接分发。
**全程在本机运行，不联网、不上传任何数据**，断网也能正常使用。另附 PDF 加密分发模式：
同一份 PDF 为每个网格生成专属打开密码 + 专属水印的副本。

**它不做这些**（避免你下载后才发现不合用）：

- ❌ 不做**表头扁平化 / 多层表头降维**——本工具的定位是「原样保留」表头，不是改造表头
- ❌ 不做**把多个表合并成一个总表**的通用合并（只在拆分时支持「同一取值跨文件合并」）
- ❌ 不做数据透视、清洗、统计分析
- ❌ **没有 Mac / Linux 版，也没有在线版**（在线版意味着要上传数据，与「数据不出本机」冲突）

> *ExcelRouter is a free, open-source (MIT) **Windows desktop tool that splits one or a whole folder
> of Excel files by the values of any column** — department, region, store, employee ID, class —
> keeping the original (even multi-row, merged) headers and cell formatting intact. Optional second
> level split (per person), automatic cross-file merge, ZIP packaging per group, plus a PDF secure
> distribution mode (per-group password + watermark). Everything runs **100% locally — no upload,
> no telemetry, works offline**. No macOS/Linux build, no web version; it does not flatten headers,
> merge workbooks into one, or do any analytics.*

---

## 🎯 谁会用到它 / Typical use cases

- 全公司销售 / 结算明细，**按部门或区域拆成多个文件**，分别发给各负责人
- 工资条、考核表，**一人一个独立文件**，避免任何人看到别人的数据
- 门店 / 网格 / 项目台账，**按门店编号拆分**并自动打包 ZIP 群发
- 学校成绩单，**按班级拆分**，班主任只拿到自己班的表
- 每月一份表、连着十几个月，**把同一个部门的数据跨文件合并成一张汇总表**
- 一份通报 PDF 要发给多个网格，**每个网格一个专属密码 + 专属水印**（泄露可溯源）

> 如果你正在搜的是这些问题 ——「怎么把一个 Excel 按某一列拆分成多个文件」「一个工作表拆成多个
> 工作簿」「批量拆分 Excel 不写 VBA / 不用 Power Query」「拆完还要保留原来的复杂表头和格式」
> 「按部门把表格分开发给不同的人」—— 那这个工具就是干这个的。

**国内用户下载更快**：[Gitee 镜像仓库](https://gitee.com/Marsandsea/Excelrouter/releases)
（与 GitHub 同步发版）· 国际用户：[GitHub Releases](../../releases)

---

## ✨ 功能特点 / Features

- **选一个字段就能拆** —— 自动识别表头，下拉选「拆分字段」，每个取值拆成一个文件，无需预先列出。
  *Pick one column and split — auto-detect header, choose a column, one file per value.*
- **三步卡片式界面** —— ①选表格 ②选字段 ③开始拆分，主按钮固定在底部；不常用的设置收进
  「▸ 高级设置」默认折叠，界面单屏不用切标签；支持把文件 / 文件夹**直接拖进窗口**。
  *Three-step card UI — pick table → pick field → split; advanced options collapse by default; drag & drop supported.*
- **拆分前先看结果（新）** —— 选定拆分字段后自动列出「将拆成哪几组」，可勾选只拆其中一部分；
  分组异常多时（比如误选了工号列）主动警告，不再导图。
  *Preview split groups before running, with optional per-value checkboxes and a warning for suspicious many-group fields.*
- **默认按原表拆分，按需跨文件合并** —— 默认每个源文件各自拆分、打包进同一个 ZIP；
  需要把同一个取值跨文件汇总成一张表时，「高级设置」里勾选「跨文件合并汇总」即可
  （修掉了旧版互相覆盖的问题）。
  *Splits each source file on its own by default, bundled into one ZIP; optionally merge the
  same value across files into a single table via Advanced Settings.*
- **只拆指定取值** —— ②里点「查看/勾选分组」勾掉不需要的即可（也可在「高级设置」手填），
  勾选结果自动记住，下次不用重设。
  *Restrict splitting to specific values via checkboxes, remembered across runs.*
- **二级拆分（按人分发）** —— 可选再按第二列细分（如 部门 → 姓名），一次产出汇总 + 到人双份结果。
  *Optional secondary split (e.g. Department → Person), producing both summary and per-person outputs.*
- **保留格式** —— 表头与数据行的字体、颜色、边框、数字格式、合并表头完整保留。
- **智能识别表头** —— 表头不在第一行也能自动找到（前 15 行启发式扫描）。
- **公式显示真实值** —— 读公式缓存值，并对「未计算的公式」提前预警。
- **可选保留公式（新）** —— 「高级设置」勾选「保留公式」，同一行内的公式（如「金额=单价×数量」）
  会平移行号后保留为活公式，方便接收人核对计算过程；跨行汇总/跨表的公式（如合计 SUM）会转成
  当前数值，不会出现算错的公式。默认关闭，主要面向单文件场景。
  *Optional formula preservation — same-row formulas stay live after splitting; cross-row totals
  are converted to their current value to avoid showing wrong results. Off by default.*
- **多 Sheet / 兼容 .xls / 批量递归** —— 一次处理整个文件夹。
- **大文件不假死** —— 几万行的单文件也有实时进度与心跳日志，界面全程响应（见 [FAQ](docs/FAQ.md#处理大文件几万行时界面卡住像死机了一样是不是崩溃了)）。
- **完成就知结果长什么样（新）** —— 拆分完成直接告诉你拆出几组、几个文件、多少行，
  哪个 sheet 被跳过也会点名，不用翻日志猜。
  *Clear result summary on completion — groups, files, rows, and anything skipped.*
- **PDF 加密分发（新）** —— 同一份 PDF 发给多个网格/部门：每个网格自动生成**专属打开密码 + 专属水印**
  （泄露可溯源）的副本，并输出「谁收什么、密码是什么」的分发清单，照着微信/邮件群发即可。
  *PDF secure distribution — per-group password-protected & watermarked copies from one PDF, plus a distribution manifest.*

> **关于格式保留的两点限制 / Limitations：**
> ① `.xls` 转换后无法保留原格式（仅保留数据）；
> ② 数据区的合并单元格暂不保留（表头的合并单元格正常保留）；
> ③ 跨文件合并按**列位置**追加，最适合「同一套模板的多个表」。
>
> 更多问题见 **[FAQ](docs/FAQ.md)**（杀毒软件误报怎么办、公式列为什么是空的、.xls 支持范围等）。

---

## 🖼️ 演示截图 / Screenshots

**数据源（3 行复杂表头——大标题 / 分组标签 / 列名，工具自动识别表头行）**

![数据源示例](docs/screenshot_source.jpg)

**拆分结果（按区域自动分组，每组含汇总 + 到人，整包 ZIP 可直接发负责人）**

![输出结果示例](docs/screenshot_output.jpg)

---

## 🚀 直接下载使用（无需安装 Python）/ Download

普通用户请直接下载打包好的程序：👉 **[前往 Releases 下载](../../releases)**

每个版本提供两种产物，**优先选 ZIP**：

| 产物 | 说明 | 适用场景 |
|---|---|---|
| `ExcelRouter-vX.X.X-win64.zip` | 文件夹形式，解压后双击里面的 exe | **推荐**，启动更快，极少触发杀毒软件误报 |
| `ExcelRouter-vX.X.X.exe` | 单文件版，下载即用无需解压 | 图方便，但个别杀毒软件可能误报（[why?](docs/FAQ.md#杀毒软件误报)） |

> 🤖 **在用 AI 助手（WorkBuddy / Claude / WPS 灵犀等）？** 也可以不装 exe：给助手装上
> **[ExcelRouter Skill](https://github.com/MarsandSea/excelrouter-skill)**——同一套核心代码的
> 命令行版，装好后对话里说一句「把这个表按部门拆到人」「给 PDF 按网格加密分发」就能直接干，
> 同样全程本机不上传。WorkBuddy 用户可在技能市场搜 `excelrouter` 一键安装
> （[SkillHub 页面](https://skillhub.cn/skills/excelrouter)）。普通办公用户建议继续用上面的图形界面版。

---

## 🧭 三步上手 / Quick Start

界面就是三张卡片，从上到下做完即可：

1. **①选择要拆的表格** —— 把文件 / 文件夹**拖进窗口**，或点「📄 选一个 Excel 文件」
   「📁 选整个文件夹（批量拆）」。
2. **②按哪个字段拆分** —— 在「拆分字段」下拉里选（如「部门」），下方自动列出
   「将拆成哪几组」，可点「查看/勾选分组」只拆一部分；选文件夹时还能勾「同时拆到人」。
3. **③开始拆分** —— 输出位置不用改（自动放进「拆分结果」文件夹），点「▶ 开始拆分」，
   完成后显示结果摘要（几组 / 几个文件 / 多少行），并自动打开输出目录。

不常用的设置（表头识别策略、只拆部分取值、取值归并、跳过值等）收在
**「▸ 高级设置（一般用不到）」**里，默认折叠，一般流程用不到点开它。

想试一下？仓库自带样本：

```bash
python examples/make_sample.py   # 生成 5 个月份的虚拟员工明细（1月A分公司明细.xlsx … 5月A分公司明细.xlsx）
```

用「📁 选整个文件夹」，输入选 `examples/`，拆分字段选「所属部门」，勾选「同时拆到人」，
（想体验跨文件合并汇总，在「高级设置」里勾上「跨文件合并汇总」），
点「▶ 开始拆分」——每个部门 5 个月的数据各自拆分（勾了合并则汇总成一张表），
同时产出按员工姓名细分的个人文件，整体打包成 ZIP 可直接发给对应负责人。

---

## 🔐 PDF 加密分发 / PDF secure distribution

同一份 PDF（通报、明细等）要发给多个网格/部门，又怕外泄？切到顶部的 **「PDF 加密分发」** 模式：

1. **①选择要分发的 PDF** —— 可多选（支持拖入），每个网格都会拿到全部所选文件。
2. **②选择密码映射清单** —— 一个 Excel 小表，每行一个网格：网格名、专属密码，可选接收人；
   没有清单点「生成模板」一键产出；没有密码思路可以不选密码列，开始时选择自动生成随机密码。
   选完自动识别列，确认「网格列 / 密码列 / 接收人列」三个下拉即可。
3. **③开始分发** —— 每个网格生成一份 **专属打开密码 + 专属水印**（网格名+日期，斜向平铺，
   泄露可溯源）的副本，并输出一份「分发清单.xlsx」（网格 | 文件 | 密码 | 接收人），照着群发即可。

> - 加密采用 **AES-256** 打开密码，忘记密码**无法找回**，请保管好清单；
> - **分发清单里有明文密码**，只留给分发人自己用，不要随文件一起发出去；
> - 源 PDF 本身带密码的暂不支持，请先解密另存。

---

## ⚖️ 和其它做法比 / Compared with other approaches

| 做法 | 批量处理多个文件 | 保留原格式 | 数据安全 | 上手成本 |
|---|---|---|---|---|
| 手动筛选 + 复制粘贴 | ❌ 一个个来 | ✅ | ✅ 本地 | 低，但极耗时易错 |
| 在线拆分网站 | 通常有限 | ⚠️ 常丢格式 | ❌ **要把表上传到别人服务器** | 低 |
| VBA 宏 | ✅ | 取决于代码 | ✅ 本地 | 高（要会写 / 改代码） |
| Power Query | ✅ | ❌ 输出是纯数据表 | ✅ 本地 | 中高（要学 M 语言、每次重配） |
| **ExcelRouter** | ✅ **整个文件夹一次拆完** | ✅ 含多层复杂表头 | ✅ **全程本机，可断网** | 低（三步点击，零编程） |

含敏感信息的报表（工资、结算、住户台账）尤其不建议走在线转换网站——**这也是本工具坚持不做
在线版、不做任何遥测上报的原因**。

---

## 💡 快问快答 / Quick FAQ

- **数据会被上传吗？安全吗？** —— 不会。全程本地处理，程序不联网、无任何遥测或自动上报，
  拔掉网线照样能用。源码开源，可自行审查。
- **下载的 exe 被杀毒软件报毒？** —— 是 PyInstaller 打包的**常见误报**，不是病毒。优先下
  ZIP 版基本不会触发，详见 [FAQ · 杀毒软件误报](docs/FAQ.md#杀毒软件误报)。
- **能把多层合并表头「拆平」成单行表头吗？** —— 不能。本工具的设计目标是**原样保留**复杂表头，
  不改造表头结构。
- **`.xls` 老格式支持吗？** —— 能拆、数据完整，但转换后**无法保留原格式**；格式重要就先用
  Excel 另存为 `.xlsx`。
- **免费吗？公司内部能用吗？** —— 免费，MIT 协议，个人和商业使用都可以，保留版权声明即可。
- **有 Mac 版 / 在线版吗？** —— 没有，只有 Windows 版。在线版意味着要把你的数据上传到别人
  的服务器，与「数据不出本机」冲突，不做；Mac / Linux 待 Star 数 100+ 后评估提供。
- **在用 AI 助手（WorkBuddy / Claude 等），可以不装 exe 吗？** —— 可以。有同一套内核的
  **ExcelRouter Skill** 版：装进你的 AI 助手后，对话里说一句「按部门拆分」「按 PDF 网格加密分发」
  就能完成，同样全程本机处理。见
  [excelrouter-skill](https://github.com/MarsandSea/excelrouter-skill)，WorkBuddy 用户可从
  [SkillHub](https://skillhub.cn/skills/excelrouter) 一键安装。

更多问题见 **[完整 FAQ](docs/FAQ.md)**。

---

## 🛠️ 开发者运行 / Run from source

```bash
git clone https://github.com/MarsandSea/excel-router.git
cd excel-router
pip install -r requirements.txt
python main.py
```

运行测试：

```bash
pip install pytest
pytest -q
```

---

## 📦 自行打包 / Build

直接运行 `build.bat`（onedir + onefile 双产物，已含必需参数），或参考
**[发版手册](docs/RELEASING.md)** 了解 CI 自动发版流程与防误报细节。

> ⚠️ `--collect-all customtkinter`、`--collect-all tkinterdnd2`、 `--add-data "config;config"` 三个参数缺一不可，
> 否则打包后的 exe 会启动崩溃、拖拽失效或找不到默认配置。

---

## 📂 项目结构 / Structure

```
excel-router/
├── main.py                  # 入口
├── config/default_config.json
├── core/
│   ├── splitter.py          # 核心拆分逻辑（表头识别 / 列枚举 / 跨文件合并）
│   ├── pdf_dist.py          # PDF 按网格加密分发（密码 / 水印 / 分发清单）
│   └── utils.py             # 文本清理 / 取值归并 / 文件名净化
├── gui/app.py               # 三步卡片式图形界面（Excel 拆分 / PDF 加密分发 双模式）
├── examples/make_sample.py  # 样本生成器
├── docs/                    # FAQ / 使用指引 / 发版手册 / 截图
├── llms.txt                 # 给 AI 助手与搜索引擎的项目事实卡
└── tests/                   # pytest 测试
```

---

## ❓ 遇到问题 / Support

先看 **[FAQ](docs/FAQ.md)**，多数问题（杀毒误报、公式列空白、.xls 限制等）都有解释。
没解决再提 [Issue](../../issues)，模板会引导你附上必要信息，处理更快。

**想提建议？** 程序内点「💬 反馈建议」可匿名反馈（1 分钟，不需要注册任何账号），
你的真实使用场景是这个工具迭代的主要依据。

---

## 📄 开源协议 / License

[MIT License](LICENSE)。可自由使用、修改、分发，请保留原作者署名。

## 👤 作者 / Author

**AbeLin** · 有问题欢迎提 [Issue](../../issues)
