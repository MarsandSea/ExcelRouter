# CLAUDE.md

本文件是给 Claude Code 的项目说明与工作契约。请在每次开始工作前阅读。

---

## 项目是什么

**ExcelRouter**：一个桌面工具（Windows + 国产化 Linux：银河麒麟 V10 / 统信 UOS），把一批 Excel
表格，**按用户指定的任意一个字段的取值**拆分成多个文件（每个取值一个文件），可选再按第二个
字段做二级拆分（「到人」），并保留原始表头格式。面向不会编程的普通办公人员，v2.1 起界面已改为
**单屏自适应**（不再是三层分档 Tab），v2.4 后期进一步重构为**三步卡片式**：
①选表格 → ②选字段 → ③开始拆分，主按钮固定底部，高级设置与日志默认折叠。
当前版本 **v2.9.0**，GitHub 仓库：`MarsandSea/ExcelRouter`。

> **v2.9.0：速度模式显性化 + 取值预览不再把抽样当全量（2026-09）**。
> 两条都来自作者在真实工单批次（401 个文件 / 29 个分组）上的实际使用反馈。
>
> ① **「拆分速度」段选从「高级设置」提到第 ③ 步主操作区**（`SPEED_STD` /
> `SPEED_FAST` 两档，直接写明代价）。`preserve_format` 是唯一能把耗时拉开数倍的
> 选项，藏在折叠面板里等于不存在——用户只会得出「这工具拆得慢」的结论，而不会
> 去翻高级设置找到它。**高级设置里原来的「保留格式」勾选框已删除**：同一个开关
> 两处控制，用户永远不确定哪个生效。`_on_speed_change` 顺带联动「保留公式」
> （极速模式下它无从谈起，直接置灰并取消勾选，而不是留一个勾了也不生效的框）。
>
> ② **取值预览的诚实性修复（核心）**。`_scan_input` 挑的「模板」只是文件夹里
> 排序最靠前的**一个**样本表，`list_values` 只枚举它，但**拆分是对全部文件生效的**。
> 旧文案直说「将拆成 6 组」，用户按它理解、跑完看到 29 组，会认为是工具算错了。
> **数字本身没错，错的是把抽样结论说成了全量结论。** 现在：
> 摘要行改成「样本「x.xlsx」里有 N 组 … ⚠ 这只是 M 个表中的 1 个」（警告色）；
> 新增 `_scan_all_values()`「🔍 扫描全部文件的分组」按钮（读遍所有表取并集，
> **显式按钮而非自动跑**——400 个表自动扫会让「选字段」这一步凭空卡几分钟，
> >30 个表时先确认）；新增 `_pick_template()`「📄 换样本表」让用户自己指定
> 有代表性的那张表（自动挑的那个未必有代表性）；`_start` 里补一道
> `askyesnocancel`，不筛选取值又处于抽样态时给出「照这样拆 / 先列全分组 / 取消」
> 三选一。另外 `_lock_to_listed()`「只拆这 N 组」按钮修掉一个语义陷阱：
> 「全选＝留空＝不筛选」在抽样场景下会咬人（样本恰好就是用户想要的那几组时，
> 全勾等于什么都没筛），这个按钮把「我只要这几组」写成一份**显式名单**。
> `_values_scope`（sample / all）是这一套文案的开关，**别让它退回只有一种口径**。
>
> ③ **「已限定」必须是显式状态 `_values_locked`，不能靠「是否全勾」反推**。
> 反推版实测会这样漏：点完「只拆这 N 组」，再随手取消一个分组又勾回来，
> `_sync_selvals` 看到「全勾」就把名单清成空串（＝拆全部），界面上的
> 「🔒 只拆其中 N 组」角标也一起没了——用户刚点过限定，不会再去确认一次。
> 现在的规则：取消勾选任何一组 → 进入已限定；全不选 / 点「全选」→ 解除限定；
> 已限定时即使全勾也照样写出名单。面板里的按钮与提示由 `_refresh_panel_hint()`
> 跟随状态刷新（面板是一次性搭起来的，状态却会在它开着时变）。
> 同时 `_set_values` 里补了一道：换字段后旧名单与新字段一个都对不上时**清空并
> 在日志里点名**——留着它会让这一跑按不存在的取值筛选、一行都拆不出来，
> 而界面还挂着「🔒 只拆其中 N 组」。
>
> **v2.8.0：麒麟（信创）适配（2026-09）**。公司配发的新机器有相当一批是移动信创版
> 银河麒麟 V10，跑不了 win64 产物。**单一代码库、不开分支**，平台差异集中到新模块
> `gui/platform_compat.py`（硬性规则：**该模块不得在模块级 import tkinter/customtkinter**，
> 否则新单测就需要显示器，当前「pytest 可无头运行」的性质会丢）。六处改动：
> ① **`TkinterDnD._require()` 补 try/except** —— 这是唯一的启动级崩溃点，tkdnd 原生库
> ABI 对不上时旧代码会把 `App.__init__` 打穿、进程直接退出。现在拆成两个标志：
> 模块级 `_DND_OK`（包 import 成功）与实例级 `self._dnd_ok`（原生库真的加载成功），
> 三个消费点都改读后者。**别把这个 try 去掉。**
> ② **`open_folder()`** 替掉两处 `subprocess.Popen(f'explorer "{p}"')` —— 裸字符串在
> POSIX 上被当成单个可执行文件名，且失败被 `except: pass` 吞掉，表现为死按钮。
> 新实现 Windows 也改成 list 形式，Linux 走 `xdg-open → gio → peony(UKUI) → ...` 回退链，
> **返回 bool，调用方必须给可见提示**。
> ③ **`_find_cjk_font()` 支持 Linux**：环境变量 `ER_CJK_FONT` → 扫字体目录按候选名匹配
> → `fc-match`（带 timeout=3，冷缓存会卡）。新增 `_font_usable()` 让 fpdf2 自己回答能不能
> 加载，不靠扩展名猜。顺手给 `_make_watermark_pdf` 的 `add_font` 加了守卫——**既有 bug**：
> 一个坏字体文件会在已经写出一部分网格之后把整轮分发打断。
> ④ 拖放载荷解析补 `file://` URI 解码（X11/XDND 发的是百分号编码 URI，旧正则解不开，
> 麒麟上**每一次拖放都会被拒**）。⑤ UI 字体候选加 Noto CJK / 文泉驿 / 方正。
> ⑥ 窗口图标补 `iconphoto(app.png)`（X11 的 `iconbitmap` 只认 XBM）。
> 打包：新增 `build_linux.sh`（**在目标机器上跑**）+ `packaging/launcher.sh.in`，
> 只产 onedir tar.gz。**麒麟版按需发布、不跟随 tag 自动跑**
> （`release-linux.yml` 只有 `workflow_dispatch`）——真机验收有成本，
> 发一份没验过的包比版本落后更糟；判断标准与操作见 `docs/RELEASING.md` §1.7。
> 另修一个平台无关的真 bug：`core/splitter.py` 的 `__tmp__`
> 过滤原来大小写敏感，`X__TMP__.XLSX` 会被当数据重新吃进去（回归用例已补）。

> **v2.7.2：修「到人」拆分名实不符 bug（2026-09）**。现象：同人跨多个源文件（7/8/9 月表都有
> 张三）时最终只有一张「张三_7月明细.xlsx」且里面装着全部月份的行——数据没丢，但文件名只带
> 第一个源文件，用户按名理解以为其他月份被覆盖丢了。旧实现：`_person_key_path` 键不含
> src_stem（无条件跨文件合并），路径却带 src_stem（只采用第一个源的名字）。第一性原理修复：
> 文件名必须诚实反映内容，与『汇总』的 merge 语义对称——**merge=True**：键
> `('到人', primary, person)`，路径 `{姓名}.xlsx`（同人的全部行进这一张，名字不带源文件）；
> **merge=False**：键 `('到人', primary, person, src_stem)`，路径 `{姓名}_{原文件名}.xlsx`
>（每人每源一张，与汇总的按原表拆分对称，也避免跨源按列位置合并错位）。
> 输出契约变化：merge 模式到人文件名去掉源文件后缀。测试补了两个跨多源回归用例
> （`test_to_person_cross_files_merge/no_merge`）——旧测试每人只出现在单个源文件，
> 这正是 bug 逃逸的原因，新用例必须保留。
>
> 同版修复 issue #1（fill_random_passwords）：`password_col` 未指定且清单已有「密码」列时
> 追加重复同名列，且 `read_mapping` 的 `header.index` 命中原空白列→生成的密码全部漏读。
> 修法：未指定时先复用既有「密码」列（只填空白格），没有才新增。回归用例
> `test_fill_random_passwords_reuse_existing_blank_col`。**注意：excelrouter skill 侧的
> 规避补丁在上游修复后可删**（skill 仓独立维护，删补丁需另跑 skill 侧测试）。

> **v2.7.1：界面品牌化精修（2026-09）**。零新依赖零体积增加，不引入 theme.json（自定义主题
> JSON 缺键有崩溃风险），全部走**代码内调色板常量 + 组件参数覆盖**：品牌主色「移动蓝」
> `PRIMARY=("#0E7FD1","#38BDF8")`（浅/深双主题取值集中在常量区，`ACCENT` 已是它的别名），
> 用于主按钮/步骤徽章/进度条/分段选择高亮/Hero 品牌色块；步骤卡片 corner_radius=12 + 1px
> 淡边框 `CARD_BORDER`；主按钮加高到 44/圆角 10；完成摘要改由浅绿横幅 `OK_BANNER_BG`
> 承载（有警告时退中性底色）。**字体统一的两个关键事实（实测）**：CTkFont 未指定 family 时
> 取 `ThemeManager.theme["CTkFont"]["family"]`（改 `FontManager._default_font` 无效），且组件
> 构造时读主题值——所以 `_init_fonts()` 必须在 root 建立后、任何组件创建前调用（现在在
> `App.__init__` 开头），生效条件是 root 存在（`tkfont.families()` 要 root）。新增页脚「🌓
> 深浅色」循环按钮（config 键 `appearance_mode`: system/light/dark，即时生效）。验收方式：
> PIL ImageGrab 三态截图（浅/深/完成）目检对比度，别只看代码。

> **v2.7：普通用户全流程 UX 升级（2026-09）**。站在「第一次打开 exe 的办公人员」角度补四个断点：
> ① **取值预览**：选定拆分字段后自动枚举该列取值（接入 `core.splitter.list_values`，此前只有
> CLI 契约、GUI 从未接入），②卡片直接显示「将拆成 N 组」，可展开勾选只拆部分取值
> （勾选结果写回 `selected_values`，全选/全不选=留空=拆全部）；N>50 橙色警告「可能选错字段」
> 并在开始时二次确认。② **拖拽**：新增 tkinterdnd2 可选依赖（`App` 变为
> `CTk + TkinterDnD.DnDWrapper` 混编，import 失败静默退回无拖拽）——表格/文件夹拖进
> ①卡片即可开始，PDF 和映射清单同理；打包三处都要加 `--collect-all tkinterdnd2`。
> ③ **小屏适配**：窗口高度按屏幕自适应（1366×768 老笔记本下 800px 会出屏）+ 记住上次
> 窗口大小位置（config 新键 `window_geometry`）；页脚新增「大字号」切换（`ui_scale`，
> 重启生效）。④ **完成摘要**：core 两个 run_* 收尾各打一行 `[SUMMARY] {json}`（机器可读，
> 追加式不改返回值/签名），GUI 渲染「拆出 N 组 · M 文件 · K 行 · Z 个 ZIP + 跳过/失败点名
> + 输出路径」多行摘要；日志里 ⚠ 行计数给「处理详情」打角标，失败时出现「📋 复制日志」。
> 其余：扫描结果带文件数+总 MB（>50 个或 >200MB 附耗时长提示）；开跑前检测 `~$` 锁文件
> 提醒「正被 Excel 打开读到的是旧内容」；「🔄 扫描字段」改名「🔄 重新识别」（自动扫描已是
> 主路径，按钮只是兜底）；「跨文件合并汇总」旁补内存风险提示。PDF 模式：②卡片加
> 「生成模板」一键产出映射清单模板（说明放第二个 sheet，不污染数据区）；不选密码列
> 开始时可选**自动生成随机密码**（新函数 `fill_random_passwords`，只填空白、不改原文件、
> 另存 `*_含密码.xlsx`）；选完清单显示「N 个网格 × M 个 PDF = K 个加密副本」预览。

> v1.0 原是「网格化管理」专用（按网格 + 工号识别）。v2.0 已**通用化**：表头自动识别、
> 按列名选列、自动枚举取值、跨文件合并。作者原来的网格 + 岗位工作流用「专家模式」
> （header_mode=keyword + 主列=网格 + 二级列=岗位 + value_alias_map 复现岗位归并）即可复现。

**作者：AbeLin。开源协议 MIT。目标是开源传播 + 建立个人声誉，不收费。**
所有源码文件顶部都有 `Copyright (c) 2026 AbeLin · MIT License` 署名，**请勿删除任何文件的版权头**。

> **品牌名变更（2026-07，v2.2）**：项目对外品牌名从「Excel 通用拆分工具 / Excel Splitter」
> 改为 **ExcelRouter**。GitHub 仓库已从 `excel-splitter` rename 为 `excel-router`
> （GitHub 自动保留旧链接 301 跳转）。当时只改了**用户可见的品牌文案**：窗口标题、
> `version.txt` 元数据、`build.bat` 产物名、README。旧版本号 v2.1 → 新版本号 v2.2。
> **2026-07-03 补充**：内部残留的旧代号也已清理——本地项目目录从 `Excel_splitter` 改名为
> `excel-router`（与 GitHub 仓库名一致），文档目录树同步更新。`core/splitter.py`、`gui/` 等
> 文件/包名描述的是**功能**（拆分、界面）而非旧品牌，**保持不变**，不要为了品牌一致去改它们。
>
> **品牌文案定稿（2026-07-03，v2.4 后）**：采用「品牌名 + 直白副标 + 差异化标语」三层结构，
> 中文副标从「Excel 业务数据自动分发工具」改为「Excel 批量拆分工具」（新人 3 秒看懂 +
> 保留"Excel/拆分"搜索词；Router 的"分发"含义降级到标语里承载）。标语「整个文件夹一键拆完：
> 按部门、区域、工号等字段自动拆分，打包分发」；特性行「保留原格式 · 跨文件自动合并 · 单个
> 文件也能拆」；界面术语「拆分字段」「🔄 扫描字段」（不再用「按这列拆 / 识别列」）。
>
> **品牌文案二次修订（2026-07-03）**：中文副标再改为 **「Excel 智能拆分工具」**——作者要求
> 突出智能拆分能力（自动识别表头、自动枚举取值、跨文件智能合并等），「批量」只体现数量维度，
> 不能传达工具的智能识别能力。统一文案：窗口标题 `ExcelRouter · Excel 智能拆分工具`；
> README H1、`version.txt` FileDescription 同步。标语/特性行/界面术语不变。
> **此为最新定稿**，「Excel 批量拆分工具」已废弃，勿再混用。
>
> **新增 PDF 加密分发模式（2026-07-12，v2.5.1 后）**：界面顶部加了
> `CTkSegmentedButton` 模式切换「Excel 拆分 / PDF 加密分发」。新模式把同一批 PDF
> 按「网格 → 密码」Excel 映射清单，为每个网格生成专属打开密码（AES-256）+ 专属水印
> （网格名+日期，斜向半透明平铺，泄露可溯源）的副本，并输出含明文密码的「分发清单.xlsx」。
> 不做自动发送（用户照清单手动微信/邮件群发）。核心在 `core/pdf_dist.py`，
> 新依赖 pypdf + fpdf2 + cryptography（本项目**首个 C 扩展依赖**是 cryptography，
> 发版后要盯 VirusTotal 误报面变化）。
>
> **默认「不跨文件合并」+ 新增「只拆这些取值」（2026-07-17）**：`merge_across_files` 默认值
> 从 `true` 改为 **`false`**——默认行为改为「按原表输出，每个源文件各自拆分，只在 ZIP 里
> 打包在一起」，用户需要跨文件合并汇总表时自己去「高级设置」勾选「跨文件合并汇总」再开。
> 起因：宽表（几百列）+ 大批量文件时，跨文件合并把所有输出工作簿整批留在内存到最后统一保存，
> 叠加保留格式模式下的逐格式复制，会把内存/CPU 都推得很高，表现为界面乃至整机卡顿；改成
> 默认不合并后，风险场景变成用户主动选择，且默认路径更符合「批量给每个人分发各自原表」的
> 常见诉求。**旧版保存的 `user_config.json` 若已写了 `merge_across_files: true`，会按 FALLBACK
> 的补键规则保留用户的选择，不会被静默改回 false**——只有全新安装/从未保存过配置时才吃新默认值。
> 同时给 `selected_values`（只拆哪些主取值）补上了 GUI 入口：高级设置里新增「只拆这些取值
> （逗号分隔，留空=全部）」文本框（`gui/app.py` 的 `_selvals_var`），此前这个 config 字段
> 虽然 `core/splitter.py` 早就支持，但界面没有对应控件，只能手改 `user_config.json`。填好后
> 随「跑完自动存配置」的机制一起持久化，下次打开程序自动带出，不用每次重填。
>
> **新增「保留公式」选项（2026-07-21）**：源起用户诉求——结算表拆分给接收人时，希望对方能看到
> 计算公式（如「金额=单价×数量」）以自行核对，减少结算争议。新增 `keep_formulas` 勾选项，
> **默认关闭**、主要面向单文件场景（批量也可用但会多占一份源文件内存，日志会提示）。
> 采用**智能保留**而非无脑保留：拆分会压缩行号（源表第 100 行可能落到输出第 5 行），
> 只有**同行公式**（如 `=D100*E100`，只引用自己这一行的其它列）能安全平移行号后保留为活公式；
> **跨行/汇总/跨表公式**（`SUM`、`VLOOKUP`、跨 sheet 引用）无法平移，一律落成当前缓存数值
> （即现有行为）——绝不把一个会显示错误结果的公式发给接收人，这比看得到公式但算错更麻烦。
> 核心是 `core/splitter.py` 的 `_is_row_safe_ref` + `_safe_translate`（用 openpyxl 自带的
> `Tokenizer`/`Translator`，**无新依赖**，纯 Python，不影响 VirusTotal 误报面）。

---

## 运行环境

- **目标平台：Windows 10/11 64 位 + 国产化 Linux**（银河麒麟 V10 / 统信 UOS，
  x86_64 与 aarch64 双架构）。**单一代码库**，平台差异收敛在 `gui/platform_compat.py`
  和 `core/pdf_dist.py` 的字体探测里，不开分支。
- Python 3.9+（麒麟 V10 SP1 常自带 3.7，打包机需要先升级；**目标机不需要装 Python**）
- 依赖见 `requirements.txt`：customtkinter、openpyxl、pandas、xlrd、pypdf、fpdf2、cryptography
- Linux 额外的系统依赖：`python3-tk`（打包机）、`fonts-wqy-zenhei` 或 `fonts-noto-cjk`
  （否则 PDF 水印中文变 `?`）、`xdg-utils`（否则「打开输出文件夹」不可用）

---

## 目录结构与各文件职责

```
excel-router/
├── main.py                    # 入口，只有几行，调用 gui.app.run()
├── CLAUDE.md                  # 本文件
├── README.md                  # 中英双语说明（含实拍截图 docs/screenshot_*.jpg）
├── llms.txt                   # 给 AI 助手/搜索引擎的项目事实卡（能做/不做/关键词，见下「可发现性」）
├── LICENSE                    # MIT 许可证
├── requirements.txt           # 运行依赖
├── requirements-dev.txt       # 开发/测试依赖（pytest、pyinstaller、ruff）
├── conftest.py                # 让 pytest 能 import core 包
├── ruff.toml                  # Ruff 静态检查配置（全仓库 lint；core/tests 存量风格已按文件豁免）
├── pyrightconfig.json         # Pyright 类型检查配置（检查范围：gui/ + main.py，core 存量未纳入）
├── .gitignore                 # 忽略 dist/build/用户配置等
├── build.bat                  # Windows 一键打包：同时产出 onedir（推荐）+ onefile 两个产物
├── build_linux.sh             # 麒麟/UOS 打包：★ 必须在目标机器上跑，只产 onedir tar.gz
├── packaging/
│   ├── launcher.sh.in         # Linux 启动脚本模板（@GLIBC@ 由 build_linux.sh 替换）
│   └── README-Linux.txt       # 随 tar.gz 一起分发的「使用说明.txt」
├── .gitattributes             # ★ 钉死 *.sh 为 LF：本机 autocrlf=true，CRLF 的 sh 在麒麟上直接起不来
├── app.ico                    # 程序图标（Windows；占位，用户可替换）
├── app.png                    # 程序图标（Linux/X11 用 iconphoto，拷自 docs/logo.png）
├── version.txt                # ★ 仅 Windows 用（PyInstaller --version-file 是 PE 资源）。
│                              #   不要为 Linux 另造版本文件，build_linux.sh 从 APP_VERSION 解析
├── config/
│   └── default_config.json    # 默认配置（通用空配置，新字段见下「数据模型」）
├── user_config.json           # 用户保存的配置（运行时生成在程序目录，不入库）
├── core/
│   ├── __init__.py
│   ├── splitter.py            # 核心拆分逻辑（最重要的文件）
│   ├── pdf_dist.py            # PDF 按网格加密分发（密码/水印/分发清单）
│   └── utils.py               # 文本清理 + 取值归并 + 文件名净化
├── gui/
│   ├── __init__.py
│   ├── app.py                 # customtkinter 三步卡片式界面（Excel/PDF 双模式、打包路径适配、队列泵）
│   └── platform_compat.py     # 平台兼容层（打开文件管理器/拖放解析/UI字体/窗口图标/锁文件）
├── examples/
│   ├── make_sample.py         # 可复现样本生成器：5 个月份 × 55 名虚拟员工，3 行合并表头
│   └── {1-5}月A分公司明细.xlsx # 生成的演示样本（跨文件合并 + 到人演示用）
├── docs/
│   ├── FAQ.md                 # 常见问题（杀毒误报、公式空白、.xls限制、大文件进度等）
│   ├── RELEASING.md           # 维护者发版手册（CI 流程、误报处理、双产物说明）
│   └── screenshot_*.jpg       # README 用截图
├── .github/
│   ├── workflows/release.yml       # tag push v* 触发：测试→双 PyInstaller 构建→打包→发 Release
│   ├── workflows/release-linux.yml # ★ 只手动触发（按需发布，不跟 tag）：almalinux:8 × 双架构
│   └── ISSUE_TEMPLATE/        # Bug/Question 结构化表单，config.yml 禁用空白 issue
└── tests/
    ├── test_platform_compat.py # 平台兼容层单测（纯函数，无 Tk、可无头跑）
    ├── test_utils.py          # utils 单元测试
    ├── test_splitter.py       # splitter 集成测试（单文件/合并、到人双产出、格式保留）
    └── test_pdf_dist.py       # pdf_dist 集成测试（加密/水印/清单/容错/停止）
```

### 数据模型（config 关键字段）

| 字段 | 含义 | 默认 |
|---|---|---|
| `header_mode` | 表头识别：`auto` / `row`(指定行号) / `keyword`(旧关键词法) | `auto` |
| `header_row` | header_mode=row 时的 1 基行号 | `1` |
| `grid_keys` / `id_keys` | header_mode=keyword 时表头需同时含的两类关键词（任一空则不要求） | `[]` |
| `split_column` | 拆分字段（按字段名，GUI 里叫「拆分字段」） | `""` |
| `selected_values` | 只拆这些主取值；空 = 自动枚举该字段所有取值 | `[]` |
| `to_person` | 是否在汇总之外**附加产出到人**（仅文件夹批量有效） | `false` |
| `person_column` | 到人按哪个字段拆 | `""` |
| `person_file_filter` | 只对文件名命中这些关键词的表做到人；空 = 全部 | `[]` |
| `make_zip` | 批量时按主取值打包 ZIP | `true` |
| `value_alias_map` | 取值归并 `{规范值:[别名...]}`（旧 position_map 的通用化身） | `{}` |
| `skip_values` | 拆分列中要忽略的取值（合计/小计/空等） | `["合计","小计","总计","平均",""]` |
| `merge_across_files` | 汇总：同取值跨源文件是否合并到一个文件（到人始终按人合并） | `false`（2026-07-17 起，见下方说明） |
| `exact_match` `preserve_format` `auto_open_output` | 同义保留 | |
| `keep_formulas` | 保留公式（仅同行公式安全平移，需先开 `preserve_format`；见下方说明） | `false` |
| `ui_mode` | 界面模式：`excel`（拆分）/ `pdf`（加密分发） | `excel` |
| `pdf_input_paths` | 待分发的 PDF 文件列表（可多个） | `[]` |
| `pdf_mapping_path` | 「网格→密码」映射清单 xlsx 路径 | `""` |
| `pdf_grid_column` / `pdf_password_column` / `pdf_receiver_column` | 映射清单里的网格/密码/接收人（选填）列名 | `""` |
| `pdf_watermark` | 是否加网格专属水印 | `true` |
| `pdf_watermark_text` | 水印模板，支持 `{grid}`（网格名）`{date}`（日期） | `"{grid} {date}"` |
| `pdf_watermark_opacity` / `pdf_watermark_angle` | 水印透明度 / 旋转角（无 GUI 入口，改配置文件生效） | `0.15` / `45` |
| `window_geometry` | 上次关闭时的窗口大小位置；空 = 按屏幕高度自适应（v2.7） | `""` |
| `ui_scale` | 界面缩放百分比（大字号=115），重启生效（v2.7） | `100` |
| `appearance_mode` | 深浅色：system / light / dark，页脚循环按钮即时切换（v2.7.1） | `"system"` |

### core/splitter.py 的关键函数

- `detect_header_row_auto(rows, max_scan=15)` —— 通用启发式：非空多、文字多、下一行像数据 → 表头行
- `find_header_row(rows, grid_keys, id_keys)` —— 关键词法（专家兜底，接收**值矩阵**而非 ws）
- `resolve_header_row(rows, config)` —— 按 header_mode 分发 auto / row / keyword
- `list_columns(file_path, config)` / `list_values(file_path, config, column)` —— 供 GUI 下拉与多选
- `detect_uncalculated_formulas(work_path)` —— 检测「有公式但无缓存值」，采样前 200 行预警
- `_cached_copy(style_obj, cache)` —— 按源样式 id 缓存复制，保留格式提速的关键（**缓存按单个源文件作用域**）
- `_is_row_safe_ref(ref, origin_row)` / `_safe_translate(formula, origin, dest)` —— `keep_formulas`
  的安全校验+平移：只有「同表、单格、引用行 == origin_row」的公式才平移保留，否则返回 `None`
  （调用方落成缓存数值）。任何解析异常都按不安全处理
- `_OutputBook` —— 一个输出文件的内存工作簿；`get_or_create_sheet` + `_append_rows` 支持**跨源文件追加**
- `_summary_key_path` / `_person_key_path` —— 计算「汇总 / 到人」两套输出树的 key 与路径
- `_matches_filter(name, keywords)` —— 文件名是否命中到人范围关键词
- `normalize_to_xlsx(file_path, log_fn)` —— .xls 转临时 .xlsx
- `process_file(..., single_file, ...)` —— **始终产出汇总、可选附加到人**，追加进共享 outputs 注册表；
  `keep_formulas` 开启（且 `preserve_format` 有效）时额外打开一份 `data_only=False` 的源工作簿取
  公式文本，逐格式单元格调用 `_safe_translate`——**同行公式**平移后仍是活公式，其余（汇总/跨行/
  跨表公式）落成当前缓存数值，绝不把「会算错」的公式发给接收人
- `run_split(config, ...)` —— 主流程：**输入可为单文件或目录** → 累积 → 统一保存 → 按主取值打 ZIP；
  `process_file` 的 `stats` 可变 dict 参数跨文件累计「跳过 sheet/失败文件」，收尾打一行
  **`[SUMMARY] {json}`**（mode/groups/files/rows/zips/skipped_sheets/failed_files/failed_saves/
  stopped/output）——GUI 菜单据渲染完成摘要；追加式输出不改返回值，skill 只是多打一行

**架构要点：pandas 负责过滤数据（向量化 mask），openpyxl 负责复制格式。**

**一次运行两套产出（v2.1）：**
- **汇总**（始终）：按 `split_column` 拆。单文件 / `merge=True` → 扁平 `{主取值}.xlsx`（跨文件合并）；
  目录 + `merge=False` → `{主取值}/汇总/{主取值}_{原文件名}.xlsx`（原文件拆分，文件名带主取值前缀，
  脱离文件夹单独看也能分清是哪个取值的）。
- **到人**（可选）：`to_person` 开 + 有 `person_column` + 文件名命中 `person_file_filter` →
  `{主取值}/到人/{姓名}_{原文件名}.xlsx`（同一人跨文件按 sheet 名合并；文件名里的 `{原文件名}`
  取自第一个产生该人输出的源文件，后续源文件命中同一人只追加行、不改文件名，见下方
  `_person_key_path` 说明）。
- 批量时对每个生成了文件夹的主取值打 `{主取值}.zip`。

**跨文件合并（v2.0 核心，仍在）：** 输出按 key 在内存累积、统一保存；第一个产生某 key 的源文件
定义表头/格式，后续文件**按列位置追加**（到人则按 sheet 名分别累积）。从根本上修掉 v1.0「同名输出互相覆盖」bug。

**数据行两种模式，由 `config['preserve_format']` 控制（默认 True）：**
- **保留格式**：用 `wb_src` 逐行复制「值 + 完整格式」。按【原始行号】读取、加法式写入
  （`rows_df.index` 即原始 0 基行号，`excel_row = src_idx + 1`），**绝不删行**。
- **快速模式**：数据行只写值（来自 pandas），最快。

表头格式（含合并单元格、列宽）两种模式都保留。`wb_src` 在 `process_file` 末尾统一 `close()`。

**限制（已在 README/界面注明）：** ① `.xls` 转换后无法保留原格式；
② 数据区合并单元格暂不保留；③ 跨文件合并按**列位置**追加，最适合「同一套模板的多个表」。

### core/pdf_dist.py 的关键函数（PDF 加密分发）

- `list_mapping_columns(xlsx_path)` —— 读映射清单第 1 行表头，供 GUI 三个列下拉（清单是用户
  专门维护的小表，表头固定第 1 行，不做启发式识别）
- `read_mapping(xlsx_path, grid_col, password_col, receiver_col)` —— 返回 `[{"grid","password","receiver"}]`；
  **密码统一转字符串**（`_cell_str` 处理 openpyxl 把 001234 读成 1234/1234.0 的问题），
  空网格/空密码行跳过并警告，重复网格后者覆盖并警告
- `_find_cjk_font()` —— 探测中文字体。Windows：扫 `%WINDIR%\Fonts` 里的 simhei/Deng 等
  （行为与 v2.7.2 逐字一致）。Linux：`ER_CJK_FONT` 环境变量 → 扫 `/usr/share/fonts` 等目录
  按 `_LINUX_FONT_CANDIDATES` 匹配 → `fc-match`（timeout=3）。
  **注意：「fpdf2 不支持 .ttc」是过期说法** —— fpdf2 2.8.x 的 `add_font` 接受
  `.ttf/.otf/.ttc/.otc`，麒麟自带中文字体恰恰多是 .ttc/.otf。但**单体简中面
  （`*SC*.otf`）必须排在 .ttc 合集前面**：`NotoSansCJK-Regular.ttc` 的第 0 面通常是日文面，
  汉字能渲染但字形是日式变体。函数名/零参签名/`str|None` 返回值是对外契约，下游 skill
  靠猴补丁替换它实现 `--font`，别改签名
- `_font_usable(path)` —— 让 fpdf2 自己回答能不能加载，不靠扩展名猜。这是「优雅降级成 `???`」
  和「分发跑到一半崩掉」的分界线，**别删**
- `_make_watermark_pdf(text, w, h, ...)` —— fpdf2 内存生成平铺斜排水印页，按（尺寸+文字）缓存
- `_stamp_and_encrypt(reader, out_path, wm_text, password, algorithm)` —— 盖水印+设密码写盘。
  **writer 每个网格必须重建**：pypdf 的 merge_page 就地改页对象，复用会导致水印跨网格叠加；
  `writer.append(reader)` 是深拷贝，不会污染共享 reader
- `write_manifest(entries, out_path)` —— 写「分发清单.xlsx」，密码列强制文本格式（`number_format='@'`）
- `write_mapping_template(path)` / `fill_random_passwords(mapping, grid_col, password_col="", out_path=None)`
  / `gen_password(n=8)`（v2.7）——映射清单模板生成/随机密码补填写。模板把「填写说明」放
  **第二个 sheet**（避免说明文字混进数据被 read_mapping 误当网格）；随机密码用 `secrets`，
  字符集排除 0O1lI 易混字符，只填空白格不覆盖手填，**写入新文件**（`*_含密码.xlsx`）不动原清单
- `run_pdf_dist(config, log_fn, progress_fn, stop_flag)` —— **签名与 run_split 完全一致**，
  GUI 队列泵零改动接入；AES-256 加密，cryptography 缺失时降级 RC4-128 + 警告；
  单网格失败记日志继续；源 PDF 自带密码的跳过不中断；收尾同样打 `[SUMMARY] {json}`

输出结构：`输出根/{网格}/{原文件名}.pdf` + `分发清单.xlsx`（网格|文件名|密码|接收人|页数|状态）。
**清单含明文密码**，完成日志里有「勿随文件发出」提醒——不要删这句。

### gui/app.py（v2.4 后期：三步卡片式重构；队列泵机制不变）

**双模式（2026-07-12 起）**：`_build_header` 尾部的 `_mode_seg` 切换「Excel 拆分 / PDF 加密分发」，
`self._mode`（"excel"/"pdf"）随 `ui_mode` 持久化。`_body` 里 `_excel_frame` / `_pdf_frame` 两个
Frame 同格互斥显示（`_apply_mode` 用 grid/grid_remove 切换，同时改③卡片标题与主按钮文案）。
PDF 模式步骤卡：`_build_pdf_step_input`（多选 PDF）+ `_build_pdf_step_map`（选映射清单 →
子线程 `list_mapping_columns` → 泵消息 `("pdf_scan", ...)` → `_on_pdf_scan` 关键词预选三个列下拉；
水印子卡片）。`_start` 开头按 `self._mode` 分派到 `_start_pdf`（自带校验链），
两模式共享 `_enter_running`（运行态 UI 准备）、③操作区、进度条、日志、`_on_done`。

**布局**（self 的 grid 行）：row0 品牌区（品牌名+副标同行、标语、特性行，作者信息在页脚）→
row1 `CTkScrollableFrame` 步骤区（weight=1）：①②卡片 + 「▸ 高级设置」折叠 →
row2 **固定操作区**③卡片（输出目录 + 主按钮 + 进度 + 状态行，永远可达）→
row3 工具条（「▸ 处理详情」日志折叠钮 | 「💬 反馈建议」「保存配置」）→ row4 日志框（默认收起，
**失败时自动展开**）→ row5 页脚（🔒 数据仅在本机处理 · 作者 · MIT）。
`_step_card()` 生成带编号圆徽章的卡片；`_ghost_button`/`_flat_button` 是次要/纯文字按钮工厂
（勿改回 `**dict` 解包样式——pyright 无法对异构 dict 解包做类型匹配）。

- **① 选表格**：「📄 选一个 Excel 文件 / 📁 选整个文件夹」两个直选按钮（替代旧 SegmentedButton+浏览）；
  输入类型由 `os.path.isdir` 自动推断，`_update_input_ui` 控制批量区 `_batch_frame` 显隐；
  路径框可手动粘贴（Return/FocusOut → `_on_path_edited`，`_scanned_path` 防重复扫描）。
- **选完输入自动三连**（`_after_pick`）：显隐批量区 + `_suggest_output` 自动推荐输出目录
  + `_scan_input` 后台扫描。**保存位置是可选项**：不填也能开始（`_start` 里会自动补默认值）。
- **输出默认值规则**（`_auto_output_for`）：文件→同目录`拆分结果`；文件夹→**文件夹里面的**
  `拆分结果`（结果永远出现在数据旁边；放输入里是安全的，run_split 会整体跳过 output_root
  子树）。`_out_auto` 记录推荐值，用户手改过就不再覆盖；启动时若上次保存的输出恰是自动值，
  也识别为自动值（换输入时跟着更新，不会钉在旧位置）。
- `_scan_input`（替代旧 `_detect_columns`）在**子线程**统计 Excel 数、找模板、`list_columns`，
  经 UI 泵消息 `("scan", (cols, n, tpl_name, is_dir, p))` 回主线程；`_on_scan` 先校验 p 未过期。
  **计数镜像 run_split 的跳过规则**（当前输出目录子树不算），结果放输入里时数字才不虚高。
  `_on_columns` 填充下拉，保存值失效时用 `_recommend_split`/`_recommend_person` 按关键词智能预选。
- **② 选字段**：拆分字段下拉 + 🔄 扫描字段；批量选项子卡片在②内（仅文件夹显示）：每组 ZIP、
  同时拆到人（未勾选时子控件置灰，`_update_person_state`）。
- **③ 开始拆分**：`⏹ 停止` 只在运行时出现、`📂 打开输出文件夹` 成功后出现（`_last_output`）；
  `_run_status` 状态行运行时**镜像最新一条日志**（截 70 字符），完成 ✅ 绿 / 停止 ⏹ 橙 / 失败 ❌ 红。
- `_start` 校验链含**三道拦截**：① 目录模式下「输出 == 输入」或「输出是输入的上层」直接阻止
  （`run_split` 会把 output_root 前缀整体跳过，这两种选法会导致“找不到任何文件”）；
  ② `_find_stale_results`：输入里有历史拆分结果目录（`拆分结果`/`*_拆分结果`/`\d{8}结果`）
  且不被本次输出覆盖时弹确认——否则旧结果会被当数据重复拆；③ 开跑前 `os.makedirs` 预检
  保存位置可写，只读/无权限当场友好报错而不是跑一半失败。
- `_on_done` 成功时**静默 `save_config(cfg)`**（失败运行不存，避免记坏参数）；启动时若上次输入
  路径仍存在则自动扫描，实现「打开即用」。
- `_collect_config()` 字段与 v2.1 完全一致（config schema 未变）；`load_config` 用 FALLBACK
  补齐缺键，旧版配置不会报错。
- **匿名反馈入口（v2.3 起）**：模块级常量 `APP_VERSION`、`FEEDBACK_URL`（当前指向 WPS 匿名问卷）。
  「💬 反馈建议」按钮 `_open_feedback` 打开问卷 + 把版本号复制进剪贴板；`_on_done` 仅在**成功完成**
  时在日志追加一行反馈引导（失败路径不加，避免像推卸责任）。**不做任何遥测/自动上报**——
  「数据不出本地」是产品卖点，反馈只能是用户主动点开的外部链接，不要改成程序内嵌表单或自动上传。

**v2.7 界面机制增量（勿回退）：** 拖拽由 `_enable_drop`/`_parse_drop_paths` 承载，
仅依赖 tkinterdnd2 存在性（`_DND_OK`），exe 缺二进制时静默无拖拽 不报错；取值预览三个状态
`_tpl_path`/`_values_all`/`_value_vars` 随扫描重置，`("values", (col, token, vals))` 泵消息
带序号防过期回填；`[SUMMARY]` 与 ⚠ 计数在 `_pump_ui` 消费日志时提取（**[SUMMARY] 行不进
日志框显示**，提取失败退回原有纯日志行为）；`_collect_config` 透传 `ui_scale`/`window_geometry`
（否则成功一次就把字号配置写丢了）；「🔄 扫描字段」已改名「🔄 重新识别」，README/FAQ 口径同步。

**拆分线程 → 主线程通信（v2.4，重要，勿回退）：** 子线程把日志/进度/完成信号 `put` 进
`self._ui_q`（`queue.Queue`），主线程 `_pump_ui` 每 100ms 用 `after(100, self._pump_ui)` 自我调度、
一次性 `get_nowait()` 排空队列后批量刷新 UI。**不要改回 `self.after(0, self._log, ...)` 这种子线程直接
排程主线程回调的写法**——大文件（几万行）时 `core/splitter.py` 会高频报告进度/心跳日志，逐条
`after(0, ...)` 会把 Tk 事件队列打满，表现为界面卡死/黑条纹无响应（v2.3 用户实测反馈的 bug）。
点击「开始拆分」的瞬间要给出即时反馈：进度条先切 `mode="indeterminate"` 播放滚动动画 + 立刻打一条
日志，第一条真实进度值到达后 `_pump_ui` 自动切回 `determinate`（见 `_stop_indeterminate`）。
`core/splitter.py` 侧配合：`process_file` 的 `tick_fn` 报告文件内部阶段进度（读取→读格式→拆分→保存），
`_append_rows` 每 200 行调一次 `heartbeat()`（内含 `time.sleep(0.001)` 让出 GIL + 达到 2000 行报一次
日志心跳）——**这两个回调不要删**，否则单文件几万行场景又会退回“进度条一开始就 100%、中间像假死”。

---

## ⚠️ 已知历史 Bug（重要，勿重蹈覆辙）

早期版本（V33）有过一个**行号错位 Bug**：在删除行之后，用删除后的行号去索引 pandas
镜像 DataFrame，导致公式列（如「到账金额」）取到错误的值。

**当前架构已规避**：本版本不做「先删行再回填」，而是 pandas 直接 `mask` 过滤出要保留的行，
表头格式单独从缓存写入。如果未来要改回「删行」逻辑，**务必先按原始行号缓存镜像数据，
删除后再按缓存顺序回填**，不要用删除后的行号索引。

### ⚠️ v1.0 的「多表覆盖」Bug（v2.0 已修，勿回退）

v1.0 逐文件独立处理、每个输出文件都新建 Workbook 后 `save` 覆盖，输出文件名又不含源表名——
同一个人/取值出现在多个源表时，后一个表直接覆盖前一个，导致数据丢失。
**v2.0 改为按 key 在内存累积、统一保存的 `_OutputBook` 机制**（见上）。
未来若改输出写法，**务必保持「同 key 跨文件追加」语义**，不要回到「逐文件新建并覆盖同名文件」。

## ⚠️ 公式列依赖缓存值（实测确认）

本工具用 pandas 读公式列的**缓存计算结果**（不是重新计算）。实测结论：

- 真实 Excel/WPS 另存的文件 → 带缓存值 → 公式列**正常**显示真实数字 ✅
- 程序(openpyxl)生成、或公式从未被 Excel 计算过的文件 → 无缓存值 → 公式列读成空白 ❌

`core/splitter.py` 里的 `detect_uncalculated_formulas()` 会提前检测这种情况并在日志里预警，
提示用户「用 Excel 打开另存后再处理」。**不要删掉这个检测**，它直接防止公式列静默丢数据。
.xls 路径不受影响（xlrd 读 .xls 时本来就只读值，转换后已是纯值）。

---

## 项目当前状态与发版流程

初始打包/部署阶段（三层界面时代的 v2.0）早已完成，GitHub 仓库、CI 自动发版、FAQ、Issue 模板、
匿名反馈入口均已上线。**日常开发不需要手动打包/手动建 Release**——正确流程是：

### 1. 本地验证（改动核心逻辑后必做）

```bash
pip install -r requirements.txt
pip install -r requirements-dev.txt
py main.py                 # 本机解释器是 py 启动器，见 [[python-launcher-gotcha]]
# 注意（2026-09 实测）：从其它 Agent 沙箱（如灵犀）的 shell 调本机 Python 时，
# 沙箱会注入 PYTHONHOME 指向它自带的环境，导致 py/真实解释器报
# pyexpat DLL load failed、pip 解析到错误 site-packages。解法：命令前加
# PYTHONHOME= PYTHONPATH= 前缀，或直接调用真实解释器绝对路径。
ruff check .               # 静态检查，必须全绿（配置见 ruff.toml，存量豁免勿扩大）
npx pyright --pythonpath "$(py -c 'import sys; print(sys.executable)')"
                           # 类型检查（范围 gui/+main.py；必须传 --pythonpath，
                           # 因为 PATH 里的 python 是商店 stub，pyright 自己找不到依赖）
py -m py_compile main.py core/*.py gui/*.py
pytest -q
```

GUI 需要图形环境；无显示器时只做静态检查 + `pytest -q`。

### 2. 发版（真正会打包发布时）

```bash
git add -A && git commit -m "feat: ..."
git push origin main
git tag vX.Y.Z && git push origin vX.Y.Z    # push tag 才会触发 CI
```

推送 `v*` 格式的 tag 会自动触发 `.github/workflows/release.yml`：跑测试 → 语法检查 →
分别用 `--onedir`（推荐，防杀毒误报）和 `--onefile` 两种模式跑 PyInstaller → 打包 →
`softprops/action-gh-release` 自动创建 GitHub Release 并上传双产物。**改版本号时三处要同步**：
`gui/app.py` 的 `APP_VERSION`、`version.txt` 里的 `filevers`/`prodvers`/`FileVersion`/`ProductVersion`
四个字段、本文件开头的「当前版本」那一行——否则 exe 属性/标题栏/文档三方对不上。
详细的发版规范、误报处理、CI 排障历史见
`docs/RELEASING.md`（例如 Windows runner 默认 `pwsh` 不展开 glob、需要给用到 `*.py` 的
step 加 `shell: bash`；`Rename-Item` 是原地重命名不是移动，后面别再接多余的 `Move-Item`）。

本地手动打包（不走 CI）用 `build.bat`，同样会产出 onedir + onefile 两个产物，两个
PyInstaller 致命坑仍然成立、缺一不可：

> 1. **`--collect-all customtkinter` 必须有**：customtkinter 依赖 `assets/themes/*.json`
>    和字体文件，缺了会导致 exe 一启动就崩溃（报错找不到 blue.json 之类）。
> 2. **`--add-data "config;config"` 必须有**：否则打包后找不到默认配置文件
>    （Windows 分号 `;` 分隔，Linux/Mac 是冒号 `:`）。
> 3. **`--collect-all tkinterdnd2` 必须有（v2.7 起）**：tkinterdnd2 自带 tkdnd 原生二进制，
>    缺了打包后拖拽静默失效（不崩，但功能没了）。build.bat 与 CI 三处已加，别删。
> 4. **`--add-data "app.png;."` 必须有（v2.8 起）**：X11 的 `iconbitmap()` 只认 XBM，
>    Linux 窗口图标走 `iconphoto(app.png)`。Windows 侧也加了，为的是让 build.bat 和
>    build_linux.sh 能逐行对照 diff。
>
> 代码已做路径适配：`gui/app.py` 用 `sys._MEIPASS` 读打包进去的默认配置，用户配置
> `user_config.json` 存到 exe 所在目录。**不要把配置读取改回纯相对路径**，打包后会失效。

麒麟 / UOS 打包用 `build_linux.sh`，对应有**四个 Linux 专属的坑**：

> 1. **★ 必须在目标机器上打包**。PyInstaller 产物的 glibc 下限 == 打包机的 glibc，
>    且**不向下兼容**。在 ubuntu-22.04（glibc 2.35）上打的包，拿到麒麟 V10（2.28~2.31）
>    上会直接报 `GLIBC_2.34 not found`。有多台不同 SP 时，在**最旧**的那台上打。
>    CI 里的等价做法是跑在 `almalinux:8` 容器（glibc 2.28，与麒麟 V10 同代）——
>    **不能用 manylinux 容器**，它的 CPython 没编 `_tkinter`，customtkinter 根本 import 不了；
>    也不能比 2.28 更低，`actions/checkout@v4`（Node 20）要求 glibc ≥ 2.28。
> 2. **绝不产 `--onefile`**。onefile 每次启动都解包到 `/tmp/_MEIxxxx`，很多信创镜像把
>    `/tmp` 挂成 `noexec` → 直接起不来。Linux 只发 onedir tar.gz。
> 3. **`--icon` / `--version-file` / `--noconsole` 三个都要去掉**：前两个是 Windows PE
>    专属（ELF 上无意义），第三个在 Linux 上是空操作，留着只会让人误以为它做了事。
>    `--add-data` 分隔符也从 `;` 换成 `:`。
> 4. **`.gitattributes` 钉死 `*.sh` 为 LF**。本机 git 是 `core.autocrlf=true`，没有这条
>    声明的话启动脚本检出成 CRLF，在麒麟上报 `/bin/sh^M: bad interpreter`，
>    表现为「双击没反应」——非常难查，别删那个文件。

### 3. 代码混淆

项目目标是开源，**不需要**混淆，除非用户明确要求。

---

## 下游消费者：excelrouter-skill（2026-08-05 新增）

本仓 `excelrouter-skill/` 目录就是那个 Claude Skill（命令行脚本 `er_inspect.py` /
`er_split.py` / `er_pdf_dist.py` / `er_list.py`，不依赖 GUI），把本仓 `core/` 蒸馏成可被 AI
助手调用的技能。**它已并入本仓，不再是独立仓库**（原独立仓库 MarsandSea/excelrouter-skill
已于 2026-09-08 归档）。

发布链路：`.github/workflows/publish-skill.yml` 在**发版打 `v*` tag 时**触发，把本仓 `core/`
同步进 `excelrouter-skill/skills/excelrouter/scripts/vendor/core/`、过滤掉 GUI 专用依赖、
同步 manifest/plugin 版本号、跑 skill 测试，然后发布到 SkillHub 与 ClawHub。
因为 core 与 skill 同仓，不再需要跨仓比对 tag —— 打 tag 那一刻复制过去的 core 就是该 tag 的内容。
这意味着：

- **`run_split()` / `run_pdf_dist()` / `list_columns()` / `list_values()` /
  `list_mapping_columns()` / `read_mapping()` 的函数签名和 config 字段已经是对外契约**——
  改函数签名、改字段名、改默认值语义，都要想一下下游 skill 会不会跟着崩（下游有自己的
  `pytest`，签名对不上会在下游 CI 上报错，但语义变了未必报错，人肉过一遍更保险）。
- **`core/` 不许引入对 `gui/` 的任何依赖**（哪怕只是延迟 import），否则下游 vendor 整个
  `core/` 目录时会直接带进 customtkinter 依赖，skill 装不起来。这条本来就是现有架构
  （GUI 对 core 单向依赖），只是现在多了一个不能违反的理由。
- 下游只跟**发布 tag** 走，不跟 `main` 分支的半成品——本仓改动只要没打 tag，就不会出现在
  skill 里。正常发版流程（commit → push → `git tag vX.Y.Z` → push tag）不用做任何额外动作，
  skill 会在下一次定时同步（或维护者手动点一下 Actions 页的 Run workflow）时自动追上。

---

## 可发现性约定（SEO / AI 召回，2026-08-05 新增）

项目零 Star、零下载的瓶颈之一是「搜不到 + AI 答不出」。为此在文档层做了固定结构，**改功能时
要顺手维护，不要让它们和实际能力脱节**（说错比不说更伤——用户下载后发现不合用就直接流失）：

- **README 首屏三段式**：`📌 一句话介绍`（是什么 + 平台 + 协议 + 数据不出本机）→ **「它不做这些」**
  （表头扁平化 / 通用多表合并 / 数据分析 / Mac / 在线版，四条否定）→ `🎯 谁会用到它`（真实场景 +
  用户会搜的原生问句）。**否定清单是刻意的**，它同时服务于「筛掉不合用的人」和「让 AI 回答边界
  问题时有据可依」，不要为了显得功能多而删掉。
- **`llms.txt`（根目录）**：给 AI 助手/搜索引擎的纯文本事实卡，也可直接粘去问答平台。**功能有
  增删时，README 首屏、`llms.txt` 的「能做/不做」两节要一起改**，三处口径必须一致。
- **`docs/FAQ.md`**：标题写成**用户真实的提问句**（AI 检索按语义召回，问句命中率远高于名词短语），
  顶部有分组目录。README 与 `docs/使用指引.md` 直接链到 FAQ 的锚点，**改 FAQ 标题必须同步改所有
  引用锚点**（现有引用：`#杀毒软件误报`、`#处理大文件…`）。注意 GitHub 锚点会吃掉标点：
  标题里的 ` / ` 会变成两个连字符，写目录链接时优先用「、」避免踩坑。
- **仓库元数据**（GitHub topics / 简介、Gitee 简介 + 标签）由**维护者手动在网页端维护**，
  文案见 `docs/RELEASING.md` 的「仓库元数据文案」一节，改品牌文案时同步更新。
- **不做关键词堆砌**：所有关键词都必须出现在通顺的句子或真实场景里。堆砌既伤人类阅读，
  现在的搜索与 AI 排序也不吃这套。

---

## 代码风格约定

- 中文注释，函数加简短 docstring
- 所有源码文件保留版权头，**不要删**
- 保持模块化：核心逻辑在 `core/`，界面在 `gui/`，不要把业务逻辑写进 `main.py`
- 错误处理：拆分过程中单个文件/网格出错不要中断整体，记录到日志继续跑

---

## 给 Claude Code 的协作提示

- 改动前先 `git status` / `git log` 看清当前状态
- 大改动前先 commit，方便回滚（`git checkout` / `git revert`）
- 修改核心逻辑后，至少做 `python -m py_compile` 语法检查
- 不确定的破坏性操作（删文件、改架构），先问用户
