# 发版手册（维护者用）

给项目维护者（当前是 AbeLin）看的操作手册：怎么发新版本、怎么防杀毒软件误报、
误报真的发生了该怎么处理。目标是**发版全程自动化，人只需要打一个 tag**。

---

## 1. 日常发版流程

1. 改动代码，本地跑一遍验证：
   ```bash
   pytest -q
   python -m py_compile main.py core/*.py gui/*.py
   ```
2. 更新版本号（三处要一致，缺一会导致 exe 属性 / 标题栏 / 文档互相对不上）：
   - `gui/app.py` 的 `APP_VERSION` 常量（窗口标题、反馈按钮里都引用它）
   - `version.txt` 的 `filevers` / `prodvers` / `FileVersion` / `ProductVersion` 四个字段
   - `CLAUDE.md` 开头「当前版本」那一行
3. Commit、push 到 `main`。
4. 打 tag 并推送，**这一步会自动触发 CI 发版**（`vX.X.X` 换成实际版本号，
   比如 `v2.5.0`；千万别抄本手册里的旧示例，早已被占用会直接 push 失败）：
   ```bash
   git tag vX.X.X
   git push origin vX.X.X
   ```
5. 去仓库 **Actions** 页看 `Release` 工作流跑完（约 3-5 分钟），跑绿后
   **Releases** 页会自动出现两个产物：
   - `ExcelRouter-vX.X.X-win64.zip`（onedir，推荐分发）
   - `ExcelRouter-vX.X.X.exe`（onefile，备选）
6. 下载其中一个到本机，脱离开发环境（换个目录）双击冒烟测试：能正常打开、
   识别列、跑通一次拆分。
7. **Release 说明里带上 AI 助手引导**（固定一句，别漏）：「在用 AI 助手
   （WorkBuddy / Claude / WPS 灵犀）？同内核 Skill 版见
   [excelrouter-skill](https://github.com/MarsandSea/excel-router/tree/main/excelrouter-skill)，
   WorkBuddy 可从 [SkillHub](https://skillhub.cn/skills/excelrouter) 一键安装。」
   —— 从桌面版用户里筛 AI 助手用户，属合规引流方向（见营销计划「客户端优先」铁律）。

CI 配置见 `.github/workflows/release.yml`。它会先跑 `pytest -q` 拦住测试不过的版本，
测试失败则整个发版流程停止，不会出坏的 Release。同一个流程还会**尽力而为**同步一份到
Gitee（见下一节），Gitee 同步失败不影响本次 GitHub 发版。

---

## 1.5 Gitee 同步（国内下载主路径）

**背景**：GitHub 在国内访问慢、目标用户（普通办公人员）多数没有账号，实测下载转化率
很低。所以 GitHub 只作为「开源可查」的权威背书，**国内实际下载入口改为 Gitee**
（`https://gitee.com/Marsandsea/Excelrouter`，国内直连快、免登录直下）。CI 每次发版会
自动把代码 + tag 镜像过去，并在 Gitee 一并创建同款 Release（同样的 zip + exe 两个产物）。

**首次启用需要在 GitHub 仓库加一个 secret**（`Settings → Secrets and variables →
Actions → New repository secret`），名字必须是 **`GITEE`**（大小写敏感，workflow 里
直接引用 `secrets.GITEE`）：

| Secret 名 | 值 |
|---|---|
| `GITEE` | Gitee 私人令牌（Gitee 个人设置 → 安全设置 → 私人令牌，新建时勾选仓库/projects 读写权限） |

推送代码走 `https://oauth2:<token>@gitee.com/...`（Gitee HTTPS 认证的固定写法，用户名
写死 `oauth2`、密码位置放令牌，不需要单独的用户名 secret）；调用 Gitee API v5 创建
Release/上传附件时，同一个令牌作为 `access_token` 参数传入。

**这一步是尽力而为（`continue-on-error: true`）**：Gitee 挂了、令牌过期、限流等任何
失败都只会在 Actions 日志里留一条黄色警告，**不会导致本次 GitHub Release 失败**——
GitHub 始终是权威源，Gitee 只是下游镜像。如果发现 Gitee 侧长期没同步上，去 Actions
日志里看 `Mirror code + tag to Gitee` / `Create Gitee Release` / `Upload assets to
Gitee Release` 这三步的报错信息，通常是 token 过期或权限不够。

代码镜像用的是**强制推送**（`git push gitee HEAD:main --force`），因为 Gitee 那边定位
是纯镜像、不接受人工直接改动；如果有人手动在 Gitee 上提交了内容，下次发版会被覆盖。

---

## 1.6 仓库元数据文案（GitHub / Gitee，手动维护）

仓库的**简介**和**标签/topics** 是站内搜索、外部搜索引擎和 AI 召回都会读的字段，
CI 不会自动维护，**改品牌文案或增删大功能时手动同步一次**。以下为定稿文案，直接复制粘贴。

### GitHub（`Settings` 页顶部 About 的齿轮按钮）

**Description**（350 字符内，中英各一句，兼顾国内外检索）：

```
Windows 桌面工具：按任意字段（部门/区域/工号）批量拆分 Excel，保留复杂表头格式，可跨文件合并、二级拆到人、自动打包 ZIP 分发；另附 PDF 加密水印分发。免安装、MIT 开源、数据不出本机 · Split a whole folder of Excel files by any column on Windows — original formatting preserved, cross-file merge, per-person output, ZIP packaging, plus PDF password + watermark distribution. Free, MIT, 100% local.
```

**Topics**（只允许小写字母/数字/连字符，最多 20 个）：

```
excel excel-splitter split-excel xlsx spreadsheet openpyxl pandas python windows desktop-app gui customtkinter office-automation batch-processing pdf-encryption watermark no-code chinese
```

也可以用 gh CLI 一次性设置：

```bash
gh repo edit MarsandSea/excel-router --add-topic excel,excel-splitter,split-excel,xlsx,spreadsheet,openpyxl,pandas,python,windows,desktop-app,gui,customtkinter,office-automation,batch-processing,pdf-encryption,watermark,no-code,chinese
```

### Gitee（仓库 → 管理 → 基本信息）

**仓库简介**：

```
ExcelRouter · Excel 智能拆分工具：把一批 Excel 按部门/区域/工号等任意字段的取值批量拆成多个文件，保留原复杂表头格式，可跨文件合并、二级拆到人、自动打包 ZIP 分发；另支持 PDF 按网格加密 + 水印分发。Windows 免安装，MIT 开源免费，数据全程本机处理不上传。
```

**标签**（Gitee 支持中文标签，国内搜索主要吃这些词）：

```
Excel  Excel拆分  批量拆分  表格拆分  复杂表头  办公自动化  报表分发  开源免费  本地处理  Windows  Python  PDF加密
```

顺手把 Gitee 的**开源许可证**选成 MIT、**编程语言**选 Python、**项目分类**选「办公软件 /
开发工具」类目——这些字段都参与 Gitee 站内搜索排序，留空等于放弃排名。

> Gitee 仓库是 CI 强制推送的纯镜像，但**简介、标签、分类这些元数据不在 git 里，
> 不会被镜像覆盖**，放心在网页端改。

---

## 2. 为什么要两种打包形态

`--onefile` 单文件版每次启动都会把 Python 解释器和依赖**解压到系统临时目录再执行**，
这个「自解压后执行」的行为模式与蠕虫/木马的启动方式高度相似，是杀毒软件（尤其
Windows Defender）**启发式误报**的主要触发点之一；另一个触发点是 PyInstaller
官方预编译 Bootloader 的哈希值被广泛用于打包各类程序（含恶意程序），本身已经
被部分厂商标记为可疑。

`--onedir` 版把依赖以普通文件夹形式摊开，运行时不需要自解压，能大幅降低这类
启发式误报概率，因此**默认推荐用户下载 onedir 版**（见 README 的下载区）。
onefile 版仍然提供，图个方便，但风险自担并在文档里提前说明。

---

## 3. 每次发版后：VirusTotal 例行检查

发布新版本后，顺手查一次两个产物在 VirusTotal 上的检出情况：

1. 打开 https://www.virustotal.com/
2. 分别上传 `ExcelRouter-vX.X.X-win64.zip`（或解压后的 exe）和
   `ExcelRouter-vX.X.X.exe`
3. 记录检出引擎数量（如 `2/70`），对比两种形态的差异
4. 如果某个杀毒引擎误报，记下厂商名字——下一步要用

这一步不阻塞发版，是**观测**，用于判断要不要走第 4 步的误报申诉。

**验收标准：国内主流杀毒软件（Windows Defender、360、QQ电脑管家、火绒等）干净即可**，
目标用户是国内普通办公人员，这几家才是真实影响他们下载/使用信任的关键。冷门/海外引擎
（如 Bkav、Gridinsoft、Yandex 等）零星误报**不需要逐个申诉**——v2.5.0 实测 3/64、
全部是这类小众引擎，国内主流全部干净，判定为可接受，未做任何申诉动作。

> **PDF 加密分发功能引入 cryptography 依赖后（2026-07）**：cryptography 是本项目
> 打包产物里**第一个 C 扩展 + 加密类库**（此前全是纯 Python），加密类 .pyd/DLL 是
> 杀毒启发式的敏感面。它是 Python 生态最主流的加密库（官方签名 wheel、PyInstaller
> 内置 hook），预期风险可控，但**引入后的第一个版本要重点做本节的 VirusTotal 检查**，
> 对比上一版检出数是否明显上升；若国内主流引擎出现新误报，按第 4 步申诉处理。

> **拖拽功能引入 tkinterdnd2 后（2026-09，v2.7）**：tkinterdnd2 自带 tkdnd 原生 dll
> （PyInstaller 由 `--collect-all tkinterdnd2` 收进产物），又增加一块原生二进制面。
> v2.7.0 发版时同样按本节做 VirusTotal 对比检查；另外发版后要实测一次**打包版的拖拽**
> （拖拖拽文件夹进窗口应该触发自动扫描），验证 collect-all 生效了。

---

## 4. 真的遇到误报：处理流程

按性价比从高到低排序，**不要一上来就上重武器**：

### 4.1 向对应厂商提交误报申诉（首选，免费）

- **Windows Defender / Microsoft**：提交到
  [Microsoft Security Intelligence 误报反馈门户](https://www.microsoft.com/en-us/wdsi/filesubmission)，
  选择 "Software developer" → 上传文件 → 说明是开源项目 + 附 GitHub 链接。
  通常数天内会更新云端库，之后同一份文件哈希不再被拦截。
- **其他厂商**（火绒、360、QQ管家等）：各家一般都有类似的「文件误报申诉」入口，
  搜「厂商名 + 误报申诉」即可找到，操作方式类似。

### 4.2 如果同一版本反复被新误报（触发式，非默认执行）

说明公共 Bootloader 的哈希污染已经比较严重，可以考虑**自编译 PyInstaller
Bootloader**：从源码克隆 PyInstaller，用本地 C 编译器（GCC）重新编译
Bootloader，产出的二进制哈希是唯一的，不在任何公共特征库里。参考
[PyInstaller 官方文档 - Building the Bootloader](https://pyinstaller.org/en/stable/bootloader-building.html)。

这一步有实际的时间成本（需要装 C 编译环境、重新验证打包流程），**只有在
第 4.1 步不够用、且误报持续影响用户信任时才做**，不是每次发版的标配。

### 4.3 不做的事（成本不划算，明确记录以免以后纠结）

- **不用 Nuitka 重写打包链路**：真编译能进一步降低误报，但迁移成本高，
  与「零维护」的项目定位冲突。
- **不买 EV 代码签名证书**：数千元/年，本项目是免费开源工具，不构成收益模型，
  不值得为此掏钱。

---

## 5. 内部版本（如果需要，不走公共 Release）

如果需要一份带公司内部署名的定制版本（例如 `CompanyName` 改成内部邮箱），
**不要**把这类版本上传到公共 GitHub Release ——会有隐私/合规风险，也会让
公共受众困惑「这是不是带后门的企业定制版」。

正确做法：本地单独跑一次 `pyinstaller`，把 `--version-file` 换成一份内部专用
的版本信息文件，产物只发到公司内部群/网盘，不进 CI、不进公共 Release。
