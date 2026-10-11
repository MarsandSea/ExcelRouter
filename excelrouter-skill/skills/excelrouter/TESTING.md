# 测试与质量保障

本 skill 随包自带测试，上架后也可自行验证：改过脚本、换过 Python 环境、
或想确认安装是否完整时，跑一遍就知道。

## 跑测试

```bash
pip install -r requirements.txt -r requirements-pdf.txt pytest
pytest -q tests/
```

`tests/` 下是**端到端冒烟测试**：以子进程方式真实调用三个 `er_*.py` 脚本，
解析它们 stdout 的最后一行 JSON——和 AI 实际调用你的方式完全一致，
所以测的不是内部函数，而是「用户/AI 真会走的那条路径」。

样本是程序现场构造的 Excel，**表头故意放在第 2 行**（第 1 行是大标题），
确保自动表头识别不是碰巧读对了第一行；并且**故意混入「合计」行**，
验证汇总行不会被当成一个拆分取值。

## 覆盖到什么

| 脚本 | 验证点 |
|---|---|
| `er_inspect.py` | 自动识别非首行表头；列出字段取值并过滤合计行；目录输入的文件计数；文件不存在时返回人话报错 |
| `er_split.py` | 缺 `--by` 时友好报错；`--dry-run` 不写盘；按字段产出预期目录；`--merge` 跨文件合并成单文件；`--values` 只拆指定取值 |
| `er_pdf_dist.py` | 读取映射清单列名；按网格加密并生成清单（断言产物 PDF 确实 `is_encrypted`）；缺列名时友好报错；xlsx 加打开密码后能用对的密码解开；只加水印不要密码列；老格式 / Office 加水印 / 两项全关在开跑前报错（Office 两条要 vendor 同步到 v2.10 才跑，之前自动跳过） |
| `er_merge.py` | 文件夹合成一张（按列名对齐调过顺序的列、合计行不合、副本只合一份、名单表不当报错）；一个文件的多个 sheet 合成一张（隐藏 sheet 不合）；表多时分几次续做、最后不留半成品；中途表格有改动从头合；混着旧拆分结果时拦下；合回拆分结果时跳过重复的「到人」；不加载 pandas；`--checked` 只合交了的表、「来源」写交表人 |
| `er_check.py` | 通报里有应交 / 未交 / 漏填到行号 / 截止时间、逐人催办；输入单个文件没给 `--by` 时报错；交了 .et / 网页格式 xls / 加密码 xlsx 的人不写成「未交」（要 vendor 同步到带 `unreadable()` 的内核才跑） |
| 不用 pandas | 屏蔽 pandas / numpy 后 `er_inspect.py`（列名、分组数、全量取值）和 `er_split.py` 照常跑完（v3.0 起） |

另有针对拆分算法与 PDF 分发内部逻辑的单元测试，见上游仓库
<https://github.com/MarsandSea/ExcelRouter> 的 `tests/`（桌面版与本 skill 共用同一套 `core/`）。

## 依赖说明

依赖分成两份，按需安装，不必一次装全：

- `requirements.txt` —— **Excel 拆分、汇总、查收只需要这 2 个**：`openpyxl`、`xlrd`
  （v3.0 起不再需要 `pandas`）
- `requirements-pdf.txt` —— 做文件加密分发时另外装：`pypdf`、`fpdf2`、`cryptography`、
  `msoffcrypto-tool`（最后这个只有给 Word / Excel / PPT 加密才用到）

脚本内置依赖自检（`check_deps`），缺什么会直接报人话错误并提示补哪个包，
不需要自己猜。
