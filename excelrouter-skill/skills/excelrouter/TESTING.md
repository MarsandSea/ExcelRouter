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
| `er_pdf_dist.py` | 读取映射清单列名；按网格加密并生成清单（断言产物 PDF 确实 `is_encrypted`）；缺列名时友好报错 |
| `_common.py` | Windows 上在导入 numpy **之前**自动把 OpenBLAS 线程数限制为 1（绕开内存分配崩溃）；用户自己设过该变量时不覆盖 |

另有针对拆分算法与 PDF 分发内部逻辑的单元测试，见上游仓库
<https://github.com/MarsandSea/excel-router> 的 `tests/`（桌面版与本 skill 共用同一套 `core/`）。

## 依赖说明

依赖分成两份，按需安装，不必一次装全：

- `requirements.txt` —— **只做 Excel 拆分时只需要这 3 个**：`openpyxl`、`pandas`、`xlrd`
- `requirements-pdf.txt` —— 做 PDF 加密分发时另外装：`pypdf`、`fpdf2`、`cryptography`

脚本内置依赖自检（`check_deps`），缺什么会直接报人话错误并提示补哪个包，
不需要自己猜。
