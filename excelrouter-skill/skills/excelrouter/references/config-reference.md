# config 字段完整参考

`er_split.py` 的每个 CLI flag 最终都会落进传给 `core.splitter.run_split()` 的一个 config
dict。大多数场景用 CLI flag 就够了；需要精细控制、或者想直接复用桌面版导出的
`user_config.json` 时，用 `--config full.json` 把整份配置文件喂进去（`--input`/`--output`
仍然从命令行参数取，会覆盖文件里的同名字段）。

这份表和 excel-router 桌面版的 `config/default_config.json` 是同一套字段——凡是桌面版
支持的，命令行都支持，只是命令行默认值经过挑选（比如 `merge_across_files` 默认关闭）。

## Excel 拆分（run_split）

| 字段 | 对应 CLI flag | 默认 | 说明 |
|---|---|---|---|
| `input_path` | `--input` | 必填 | 单个 Excel 文件或一个目录（目录会递归找 `.xlsx`/`.xls`） |
| `output_path` | `--output` | 必填 | 输出根目录；真实产出在其下 `{MMDDHHMM}结果/` 子目录，务必读返回 JSON 的 `output_path` |
| `split_column` | `--by` | 必填（除非用 `--config`） | 主拆分字段，按列名 |
| `selected_values` | `--values a,b` | 空=全部 | 只拆这些取值；留空自动枚举该列全部取值 |
| `person_column` | `--to-person 字段` | 空 | 到人二级拆分字段；填了就自动打开 `to_person`（仅目录输入有效，单文件会被忽略） |
| `person_file_filter` | `--person-filter kw1,kw2` | 空=全部 | 到人时只处理**文件名**命中这些关键词的源表 |
| `header_mode` | `--header-mode` | `auto` | `auto`（启发式自动识别）/ `row`（指定行号）/ `keyword`（关键词法，专家兜底） |
| `header_row` | `--header-row` | `1` | `header_mode=row` 时的 1 基行号 |
| `grid_keys` / `id_keys` | `--grid-keys` / `--id-keys`（逗号分隔） | 空 | `header_mode=keyword` 时表头需同时含的两类关键词 |
| `value_alias_map` | `--alias-json '{"东区":["东片区"]}'` | 空 | 取值归并：把同一个业务概念的不同写法算作一个取值 |
| `skip_values` | `--skip-values a,b` | `合计,小计,总计,平均,(空)` | 拆分列里要忽略的取值（一般是汇总行） |
| `exact_match` | `--fuzzy-match`（取反） | `True` | `--values` 是否精确匹配；`--fuzzy-match` 改成包含匹配 |
| `merge_across_files` | `--merge` | `False` | 同一取值跨源文件是否合并到一个输出文件（到人始终按人合并，不受此项影响） |
| `make_zip` | `--no-zip`（取反） | `True` | 批量拆分后，每个产出了文件夹的主取值是否打 ZIP |
| `preserve_format` | `--fast`（取反） | `True` | 关闭后数据行只写值不保留格式，明显更快 |
| `keep_formulas` | `--keep-formulas` | `False` | 尽量保留"同行公式"为活公式；跨行/汇总/跨表公式仍落成缓存数值。需要 `preserve_format` 同时开启才有意义 |

## PDF 加密分发（run_pdf_dist）

| 字段 | 对应 CLI flag | 默认 | 说明 |
|---|---|---|---|
| `pdf_input_paths` | `--pdf a.pdf b.pdf` | 必填 | 待分发的 PDF，可多个 |
| `pdf_mapping_path` | `--mapping` | 必填 | "网格→密码"映射清单 xlsx，表头必须在第 1 行 |
| `pdf_grid_column` / `pdf_password_column` / `pdf_receiver_column` | `--grid-col` / `--password-col` / `--receiver-col` | 空（前两个必填） | 映射清单里对应的列名；接收人列选填 |
| `output_path` | `--output` | 空 | 不填则用"第一个 PDF 所在目录/分发结果"，**不带时间戳**（和 Excel 拆分不同） |
| `pdf_watermark` | `--no-watermark`（取反） | `True` | 是否加网格专属水印 |
| `pdf_watermark_text` | `--watermark-text` | `"{grid} {date}"` | 水印模板，`{grid}` 替换成网格名，`{date}` 替换成当天日期 |
| `pdf_watermark_opacity` | `--opacity` | `0.15` | 水印透明度 0~1 |
| `pdf_watermark_angle` | `--angle` | `45` | 水印旋转角度 |

加密算法：装了 `cryptography` 用 AES-256，没装降级 RC4-128（stderr 会警告）。
中文水印需要系统里有中文字体（Windows 自动探测 `C:\Windows\Fonts`）；非 Windows 环境
或想指定字体，用 `--font 字体.ttf` 路径。

## 两套输出结构的区别

**Excel 拆分**（`run_split`）：
- `merge_across_files=False`（默认）：`{output_path}/{主取值}/汇总/{原文件名}.xlsx`
- `merge_across_files=True`：`{output_path}/{主取值}.xlsx`（跨文件合并成一个扁平文件）
- 到人（`to_person` 开）：额外产出 `{output_path}/{主取值}/到人/{姓名}.xlsx`（同一人跨文件按 sheet 合并）
- 批量时，每个产出了文件夹的主取值会打 `{主取值}.zip`

**PDF 分发**（`run_pdf_dist`）：
- `{output_path}/{网格}/{原PDF文件名}.pdf`（加密+水印后的副本）
- `{output_path}/分发清单.xlsx`（网格|文件名|密码|接收人|页数|状态，**含明文密码**）
