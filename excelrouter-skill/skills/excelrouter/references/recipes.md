# 常见场景 → 命令配方

都假设当前目录在 `skills/excelrouter/` 下（脚本相对路径按此写；换成绝对路径同样能跑）。

## 1. "把这张表按部门拆开，每个部门一个文件"

```bash
python scripts/er_inspect.py --input 明细.xlsx
# 看到 columns 里有"部门"字段之后：
python scripts/er_inspect.py --input 明细.xlsx --column 部门
# 确认取值列表没有意外项之后：
python scripts/er_split.py --input 明细.xlsx --output 拆分结果 --by 部门
```

## 2. "这一整个文件夹的月度报表，每个人整理出一份自己的记录"

到人拆分只在**目录输入**下生效，且需要指定文件名过滤（避免把无关的表也拆进去）：

```bash
python scripts/er_inspect.py --input 月度报表/ --column 姓名
python scripts/er_split.py --input 月度报表/ --output 拆分结果 \
    --by 部门 --to-person 姓名 --person-filter 明细,报表
```
产出既有按部门的汇总，也有 `{部门}/到人/{姓名}.xlsx`（同一人跨文件按 sheet 合并）。

## 3. "只要东区和西区的数据，其它区域不用管"

```bash
python scripts/er_split.py --input 一批表/ --output 拆分结果 --by 区域 --values 东区,西区
```

## 4. "同一个人在好几个文件里都有数据，希望合成一张总表"

默认按原表分别拆分；要合并成一张扁平表用 `--merge`：

```bash
python scripts/er_split.py --input 一批表/ --output 拆分结果 --by 区域 --merge
```
产出变成 `拆分结果/{时间戳}结果/{区域}.xlsx`（不再是文件夹套文件夹）。

## 5. "部门这一列有的写'销售部'有的写'销售部门'，是同一个部门"

```bash
python scripts/er_split.py --input 一批表/ --output 拆分结果 --by 部门 \
    --alias-json '{"销售部":["销售部门","销售组"]}'
```

## 6. "表头不在第一行，自动识别失败了"

先看 `er_inspect.py` 报的 `header_row` 是不是 -1 或者明显不对，再手动指定：

```bash
python scripts/er_inspect.py --input 明细.xlsx --header-mode row --header-row 3
python scripts/er_split.py --input 明细.xlsx --output 拆分结果 --by 部门 --header-mode row --header-row 3
```

## 7. "文件很大/几百列，跑起来太慢"

放弃保留格式换速度：

```bash
python scripts/er_split.py --input 大表.xlsx --output 拆分结果 --by 部门 --fast
```

## 8. "接收人要能看到金额是怎么算出来的公式"

```bash
python scripts/er_split.py --input 结算表.xlsx --output 拆分结果 --by 部门 --keep-formulas
```
只有"同一行内"的公式（如 `=D2*E2`）能安全保留成活公式；跨行/汇总/跨表的公式会落成当前数值——
这是为了不把一个会显示错误结果的公式发给对方。

## 9. "先看看会拆出哪些文件，不要真的写盘"

```bash
python scripts/er_split.py --input 明细.xlsx --output 拆分结果 --by 部门 --dry-run
```

## 10. "按网格给一批 PDF 分别加密码"

```bash
python scripts/er_pdf_dist.py --mapping 网格密码表.xlsx --list-columns
python scripts/er_pdf_dist.py --pdf 报告.pdf --mapping 网格密码表.xlsx \
    --grid-col 网格 --password-col 密码 --receiver-col 接收人 --output 分发结果
```

## 11. "水印文字想自定义，不想要默认的'网格+日期'"

```bash
python scripts/er_pdf_dist.py --pdf 报告.pdf --mapping 网格密码表.xlsx \
    --grid-col 网格 --password-col 密码 --output 分发结果 \
    --watermark-text "内部资料·{grid}·{date}·请勿外传"
```

## 12. "一个文件里有两个表，要按网格拆、网格里再按各自的人拆（一步到位）"

多 sheet 文件（如「结对子维表」按 1看管人、「渠道维表」按 看管 渠道管理员），
两个表的人字段不同。用 `--to-person` 的 `sheet:列` 映射一次跑完，不用先拆网格再喂回去：

```bash
# 先看两个 sheet 各自的字段（v2.6.1 起 er_inspect 会列出每个 sheet）
python scripts/er_inspect.py --input 城区维表.xlsx --column 网格
python scripts/er_inspect.py --input 城区维表.xlsx --column "看管 渠道管理员"   # 跨 sheet 搜索，不再误报

# 一步：按网格拆 + 网格内按各自人字段拆到人
python scripts/er_split.py --input 城区维表.xlsx --output 拆分结果 \
    --by 网格 --to-person "结对子维表:1看管人,渠道维表:看管 渠道管理员" --no-zip
# 产出：拆分结果/{网格}.xlsx（网格汇总）+ 拆分结果/{网格}/到人/{姓名}_{网格}.xlsx
```

嫌写映射麻烦，可让程序按关键词自动挑每个 sheet 的人字段：

```bash
python scripts/er_split.py --input 城区维表.xlsx --output 拆分结果 --by 网格 --to-person auto
```

> 内部实现：vendor 先把文件按网格拆好（格式完整保留），scripts 包装层（`_person_split.py`）
> 再读这些网格文件、按每个 sheet 的人字段把行拆给对应的人。全程不修改 `vendor/core/`，
> 上游同步不会把这段补丁冲掉。
```

## 13. "要发 PDF 保密通报，但连网格密码表都还没有"

先用 `--make-template` 生成一份模板（第 1 个 sheet 是待填表格，第 2 个 sheet 是给用户看的
填写说明），让用户填好网格和接收人，再回来走配方 10/14：

```bash
python scripts/er_pdf_dist.py --make-template 网格密码表.xlsx
```

## 14. "密码不想自己想，让工具随机生成"

给映射清单里密码为空的网格自动生成 8 位随机密码（排除 0O1lI 等易混字符），另存一份
「原名_含密码.xlsx」，原清单不动；已手填的密码保留不覆盖。之后用新清单正式分发：

```bash
python scripts/er_pdf_dist.py --mapping 网格密码表.xlsx --grid-col 网格 --random-password
# 返回 JSON 的 mapping_with_passwords 就是新清单路径，password_col 告诉你密码列叫什么
```
