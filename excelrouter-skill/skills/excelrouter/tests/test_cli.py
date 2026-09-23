"""三个 er_*.py CLI 的冒烟测试：以子进程方式跑（就是 Claude 实际会怎么调），
解析 stdout 最后一行 JSON。样本构造沿用 excel-router 上游 tests/ 的写法
（表头故意不放第一行、带样式、含合计行），确保 CLI 包装层没有偷懒漏传参数。
"""
import json
import os
import subprocess
import sys

import openpyxl
import pytest
from openpyxl.styles import Font, PatternFill

# 本文件位于 <skill 目录>/tests/ 下：向上两级就是 skill 目录本身（scripts/ 与它同级）。
# 测试随 skill 一起发布上架，所以路径必须按「包内位置」推导，不能依赖仓库根目录。
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(SKILL_DIR, "scripts")

HEADERS = ["工号", "姓名", "部门", "城市", "金额"]


def _make_book(path, rows, title="月度报表"):
    """表头在第 2 行、第 1 行是大标题——专测自动表头识别不是碰巧读对了第一行。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = title
    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    for i, r in enumerate(rows, start=3):
        for c, v in enumerate(r, 1):
            ws.cell(row=i, column=c, value=v)
    wb.save(path)


def _run(script, *args):
    """跑一个 er_*.py，返回 (returncode, json_obj, stderr_text)。"""
    proc = subprocess.run(
        [sys.executable, os.path.join(SCRIPTS, script), *args],
        capture_output=True, text=True, encoding="utf-8",
    )
    stdout_lines = [line for line in proc.stdout.splitlines() if line.strip()]
    assert stdout_lines, f"stdout 应至少有一行 JSON，实际为空。stderr:\n{proc.stderr}"
    obj = json.loads(stdout_lines[-1])
    return proc.returncode, obj, proc.stderr


@pytest.fixture
def two_books(tmp_path):
    inp = tmp_path / "in"
    out = tmp_path / "out"
    inp.mkdir()
    out.mkdir()
    _make_book(inp / "A.xlsx", [
        ["001", "张三", "销售部", "北京", 100],
        ["002", "李四", "销售部", "北京", 200],
        ["003", "王五", "销售部", "上海", 150],
        ["004", "赵六", "技术部", "北京", 300],
        ["", "合计", "", "", 750],
    ])
    _make_book(inp / "B.xlsx", [
        ["005", "钱七", "销售部", "广州", 120],
        ["006", "孙八", "技术部", "深圳", 220],
    ])
    return str(inp), str(out)


# ---------- er_inspect.py ----------

def test_inspect_single_file_detects_header_and_columns(tmp_path):
    f = tmp_path / "a.xlsx"
    _make_book(f, [["001", "张三", "销售部", "北京", 100]])
    rc, obj, _ = _run("er_inspect.py", "--input", str(f))
    assert rc == 0
    assert obj["ok"] is True
    assert obj["is_dir"] is False
    assert obj["excel_count"] == 1
    assert obj["header_row"] == 2
    assert obj["columns"] == HEADERS


def test_inspect_lists_values_and_skips_total_row(tmp_path):
    f = tmp_path / "a.xlsx"
    _make_book(f, [
        ["001", "张三", "销售部", "北京", 100],
        ["002", "李四", "技术部", "北京", 200],
        ["", "合计", "", "", 300],
    ])
    rc, obj, _ = _run("er_inspect.py", "--input", str(f), "--column", "部门")
    assert rc == 0
    assert obj["values"] == ["技术部", "销售部"]


def test_inspect_directory_counts_files(two_books):
    inp, _out = two_books
    rc, obj, _ = _run("er_inspect.py", "--input", inp)
    assert rc == 0
    assert obj["is_dir"] is True
    assert obj["excel_count"] == 2


def test_inspect_directory_values_are_flagged_as_sample(two_books):
    """目录输入默认只读「样本表」，取值数字本身没错，错的是把抽样结论说成全量结论
    —— 上游 v2.9.0 在 GUI 侧修的就是这个。这里 A.xlsx 只有北京/上海，
    广州/深圳在 B.xlsx 里，样本口径必然看不全，所以必须自报 values_scope。"""
    inp, _out = two_books
    rc, obj, _ = _run("er_inspect.py", "--input", inp, "--column", "城市")
    assert rc == 0
    assert obj["values_scope"] == "sample"
    assert obj["values_sample_file"] == "A.xlsx"      # 排序后固定是它，换次运行不该变
    assert "抽样" in obj["values_warning"]
    assert "广州" not in obj["values"]                # 正是抽样会漏掉的那个


def test_inspect_all_files_unions_values_across_books(two_books):
    """--all-files 读遍所有表取并集，并把口径标成 all、不再带警告。"""
    inp, _out = two_books
    rc, obj, _ = _run("er_inspect.py", "--input", inp, "--column", "城市", "--all-files")
    assert rc == 0
    assert obj["values_scope"] == "all"
    assert obj["values_from_files"] == 2
    assert set(obj["values"]) == {"北京", "上海", "广州", "深圳"}
    assert "values_warning" not in obj


def test_inspect_single_file_values_scope_is_single(tmp_path):
    """单文件没有抽样问题，不该吓唬用户报 sample。"""
    book = tmp_path / "one.xlsx"
    _make_book(str(book), [["001", "张三", "销售部", "北京", 100]])
    rc, obj, _ = _run("er_inspect.py", "--input", str(book), "--column", "城市")
    assert rc == 0
    assert obj["values_scope"] == "single"
    assert "values_warning" not in obj


def test_inspect_missing_input_reports_friendly_error(tmp_path):
    rc, obj, _ = _run("er_inspect.py", "--input", str(tmp_path / "not-exist.xlsx"))
    assert rc == 1
    assert obj["ok"] is False
    assert "找不到" in obj["error"]


# ---------- er_split.py ----------

def test_split_missing_by_reports_friendly_error(two_books):
    inp, out = two_books
    rc, obj, _ = _run("er_split.py", "--input", inp, "--output", out)
    assert rc == 1
    assert obj["ok"] is False
    assert "--by" in obj["error"]


def test_split_dry_run_does_not_write_files(two_books):
    inp, out = two_books
    rc, obj, _ = _run("er_split.py", "--input", inp, "--output", out, "--by", "部门", "--dry-run")
    assert rc == 0
    assert obj["dry_run"] is True
    assert obj["config"]["split_column"] == "部门"
    assert os.listdir(out) == []


def test_split_produces_expected_output_files(two_books):
    inp, out = two_books
    rc, obj, stderr = _run("er_split.py", "--input", inp, "--output", out, "--by", "部门")
    assert rc == 0, stderr
    assert obj["ok"] is True
    output_path = obj["output_path"]
    assert os.path.isdir(output_path)
    # 默认 merge_across_files=False：按原表各自拆分，两个源文件各自产出「部门」子目录
    produced = set(os.listdir(output_path))
    assert "运行日志.txt" in produced
    # A.xlsx 和 B.xlsx 都含「销售部」，都产出该主取值目录
    assert "销售部" in produced or "销售部.zip" in produced


def test_split_merge_flag_merges_across_files(two_books):
    inp, out = two_books
    rc, obj, stderr = _run("er_split.py", "--input", inp, "--output", out, "--by", "部门", "--merge", "--no-zip")
    assert rc == 0, stderr
    output_path = obj["output_path"]
    # merge=True 时同取值跨文件合并为扁平文件：{部门}.xlsx 直接在输出根
    assert os.path.exists(os.path.join(output_path, "销售部.xlsx"))


def test_split_selected_values_filters(two_books):
    inp, out = two_books
    rc, obj, stderr = _run("er_split.py", "--input", inp, "--output", out, "--by", "部门",
                            "--values", "技术部", "--merge", "--no-zip")
    assert rc == 0, stderr
    output_path = obj["output_path"]
    assert os.path.exists(os.path.join(output_path, "技术部.xlsx"))
    assert not os.path.exists(os.path.join(output_path, "销售部.xlsx"))


# ---------- er_pdf_dist.py ----------

def _make_pdf(path, pages=1):
    from fpdf import FPDF
    pdf = FPDF()
    pdf.set_auto_page_break(False)
    for i in range(pages):
        pdf.add_page()
        pdf.set_font("Helvetica", size=14)
        pdf.text(30, 40, f"SOURCE page{i + 1}")
    pdf.output(str(path))


def _make_mapping(path, rows, headers=("GridName", "Pwd", "Receiver")):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(list(headers))
    for r in rows:
        ws.append(list(r))
    wb.save(path)


def test_pdf_dist_list_columns(tmp_path):
    mp = tmp_path / "map.xlsx"
    _make_mapping(mp, [("GridA", "001234", "Zhang")])
    rc, obj, _ = _run("er_pdf_dist.py", "--mapping", str(mp), "--list-columns")
    assert rc == 0
    assert obj["columns"] == ["GridName", "Pwd", "Receiver"]


def test_pdf_dist_encrypts_and_writes_manifest(tmp_path):
    pdf_path = tmp_path / "report.pdf"
    _make_pdf(pdf_path)
    mp = tmp_path / "map.xlsx"
    _make_mapping(mp, [("GridA", "001234", "Zhang"), ("GridB", 8888, "Li")])
    out = tmp_path / "dist"
    rc, obj, stderr = _run(
        "er_pdf_dist.py", "--pdf", str(pdf_path), "--mapping", str(mp),
        "--grid-col", "GridName", "--password-col", "Pwd", "--receiver-col", "Receiver",
        "--output", str(out), "--no-watermark",
    )
    assert rc == 0, stderr
    assert obj["ok"] is True
    assert os.path.exists(obj["manifest"])
    assert os.path.exists(os.path.join(obj["output_path"], "GridA", "report.pdf"))

    from pypdf import PdfReader
    reader = PdfReader(os.path.join(obj["output_path"], "GridA", "report.pdf"))
    assert reader.is_encrypted


def test_pdf_dist_random_password_reuses_existing_blank_column(tmp_path):
    """--random-password 不传 --password-col 时，清单里已有的空白「密码」列应被就地填满，
    而不是在表尾再追加一个同名列（追加会让 read_mapping 命中原来那个空列 → 密码全漏读）。

    上游 v2.7.2 修了这个（issue #1），包装层原本有一段规避、现已删除——
    这个用例就是那段规避的替代品：vendor 一旦退回旧版本，这里会立刻红。
    """
    mp = tmp_path / "grid.xlsx"
    _make_mapping(mp, [("GridA", None, "Zhang"), ("GridB", None, "Li")],
                  headers=("网格", "密码", "接收人"))
    rc, obj, stderr = _run("er_pdf_dist.py", "--mapping", str(mp),
                           "--grid-col", "网格", "--random-password")
    assert rc == 0, stderr
    assert obj["generated"] == 2
    out = obj["mapping_with_passwords"]
    wb = openpyxl.load_workbook(out)
    header = [c.value for c in wb.active[1]]
    assert header.count("密码") == 1, f"密码列被重复追加了：{header}"
    assert obj["password_col"] == "密码"
    # 真的填进去了，而且是填在原来那一列
    pwds = [r[header.index("密码")].value for r in wb.active.iter_rows(min_row=2)]
    assert all(p for p in pwds), f"密码没填进既有列：{pwds}"


def test_pdf_dist_missing_columns_reports_friendly_error(tmp_path):
    pdf_path = tmp_path / "report.pdf"
    _make_pdf(pdf_path)
    mp = tmp_path / "map.xlsx"
    _make_mapping(mp, [("GridA", "001234", "Zhang")])
    rc, obj, _ = _run("er_pdf_dist.py", "--pdf", str(pdf_path), "--mapping", str(mp))
    assert rc == 1
    assert obj["ok"] is False
    assert "--grid-col" in obj["error"]


@pytest.mark.skipif(os.name != "nt", reason="仅 Windows 有 OpenBLAS 多线程崩溃问题")
def test_common_pins_openblas_threads():
    """_common 必须在导入 numpy 之前把 OpenBLAS 线程数限制为 1。

    真实故障：Windows 部分环境 numpy 底层 OpenBLAS 多线程会因内存分配失败直接崩，
    报错里带 Memory allocation / Intel MKL，与用户数据无关。
    这个测试防止有人日后调整导入顺序、把设置挪到 numpy 之后（那样就失效了）。
    """
    code = (
        "import sys, os;"
        "sys.path.insert(0, %r);"
        "os.environ.pop('OPENBLAS_NUM_THREADS', None);"
        "import _common;"
        "assert 'numpy' not in sys.modules, 'numpy 必须在 _common 之后才被导入';"
        "print(os.environ.get('OPENBLAS_NUM_THREADS', ''))" % SCRIPTS
    )
    proc = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True)
    assert proc.returncode == 0, proc.stderr
    assert proc.stdout.strip() == "1"


@pytest.mark.skipif(os.name != "nt", reason="仅 Windows 有 OpenBLAS 多线程崩溃问题")
def test_common_does_not_override_user_setting():
    """用户自己设过 OPENBLAS_NUM_THREADS 时要尊重他的选择（setdefault 语义）。"""
    env = dict(os.environ, OPENBLAS_NUM_THREADS="8")
    code = (
        "import sys, os;"
        "sys.path.insert(0, %r);"
        "import _common;"
        "print(os.environ.get('OPENBLAS_NUM_THREADS', ''))" % SCRIPTS
    )
    proc = subprocess.run([sys.executable, "-c", code], capture_output=True,
                          text=True, env=env)
    assert proc.returncode == 0, proc.stderr
    assert proc.stdout.strip() == "8"
