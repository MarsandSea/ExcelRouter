"""core/pdf_dist.py 的集成测试：加密、水印、分发清单、容错、停止。

水印文字用例统一用 ASCII 网格名——CI（GitHub Windows runner）字体环境
不可控，避免与中文字体探测耦合；中文水印在本机冒烟里验证。
"""
import openpyxl
import pytest
from fpdf import FPDF
from pypdf import PdfReader

from core.pdf_dist import (
    list_mapping_columns, read_mapping, run_pdf_dist, MANIFEST_NAME,
)


# ---------- 样本构造 ----------

def _make_pdf(path, pages=2, encrypt_pwd=None):
    pdf = FPDF()
    pdf.set_auto_page_break(False)
    for i in range(pages):
        pdf.add_page()
        pdf.set_font("Helvetica", size=14)
        pdf.text(30, 40, f"SOURCE-MARK page{i + 1}")
    if encrypt_pwd:
        from pypdf import PdfWriter
        import io
        reader = PdfReader(io.BytesIO(bytes(pdf.output())))
        writer = PdfWriter()
        writer.append(reader)
        writer.encrypt(user_password=encrypt_pwd, algorithm="AES-256")
        with open(path, "wb") as f:
            writer.write(f)
    else:
        pdf.output(str(path))


def _make_mapping(path, rows, headers=("GridName", "Pwd", "Receiver")):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(list(headers))
    for r in rows:
        ws.append(list(r))
    wb.save(path)


@pytest.fixture
def workspace(tmp_path):
    src = tmp_path / "report.pdf"
    _make_pdf(src)
    mp = tmp_path / "map.xlsx"
    _make_mapping(mp, [
        ("GridA", "001234", "Zhang"),   # 文本密码，前导零
        ("GridB", 8888, "Li"),          # 数字密码（openpyxl 读成 int/float）
        ("GridC", None, "Wang"),        # 空密码，应跳过
    ])
    out = tmp_path / "out"
    cfg = {
        "pdf_input_paths": [str(src)],
        "pdf_mapping_path": str(mp),
        "pdf_grid_column": "GridName",
        "pdf_password_column": "Pwd",
        "pdf_receiver_column": "Receiver",
        "pdf_watermark": True,
        "pdf_watermark_text": "{grid} {date}",
        "output_path": str(out),
    }
    return tmp_path, cfg


def _read_manifest(out_dir):
    wb = openpyxl.load_workbook(out_dir / MANIFEST_NAME)
    ws = wb.active
    rows = [[c.value for c in r] for r in ws.iter_rows(min_row=2)]
    wb.close()
    return rows


# ---------- 用例 ----------

def test_list_mapping_columns(workspace):
    tmp, cfg = workspace
    assert list_mapping_columns(cfg["pdf_mapping_path"]) == ["GridName", "Pwd", "Receiver"]


def test_read_mapping_password_as_text(workspace):
    tmp, cfg = workspace
    rows, warnings = read_mapping(cfg["pdf_mapping_path"], "GridName", "Pwd", "Receiver")
    assert [r["grid"] for r in rows] == ["GridA", "GridB"]
    assert rows[0]["password"] == "001234"      # 前导零不丢
    assert rows[1]["password"] == "8888"        # 不会变成 8888.0
    assert any("GridC" in w for w in warnings)  # 空密码警告

    with pytest.raises(ValueError):
        read_mapping(cfg["pdf_mapping_path"], "GridName", "不存在的列")


def test_encrypt_right_and_wrong_password(workspace):
    tmp, cfg = workspace
    out = run_pdf_dist(cfg, log_fn=lambda m: None)
    assert out

    r = PdfReader(tmp / "out" / "GridA" / "report.pdf")
    assert r.is_encrypted
    assert r.decrypt("001234") != 0
    assert "SOURCE-MARK page1" in r.pages[0].extract_text()

    r2 = PdfReader(tmp / "out" / "GridB" / "report.pdf")
    assert r2.decrypt("wrong") == 0
    assert r2.decrypt("8888") != 0

    # 空密码的 GridC 不产出
    assert not (tmp / "out" / "GridC").exists()


def test_watermark_on_and_off(workspace):
    tmp, cfg = workspace
    run_pdf_dist(cfg, log_fn=lambda m: None)
    r = PdfReader(tmp / "out" / "GridA" / "report.pdf")
    r.decrypt("001234")
    for page in r.pages:                       # 每页都有该网格的水印
        assert "GridA" in page.extract_text()

    cfg2 = dict(cfg, pdf_watermark=False, output_path=str(tmp / "out2"))
    run_pdf_dist(cfg2, log_fn=lambda m: None)
    r = PdfReader(tmp / "out2" / "GridA" / "report.pdf")
    r.decrypt("001234")
    assert "GridA" not in r.pages[0].extract_text()


def test_manifest_content(workspace):
    tmp, cfg = workspace
    run_pdf_dist(cfg, log_fn=lambda m: None)
    rows = _read_manifest(tmp / "out")
    assert len(rows) == 2                       # GridC 跳过，不进清单
    by_grid = {r[0]: r for r in rows}
    assert by_grid["GridA"][2] == "001234"      # 密码为字符串、前导零保留
    assert by_grid["GridA"][3] == "Zhang"
    assert by_grid["GridA"][4] == 2             # 页数
    assert "完成" in by_grid["GridB"][5]


def test_bad_filename_and_encrypted_source(tmp_path):
    """非法文件名字符的网格照常产出（净化）；自带密码的源 PDF 跳过但不中断。"""
    ok_src = tmp_path / "ok.pdf"
    _make_pdf(ok_src)
    locked = tmp_path / "locked.pdf"
    _make_pdf(locked, encrypt_pwd="x")
    mp = tmp_path / "map.xlsx"
    _make_mapping(mp, [('Grid/A:B?', "111", "")])
    logs = []
    cfg = {
        "pdf_input_paths": [str(ok_src), str(locked)],
        "pdf_mapping_path": str(mp),
        "pdf_grid_column": "GridName",
        "pdf_password_column": "Pwd",
        "pdf_watermark": False,
        "output_path": str(tmp_path / "out"),
    }
    out = run_pdf_dist(cfg, log_fn=logs.append)
    assert out
    assert any("locked.pdf" in m for m in logs)          # 带密码源被警告跳过
    safe_dir = tmp_path / "out" / "Grid_A_B_"
    assert (safe_dir / "ok.pdf").exists()                # 净化后的目录名


def test_stop_flag(workspace):
    tmp, cfg = workspace
    logs = []
    out = run_pdf_dist(cfg, log_fn=logs.append, stop_flag=lambda: True)
    assert out is None                                    # 一个网格都没做
    assert any("停止" in m for m in logs)

# ---------- 模板 / 随机密码 / 收尾摘要行（v2.7）----------

def test_write_mapping_template(tmp_path):
    """模板：首行表头三列，第二个 sheet 放说明（不影响 read_mapping 只读第一个 sheet）。"""
    from core.pdf_dist import write_mapping_template
    p = tmp_path / "tpl.xlsx"
    write_mapping_template(str(p))
    assert list_mapping_columns(str(p)) == ["网格", "密码", "接收人"]
    wb = openpyxl.load_workbook(p)
    assert "填写说明" in wb.sheetnames
    wb.close()

def test_fill_random_passwords_blanks_only(tmp_path):
    """只填空白的密码单元格；已手填的保留；原文件不被改动。"""
    from core.pdf_dist import fill_random_passwords
    mp = tmp_path / "m.xlsx"
    _make_mapping(mp, [("G1", None, ""), ("G2", "keepme", ""), ("G3", None, "")])
    out, n = fill_random_passwords(str(mp), "GridName", "Pwd")
    assert n == 2                       # 只补 G1 / G3
    assert out.endswith("_含密码.xlsx")
    rows, _ = read_mapping(out, "GridName", "Pwd")
    pw = {r["grid"]: r["password"] for r in rows}
    assert pw["G2"] == "keepme"         # 已填的不覆盖
    assert len(pw["G1"]) == 8 and len(pw["G3"]) == 8
    assert pw["G1"] != pw["G3"]
    rows0, _ = read_mapping(str(mp), "GridName", "Pwd")
    assert [r["grid"] for r in rows0] == ["G2"]   # 原文件保持原样

def test_fill_random_passwords_new_column(tmp_path):
    """清单没有密码列时新增「密码」列；密码字符集排除易混字符。"""
    from core.pdf_dist import fill_random_passwords, gen_password
    mp = tmp_path / "m2.xlsx"
    _make_mapping(mp, [("东区", "", ""), ("西区", "", "")], headers=("网格",))
    out, n = fill_random_passwords(str(mp), "网格")
    assert n == 2
    rows, _ = read_mapping(out, "网格", "密码")
    assert len(rows) == 2
    for _ in range(50):
        pwd = gen_password()
        assert len(pwd) == 8
        assert not set(pwd) & set("0O1lI")   # 无易混字符

def test_summary_line_pdf(workspace):
    """run_pdf_dist 收尾应打恰好一行 [SUMMARY] JSON。"""
    import json
    tmp, cfg = workspace
    logs = []
    out = run_pdf_dist(cfg, log_fn=logs.append)
    summary = [m for m in logs if m.startswith("[SUMMARY] ")]
    assert len(summary) == 1
    s = json.loads(summary[0][len("[SUMMARY] "):])
    assert s["mode"] == "pdf"
    assert s["grids_ok"] == 2           # GridA / GridB（GridC 空密码被跳过）
    assert s["grids_fail"] == 0
    assert s["files"] == 2
    assert s["output"] == out
    assert s["manifest"]
