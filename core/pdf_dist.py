"""
Excel 智能拆分工具 - PDF 按网格加密分发
Copyright (c) 2026 AbeLin
MIT License

把同一批 PDF 按「网格 → 密码」映射清单，为每个网格生成一份
带专属打开密码 + 专属水印（可溯源）的副本，并输出分发清单。

依赖：pypdf（读写/加密）、fpdf2（生成水印页）、cryptography（AES-256 后端，
缺失时降级 RC4-128 并警告）。openpyxl 读映射表、写分发清单。
"""

import io
import os
import time

from openpyxl import Workbook, load_workbook
from pypdf import PdfReader, PdfWriter

from core.utils import safe_filename, soft_clean

MANIFEST_NAME = "分发清单.xlsx"

# 水印字体候选（只收 .ttf——fpdf2 不支持 .ttc，msyh.ttc 之类直接跳过）
_FONT_CANDIDATES = ["simhei.ttf", "Deng.ttf", "msyhbd.ttf", "simfang.ttf", "simkai.ttf"]


def _noop(*_args, **_kwargs):
    pass


def list_mapping_columns(xlsx_path):
    """读映射清单首个 sheet 的第 1 行表头，供 GUI 下拉选择网格列/密码列/接收人列。

    映射清单是用户专门维护的小表，默认表头就在第 1 行，不做启发式识别。
    """
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            return [soft_clean(v) for v in row if soft_clean(v)]
        return []
    finally:
        wb.close()


def _cell_str(v):
    """单元格值转字符串：数字型密码防止 openpyxl 读成 1234.0 / 丢前导零后的 int。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return soft_clean(v)


def read_mapping(xlsx_path, grid_col, password_col, receiver_col=""):
    """读「网格 → 密码（→ 接收人）」映射。

    返回 (rows, warnings)：rows 为 [{"grid", "password", "receiver"}]，
    warnings 为需要记进日志的问题（空网格/空密码行跳过、重复网格覆盖）。
    """
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        header = None
        idx = {}
        rows, warnings = [], []
        seen = {}
        for r, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if header is None:
                header = [soft_clean(v) for v in row]
                for name, col in ((grid_col, "grid"), (password_col, "password"),
                                  (receiver_col, "receiver")):
                    if name and name in header:
                        idx[col] = header.index(name)
                if "grid" not in idx or "password" not in idx:
                    missing = [n for n, c in ((grid_col, "grid"), (password_col, "password"))
                               if c not in idx]
                    raise ValueError(f"映射清单里找不到列：{'、'.join(missing)}")
                continue
            grid = _cell_str(row[idx["grid"]] if idx["grid"] < len(row) else None)
            pwd = _cell_str(row[idx["password"]] if idx["password"] < len(row) else None)
            recv = ""
            if "receiver" in idx and idx["receiver"] < len(row):
                recv = _cell_str(row[idx["receiver"]])
            if not grid and not pwd:
                continue                      # 整行空白，静默跳过
            if not grid:
                warnings.append(f"第 {r} 行网格为空，已跳过")
                continue
            if not pwd:
                warnings.append(f"第 {r} 行「{grid}」密码为空，已跳过")
                continue
            if grid in seen:
                warnings.append(f"网格「{grid}」重复出现，以第 {r} 行为准")
                rows[seen[grid]] = {"grid": grid, "password": pwd, "receiver": recv}
                continue
            seen[grid] = len(rows)
            rows.append({"grid": grid, "password": pwd, "receiver": recv})
        return rows, warnings
    finally:
        wb.close()


def _find_cjk_font():
    """在 Windows 系统字体目录里找一款可用的中文 .ttf。找不到返回 None。"""
    fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    for name in _FONT_CANDIDATES:
        p = os.path.join(fonts_dir, name)
        if os.path.exists(p):
            return p
    return None


_wm_cache = {}


def _make_watermark_pdf(text, w_pt, h_pt, *, angle=45, opacity=0.15, font_path=None):
    """用 fpdf2 生成一页平铺斜排水印 PDF，返回字节串。按（尺寸+文字）缓存。

    没有中文字体时退回内置 Helvetica，非 ASCII 字符替换成 '?'（上层已警告过）。
    """
    key = (round(w_pt), round(h_pt), text, angle, opacity, font_path)
    if key in _wm_cache:
        return _wm_cache[key]

    from fpdf import FPDF
    pdf = FPDF(unit="pt", format=(w_pt, h_pt))
    pdf.set_auto_page_break(False)
    pdf.add_page()
    if font_path:
        pdf.add_font("wm", "", font_path)
        pdf.set_font("wm", size=18)
        draw_text = text
    else:
        pdf.set_font("Helvetica", size=18)
        draw_text = text.encode("ascii", "replace").decode("ascii")
    pdf.set_text_color(128, 128, 128)
    text_w = pdf.get_string_width(draw_text)
    step_x = max(text_w + 60, 180)
    step_y = 120
    diag = (w_pt ** 2 + h_pt ** 2) ** 0.5
    with pdf.local_context(fill_opacity=opacity):
        # 绕页面中心旋转后铺一个盖满对角线的网格，保证旋转后无死角
        with pdf.rotation(angle, w_pt / 2, h_pt / 2):
            y = h_pt / 2 - diag / 2
            row = 0
            while y < h_pt / 2 + diag / 2:
                x = w_pt / 2 - diag / 2 + (step_x / 2 if row % 2 else 0)
                while x < w_pt / 2 + diag / 2:
                    pdf.text(x, y, draw_text)
                    x += step_x
                y += step_y
                row += 1
    data = bytes(pdf.output())
    _wm_cache[key] = data
    return data


def _encrypt_algorithm(log_fn):
    """优先 AES-256；cryptography 缺失时降级 RC4-128 并警告（不中断）。"""
    try:
        import cryptography  # noqa: F401
        return "AES-256"
    except ImportError:
        log_fn("⚠ 未安装 cryptography，加密降级为 RC4-128（强度较弱，建议 pip install cryptography）")
        return "RC4-128"


def _stamp_and_encrypt(reader, out_path, wm_text, password, algorithm,
                       *, opacity=0.15, angle=45, font_path=None):
    """把 reader 的所有页盖上水印、设打开密码，写到 out_path。返回页数。

    writer 每个网格都要新建：pypdf 的 merge_page 会就地修改页对象，
    复用会导致水印跨网格叠加。append(reader) 会把页深拷贝进 writer，
    所以改 writer.pages 不会污染共享的 reader。
    """
    writer = PdfWriter()
    writer.append(reader)
    if wm_text:
        wm_readers = {}
        for page in writer.pages:
            w, h = float(page.mediabox.width), float(page.mediabox.height)
            k = (round(w), round(h))
            if k not in wm_readers:
                data = _make_watermark_pdf(wm_text, w, h, angle=angle,
                                           opacity=opacity, font_path=font_path)
                wm_readers[k] = PdfReader(io.BytesIO(data))
            page.merge_page(wm_readers[k].pages[0])
    writer.encrypt(user_password=password, algorithm=algorithm)
    with open(out_path, "wb") as f:
        writer.write(f)
    return len(writer.pages)


def write_manifest(entries, out_path):
    """写分发清单：网格 | 文件名 | 密码 | 接收人 | 页数 | 状态。

    密码列强制文本格式，防止 Excel 把 001234 显示成 1234。
    清单含明文密码——只给分发人自己用，别随文件一起发出去。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "分发清单"
    ws.append(["网格", "文件名", "密码", "接收人", "页数", "状态"])
    for e in entries:
        ws.append([e["grid"], e.get("file", ""), e["password"],
                   e.get("receiver", ""), e.get("pages", ""), e.get("status", "")])
        ws.cell(row=ws.max_row, column=3).number_format = "@"
    widths = {"A": 18, "B": 40, "C": 16, "D": 14, "E": 8, "F": 24}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    wb.save(out_path)


def run_pdf_dist(config, log_fn=None, progress_fn=None, stop_flag=None):
    """PDF 加密分发主流程。签名与 run_split 一致，GUI 队列泵直接复用。

    单个网格/文件失败不中断整体，记日志继续（项目错误处理约定）。
    返回输出根目录；一个都没成功时返回 None。
    """
    log_fn = log_fn or _noop
    progress_fn = progress_fn or _noop
    stop_flag = stop_flag or (lambda: False)

    pdf_paths = [p for p in config.get("pdf_input_paths", []) if p]
    mapping_path = config.get("pdf_mapping_path", "")
    output_root = config.get("output_path", "")
    if not pdf_paths:
        log_fn("❌ 没有选择要分发的 PDF 文件")
        return None
    for p in pdf_paths:
        if not os.path.exists(p):
            log_fn(f"❌ 找不到 PDF：{p}")
            return None
    if not mapping_path or not os.path.exists(mapping_path):
        log_fn("❌ 找不到映射清单文件")
        return None
    if not output_root:
        output_root = os.path.join(os.path.dirname(pdf_paths[0]), "分发结果")

    rows, warnings = read_mapping(
        mapping_path,
        config.get("pdf_grid_column", ""),
        config.get("pdf_password_column", ""),
        config.get("pdf_receiver_column", ""),
    )
    for w in warnings:
        log_fn(f"⚠ 映射清单：{w}")
    if not rows:
        log_fn("❌ 映射清单里没有有效的「网格 + 密码」行")
        return None
    log_fn(f"📋 映射清单：{len(rows)} 个网格；待分发 PDF：{len(pdf_paths)} 个")

    watermark_on = config.get("pdf_watermark", True)
    wm_template = config.get("pdf_watermark_text", "{grid} {date}") or "{grid} {date}"
    opacity = float(config.get("pdf_watermark_opacity", 0.15))
    angle = float(config.get("pdf_watermark_angle", 45))
    font_path = None
    if watermark_on:
        font_path = _find_cjk_font()
        if not font_path:
            log_fn("⚠ 未找到中文字体（simhei.ttf 等），水印中的中文会显示为 '?'")

    algorithm = _encrypt_algorithm(log_fn)
    os.makedirs(output_root, exist_ok=True)
    date_str = time.strftime("%Y-%m-%d")

    # 每个源 PDF 只开一个 reader，跨网格复用（writer 每网格重建，见 _stamp_and_encrypt）
    readers = {}
    for p in pdf_paths:
        try:
            r = PdfReader(p)
            if r.is_encrypted:
                log_fn(f"⚠ 「{os.path.basename(p)}」本身带密码，无法处理，已跳过（请先解密另存）")
                continue
            readers[p] = r
        except Exception as e:
            log_fn(f"⚠ 读取「{os.path.basename(p)}」失败：{e}，已跳过")
    if not readers:
        log_fn("❌ 没有可处理的 PDF")
        return None

    entries = []
    total = len(rows) * len(readers)
    done = 0
    ok_grids, fail_grids = 0, 0
    stopped = False
    for row in rows:
        if stop_flag():
            log_fn("⏹ 已停止，剩余网格未处理")
            stopped = True
            break
        grid, password, receiver = row["grid"], row["password"], row["receiver"]
        grid_dir = os.path.join(output_root, safe_filename(grid))
        wm_text = wm_template.replace("{grid}", grid).replace("{date}", date_str) \
            if watermark_on else ""
        grid_ok = True
        for src, reader in readers.items():
            base = safe_filename(os.path.basename(src))
            out_path = os.path.join(grid_dir, base)
            try:
                os.makedirs(grid_dir, exist_ok=True)
                pages = _stamp_and_encrypt(reader, out_path, wm_text, password, algorithm,
                                           opacity=opacity, angle=angle, font_path=font_path)
                entries.append({"grid": grid, "password": password, "receiver": receiver,
                                "file": os.path.relpath(out_path, output_root),
                                "pages": pages, "status": "✅ 完成"})
            except Exception as e:
                grid_ok = False
                log_fn(f"❌ 「{grid}」× {base} 失败：{e}")
                entries.append({"grid": grid, "password": password, "receiver": receiver,
                                "file": os.path.relpath(out_path, output_root),
                                "status": f"失败：{e}"})
            done += 1
            progress_fn(done / total * 0.95)
        if grid_ok:
            ok_grids += 1
            log_fn(f"🔐 {grid} → {len(readers)} 个文件已加密")
        else:
            fail_grids += 1

    manifest_path = os.path.join(output_root, MANIFEST_NAME)
    try:
        write_manifest(entries, manifest_path)
        log_fn(f"📄 分发清单已生成：{MANIFEST_NAME}")
    except Exception as e:
        log_fn(f"⚠ 分发清单写入失败：{e}")
    progress_fn(1.0)

    if stopped:
        log_fn(f"⏹ 已停止：完成 {ok_grids} 个网格")
    else:
        log_fn(f"✅ 完成：{ok_grids} 个网格成功" + (f"，{fail_grids} 个失败" if fail_grids else ""))
    log_fn("🔒 提醒：分发清单里有明文密码，只留给自己用，不要随文件一起发出去")
    return output_root if ok_grids else None
