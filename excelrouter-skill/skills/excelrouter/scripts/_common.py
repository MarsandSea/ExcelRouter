"""
excelrouter skill 的公共小工具：把 vendor/core 接到 sys.path、统一的默认配置、
统一的 stdout/stderr 输出约定、友好的依赖缺失提示。

Copyright (c) 2026 AbeLin
MIT License

约定（三个 er_*.py 脚本都遵守，Claude 靠这个约定解析结果）：
  - 日志/进度/警告一律写 stderr（人类可读，随便打印）
  - 最终结果只有一行 JSON，写 stdout（用 emit()），不要在这行前后混入别的 print
  - 失败退出码非 0，且 stdout 仍然是一行 JSON（含 "error" 字段），不是裸 Python traceback
"""

import json
import os
import sys

_VENDOR_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor")

# 上游 config/default_config.json 的镜像（截至上游 v2.7.0 tag），外加本 CLI 自有的
# no_timestamp（见文件末尾注释）。之所以在这里重新声明一份而不是运行时读上游的
# default_config.json，是因为上游那份文件是 GUI 专用路径（gui/app.py 用 sys._MEIPASS
# 找它），本 CLI 不依赖上游目录结构，只依赖 vendor/core 这一个包。
#
# 每次同步上游后都要对照 excel-router 仓库的 config/default_config.json 核一遍。
# 【有意省略的两个上游字段】auto_open_output、ui_mode 是纯 GUI 行为（跑完自动打开输出
# 目录、界面当前处于哪个模式），命令行形态用不上，CLI 侧对应行为见 SKILL.md 的
# 「跑完之后必须把结果交给用户」一节。不要为了"对齐上游"把它们补进来。
DEFAULTS = {
    "input_path": "",
    "output_path": "",
    "header_mode": "auto",
    "header_row": 1,
    "grid_keys": [],
    "id_keys": [],
    "split_column": "",
    "selected_values": [],
    "person_column": "",
    "to_person": False,
    "person_file_filter": [],
    "value_alias_map": {},
    "skip_values": ["合计", "小计", "总计", "平均", ""],
    "merge_across_files": False,
    "make_zip": True,
    # no_timestamp 是本 CLI 自有的字段，上游 core 里没有（内核无条件产出
    # {输出根}/{MMDDHHMM}结果 子目录）。它由 scripts 包装层的 flatten_output() 实现：
    # 内核跑完后把内容提升到 --output 根、再删掉时间戳目录。传给 core 也不报错，
    # core 只是忽略它——真正的扁平化动作发生在包装层，见 er_split.py。
    "no_timestamp": False,
    "exact_match": True,
    "preserve_format": True,
    "keep_formulas": False,
    "pdf_input_paths": [],
    "pdf_mapping_path": "",
    "pdf_grid_column": "",
    "pdf_password_column": "",
    "pdf_receiver_column": "",
    "pdf_watermark": True,
    "pdf_watermark_text": "{grid} {date}",
    "pdf_watermark_opacity": 0.15,
    "pdf_watermark_angle": 45,
}


def bootstrap():
    """把 vendor/ 插进 sys.path，让 `import core.splitter` 生效；同时把 stdout/stderr
    锁定为 UTF-8。三个 er_*.py 的输出全是中文（字段名、日志、JSON 里的中文取值），
    但 Windows 控制台/子进程管道默认走系统代码页（GBK 等），不强制 reconfigure 的话，
    脚本被当子进程调用时会在管道另一端解码失败——这不是理论风险，是本仓 CI 首次跑
    测试就实测踩到的坑。幂等，可重复调用。"""
    if _VENDOR_DIR not in sys.path:
        sys.path.insert(0, _VENDOR_DIR)
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8")
            except Exception:
                pass


def build_config(**overrides):
    """DEFAULTS 深拷贝后按 overrides 覆盖。绝不用 dict.get 兜底默认值——
    上游 core/splitter.py 内部 .get 的兜底值和 config/default_config.json 不完全一致
    （例如 merge_across_files 前者兜底 True、后者默认 False），CLI 必须显式传全量字段，
    不能依赖 core 内部兜底，否则命令行版本和桌面版默认行为会对不上。"""
    cfg = json.loads(json.dumps(DEFAULTS))  # 简单深拷贝，DEFAULTS 都是基础类型
    for k, v in overrides.items():
        if v is not None:
            cfg[k] = v
    return cfg


def emit(obj):
    """结果打一行 JSON 到 stdout。整个进程生命周期只应该调用一次。"""
    sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def emit_error(message, **extra):
    """标准化的失败输出：stdout 仍是合法 JSON（带 error 字段），退出码 1。"""
    obj = {"ok": False, "error": message}
    obj.update(extra)
    emit(obj)
    sys.exit(1)


def eprint(msg):
    """日志/进度输出到 stderr，不污染 stdout 的结果 JSON。"""
    print(msg, file=sys.stderr, flush=True)


def check_deps(need_pdf=False):
    """提前体检依赖，缺了给出人话提示而不是甩一截 ImportError 堆栈。"""
    missing = []
    for mod, pip_name in (("pandas", "pandas"), ("openpyxl", "openpyxl")):
        try:
            __import__(mod)
        except ImportError:
            missing.append(pip_name)
    if missing:
        emit_error(
            f"缺少运行依赖：{', '.join(missing)}。请先在 skills/excelrouter/ 目录下运行："
            f" pip install -r requirements.txt"
        )
    if need_pdf:
        pdf_missing = []
        for mod, pip_name in (("pypdf", "pypdf"), ("fpdf", "fpdf2")):
            try:
                __import__(mod)
            except ImportError:
                pdf_missing.append(pip_name)
        if pdf_missing:
            emit_error(
                f"PDF 加密分发缺少依赖：{', '.join(pdf_missing)}。"
                f" 请运行：pip install -r requirements.txt"
                f"（cryptography 缺失不影响运行，只是加密强度会从 AES-256 降级到 RC4-128）"
            )
