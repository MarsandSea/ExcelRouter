"""
excelrouter skill 的公共小工具：把 vendor/core 接到 sys.path、统一的默认配置、
统一的 stdout/stderr 输出约定、友好的依赖缺失提示。

Copyright (c) 2026 AbeLin
MIT License

约定（所有 er_*.py 脚本都遵守，Claude 靠这个约定解析结果）：
  - 日志/进度/警告一律写 stderr（人类可读，随便打印）
  - 最终结果只有一行 JSON，写 stdout（用 emit()），不要在这行前后混入别的 print
  - 失败退出码非 0，且 stdout 仍然是一行 JSON（含 "error" 字段），不是裸 Python traceback
"""

import importlib
import importlib.util
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
    # v3.0：隐藏的 sheet 默认不发、隐藏列清空（仍隐藏）、隐藏的行照常发；True = 全部照样发出
    "keep_hidden": False,
    "pdf_input_paths": [],
    "pdf_mapping_path": "",
    "pdf_grid_column": "",
    "pdf_password_column": "",
    "pdf_receiver_column": "",
    "pdf_encrypt": True,
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


# ---------- 方案（v3.0）：桌面版「📋 方案」存的小 JSON，技能也能读、能存，格式只有一份 ----------

def add_recipe_args(ap):
    ap.add_argument("--recipe", default="", help="套用一个方案文件（桌面版「📋 方案」存的 .json）；命令行明确给的参数优先")
    ap.add_argument("--save-recipe", default="", help="把这次实际用的设置存成方案文件（.json 路径）")


def _explicit_dests(ap, argv):
    """命令行上明确写了的参数（方案不能覆盖它们）。"""
    out = set()
    for action in ap._actions:
        for opt in action.option_strings:
            if any(a == opt or a.startswith(opt + "=") for a in argv):
                out.add(action.dest)
    return out


def apply_recipe_args(ap, args, mode, argmap, argv=None):
    """--recipe：把方案里的设置填进还没在命令行上写的参数。返回 (方案, 说明, 没有对应参数的设置)。

    argmap：{方案里的键: (参数 dest, 转换函数(值, 全部设置) -> 参数值)}。模式对不上、文件坏了直接报错退出。
    """
    if not getattr(args, "recipe", ""):
        return None, [], {}
    from core import recipe as R
    try:
        rec, notes = R.load_recipe(os.path.abspath(args.recipe), mode=mode)
    except R.RecipeError as e:
        emit_error(f"方案用不了：{e}")
    explicit = _explicit_dests(ap, sys.argv[1:] if argv is None else argv)
    extra = {}
    for k, v in rec["settings"].items():
        if k in argmap:
            dest, conv = argmap[k]
            if dest not in explicit:
                setattr(args, dest, conv(v, rec["settings"]))
        else:
            extra[k] = v
    return rec, notes, extra


def save_recipe_from(args, config, mode, columns=None, input_name=""):
    """--save-recipe：把这次实际用的设置存成方案，返回路径（没要求存时返回 ""）。"""
    if not getattr(args, "save_recipe", ""):
        return ""
    from core import recipe as R
    path = os.path.abspath(args.save_recipe)
    if not path.lower().endswith(".json"):
        path += ".json"
    name = os.path.splitext(os.path.basename(path))[0]
    try:
        R.save_recipe(R.make_recipe(config, mode, name, columns=columns, input_name=input_name), path)
    except Exception as e:
        eprint(f"⚠️ 方案没存成：{e}")
        return ""
    return path


def _join(v, _s=None):
    return ",".join(str(x) for x in v)


def eprint(msg):
    """日志/进度输出到 stderr，不污染 stdout 的结果 JSON。"""
    print(msg, file=sys.stderr, flush=True)


_SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 国内办公网络直连 pypi.org 很慢（v2.x 时实测 pandas+numpy 两分钟没装完）。按顺序试，失败即换下一个。
# 2026-10 公司网络实测索引响应：腾讯云 0.3s、阿里云 1.5s、清华 6s（pip 直接报 versions: none）、pypi.org 5.7s。
_PIP_MIRRORS = (
    "https://mirrors.cloud.tencent.com/pypi/simple",
    "https://mirrors.aliyun.com/pypi/simple",
)
_PIP_MIRROR = _PIP_MIRRORS[0]  # 报错提示里给用户手动用的那个

# v3.0 起拆分也不用 pandas 了：拆分、汇总、查收都只要这两个
_LITE_DEPS = (("openpyxl", "openpyxl"), ("xlrd", "xlrd"))
_EXCEL_DEPS = _LITE_DEPS
_PDF_DEPS = (("pypdf", "pypdf"), ("fpdf", "fpdf2"), ("cryptography", "cryptography"))
# 给 Word / Excel / PPT 加密才要（也写在 requirements-pdf.txt 里）；只发 PDF 时不检查，
# 免得内网装不上它的用户每次跑 PDF 分发都去试一次安装
_OFFICE_DEPS = (("msoffcrypto", "msoffcrypto-tool"),)


def _missing(deps):
    return [pip for mod, pip in deps if importlib.util.find_spec(mod) is None]


# 宿主 Agent 会掐断跑太久的命令（WorkBuddy 实测 120 秒 SIGTERM）。自动安装必须在这之前
# 自己收手并给出「再跑一次」的明确结论，否则 AI 只看到一个被杀掉的空结果。
_INSTALL_BUDGET = 95


def _pip_install(req_file, packages=None):
    """用**当前解释器**装依赖（不用裸 pip：多个 Python 并存时 pip 常装到别的环境里，
    装完照样报缺——WorkBuddy 就自带一个独立的 Python）。

    返回 "ok" / "timeout" / "fail"。packages 给了就只装这几个包（版本要求照 req_file 里写的）。

    **国内镜像优先**：用户几乎都在国内，直连 pypi.org 实测很慢（v2.x 装 pandas+numpy 两分钟都没装完、
    被宿主杀掉）；用户若自己配过 pip 源（pip.ini / PIP_INDEX_URL）就尊重他的配置。
    两轮共用一个总时限，超时不算失败：已下载的包在 pip 缓存里，再跑一次会快很多。"""
    import subprocess
    import time
    target = ["-r", req_file] if not packages else _pinned(req_file, packages)
    base = [sys.executable, "-m", "pip", "install", "--disable-pip-version-check",
            "--no-warn-script-location", "--prefer-binary", "--timeout", "15", "--retries", "1",
            "-q"] + target
    user_index = bool(os.environ.get("PIP_INDEX_URL")) or _has_pip_config()
    rounds = [[]] if user_index else [["-i", m] for m in _PIP_MIRRORS] + [[]]
    deadline = time.time() + _INSTALL_BUDGET
    for extra in rounds:
        left = deadline - time.time()
        if left < 10:
            return "timeout"
        try:
            r = subprocess.run(base + extra, stdout=sys.stderr, stderr=sys.stderr, timeout=left)
        except subprocess.TimeoutExpired:
            return "timeout"
        except OSError:
            continue
        if r.returncode == 0:
            return "ok"
    return "fail"


def _pinned(req_file, packages):
    """从 requirements 文件里取这几个包的版本要求（如 openpyxl>=3.1.0），取不到就只写包名。"""
    import re
    specs = {}
    try:
        with open(req_file, encoding="utf-8") as f:
            for line in f:
                line = line.split("#", 1)[0].strip()
                m = re.match(r"^([A-Za-z0-9_.-]+)", line)
                if m:
                    specs[m.group(1).lower()] = line
    except OSError:
        pass
    return [specs.get(p.lower(), p) for p in packages]


def _has_pip_config():
    """用户是否给 pip 配过源（有就别拿镜像覆盖他的选择，可能是公司内网源）。"""
    import subprocess
    try:
        out = subprocess.run([sys.executable, "-m", "pip", "config", "list"],
                             capture_output=True, text=True, timeout=15).stdout
    except Exception:
        return False
    return "index-url" in out


def ensure_deps(need_pdf=False, need_office=False, lite=False):
    """缺依赖就自动装上（首次使用），装不上再报人话错误。

    **必须在 import openpyxl / core.* 之前调用**——否则缺包时 Python 直接甩
    ImportError 堆栈，这里的人话提示根本走不到（旧版 check_deps 就是这么失效的）。

    为什么默认自动装：用户的诉求是「把表拆了」，不是「配环境」。首次使用卡在
    pip 上是这个技能最大的流失点；装的只是 requirements*.txt 里声明的包。
    不想自动装（内网、受管环境）就设环境变量 ER_NO_AUTO_INSTALL=1。

    Excel 这一组（拆分、汇总、查收）v3.0 起都只要 openpyxl + xlrd，只装缺的那几个，几秒钟。
    lite 是旧参数，现在和默认一样，留着给已有调用方。
    """
    groups = [(_EXCEL_DEPS, "requirements.txt")]
    if need_pdf:
        groups.append((_PDF_DEPS, "requirements-pdf.txt"))
    if need_office:
        groups.append((_OFFICE_DEPS, "requirements-pdf.txt"))
    for deps, req in groups:
        _install_group(deps, req, only=deps is _EXCEL_DEPS)


def _install_group(deps, req, only=False):
    """检查一组依赖，缺了就装；装不上 emit_error 退出。only=True 时只装缺的那几个包，不装整份 req。"""
    missing = _missing(deps)
    # cryptography 缺失只是加密降级（AES-256 → RC4-128），不值得为它阻塞
    hard = [m for m in missing if m != "cryptography"]
    if not missing:
        return
    req_path = os.path.join(_SKILL_DIR, req)
    auto = os.environ.get("ER_NO_AUTO_INSTALL") != "1"
    if auto:
        eprint(f"首次使用，正在自动安装依赖：{', '.join(missing)}（{'几秒钟' if only else '约 1 分钟'}，只装一次）…")
        status = _pip_install(req_path, missing if only else None)
        importlib.invalidate_caches()  # 否则刚装好的包在本进程里仍然 find_spec 不到
        if not [m for m in _missing(deps) if m != "cryptography"]:
            eprint("✅ 依赖安装完成")
            return
        if status == "timeout":
            emit_error(
                "首次使用正在安装运行依赖，网络较慢还没装完（已下载的部分会保留）。"
                "直接再运行一次同一条命令即可接着装，通常第二次很快。",
                retry=True, installing=missing)
    if hard:
        how = f'"{sys.executable}" -m pip install ' + (" ".join(hard) if only else f'-r "{req_path}"')
        emit_error(
            f"缺少运行依赖：{', '.join(hard)}"
            f"{'，自动安装没成功' if auto else '（ER_NO_AUTO_INSTALL=1，未自动安装）'}。请手动运行："
            f" {how}（网络慢可在末尾加 -i {_PIP_MIRROR}）",
            missing=hard)


# 旧名字保留给外部调用方
check_deps = ensure_deps


def sheet_headers(file_path, config):
    """返回 [(sheet_name, header_row_1based, [columns]), ...]，覆盖文件里每一个 sheet。
    只读每个 sheet 前 20 行，快。调用前须已 ensure_deps()。"""
    from core.splitter import sheet_headers as _core_sheet_headers
    return _core_sheet_headers(file_path, config)


def match_column(columns, target):
    """与内核 _find_col 同一规则（先精确、再包含），但把**所有**包含命中都报出来：
    返回 (实际命中的列名或 None, 其余也命中的列名列表)。内核拿第一个包含命中就走，
    「部门」同时撞上「所属部门」「部门编码」时用户却不知道按的是哪列。"""
    from core.utils import soft_clean
    t = soft_clean(target)
    if not t:
        return None, []
    if t in columns:
        return t, []
    hits = [c for c in columns if c and (t in c or c in t)]
    return (hits[0], hits[1:]) if hits else (None, [])


def default_output_for(input_path, leaf="拆分结果"):
    """输出目录默认值，与桌面版 _auto_output_for 同一规则：
    文件 → 同目录/拆分结果；文件夹 → 文件夹里面/拆分结果（结果永远在数据旁边；
    放在输入里是安全的，run_split 会整体跳过 output_root 子树）。"""
    p = os.path.abspath(input_path)
    base = p if os.path.isdir(p) else os.path.dirname(p)
    return os.path.join(base, leaf)


def open_folder(path):
    """在用户的文件管理器里打开目录（桌面版「跑完自动弹出结果」的等价物）。
    失败不抛异常，返回 bool——调用方据此告诉 AI 要不要自己把路径念给用户。"""
    import subprocess
    try:
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
            return True
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
            return True
        for cmd in ("xdg-open", "gio", "peony", "caja", "nautilus", "dde-file-manager"):
            exe = __import__("shutil").which(cmd)
            if exe:
                args = [exe, "open", path] if cmd == "gio" else [exe, path]
                subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                return True
    except Exception:
        pass
    return False


def list_outputs(root, exts=("zip", "xlsx", "pdf")):
    """扫产出目录里的真实文件（绝对路径列表）：ZIP 排最前（批量时它就是
    「拿去发」的交付单位），其余按路径排序。取前几个给 AI 直接 present_files。"""
    found = []
    for dirpath, _d, names in os.walk(root):
        for fn in names:
            low = fn.lower()
            if low.startswith("~$") or low.endswith("__tmp__.xlsx"):
                continue
            ext = low.rsplit(".", 1)[-1] if "." in low else ""
            if ext in exts:
                found.append(os.path.join(dirpath, fn))
    rank = {e: i for i, e in enumerate(exts)}
    found.sort(key=lambda f: (rank.get(f.lower().rsplit(".", 1)[-1], 99), f))
    return found
