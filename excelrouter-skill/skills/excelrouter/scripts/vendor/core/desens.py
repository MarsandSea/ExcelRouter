"""数据脱敏：表格里的手机号、证件号、银行卡号、姓名、地址等个人信息，按用途打码 / 换代号 / 泛化 / 清空，处理完再复检。

Copyright (c) 2026 AbeLin · MIT License

    job = Job.from_config({"scene": "发群通报"})        # 空 / 关掉 → None
    job.sheet(ws, h, last, names, where)                 # 识别 + 在内存里改（源文件不动）
    hits = job.scan_files(paths)                         # 落盘后重扫：处理过的类型还有完整号码，就是没洗干净
    job.write_report(path, files)                        # 脱敏报告：只有列名、类型、格数和虚构示例
    run_desens(config, log_fn, progress_fn, stop_flag)   # 单独脱敏一个表 / 一个文件夹，签名同 run_split

识别靠规则：列名关键词 + 抽样校验（身份证校验码和出生日期、银行卡 Luhn、手机号段），不保证找全，所以有复检兜底；
复检不过就不交付（单独脱敏时不过的文件根本不落盘），宁可没产出，不出错产出。
报告、日志、[SUMMARY] 里只有列名、类型、个数和虚构示例，没有任何真实值。

处理方式对应 GB/T 37964-2019 的去标识化技术：打码＝部分屏蔽；换代号＝假名化（HMAC-SHA256，同一份种子下
同类型同值同代号、跨文件一致，没有种子推不回原值）；泛化＝地址到区县、日期到年、坐标降精度、金额分段；清空＝抑制。

规矩同 xlio：不导入 pandas，不碰删文件、起子进程、读环境变量这类调用，注释和字符串避开内部打包的禁用字眼。
"""
import base64
import datetime
import hashlib
import hmac
import io
import json
import os
import re
import time

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.packaging.core import DocumentProperties
from openpyxl.utils import get_column_letter

from core.utils import safe_filename, soft_clean
from core.xlio import (HEAD_SCAN, _iter_sheets_ro, _no_formulas, col_names, hidden_info, key_value, open_book, put,
                       resolve_header_row)

SAMPLE_ROWS = 200                   # 识别时每列看前这么多行
RESULT_DIR = "脱敏结果"              # 单独脱敏的结果固定放在输入旁边的这个文件夹
REPORT_NAME = "脱敏报告.xlsx"
FAIL_NOTE = "未通过复检_请勿发出.txt"

METHODS = {"mask": "打码", "pseudo": "换代号", "generalize": "泛化", "clear": "清空", "keep": "照发"}
SCENES = ("内部下发", "发群通报", "交给AI分析", "对外报送")
SCENE_TEXT = {
    "内部下发": "证件号、银行卡号打码，其余照发（客户经理外呼要完整号码）",
    "发群通报": "识别出的个人信息全部打码，文字里夹带的号码也打码",
    "交给AI分析": "姓名、号码、证件、编码换成一致的代号（还能统计、关联，认不出人），地址到区县、日期到年",
    "对外报送": "个人信息整列清空，编码换成代号",
    "自定义": "按指定的列处理，没指定的照发；文字里夹带的手机号、证件号、银行卡号也打码",
}
# 能在一段文字里认出来、也能在复检里扫出来的三类
EMBED_KINDS = ("身份证号", "银行卡号", "手机号")

# =====================================================
# 校验：单看一个值像不像
# =====================================================

_ID_W = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
_ID_CHECK = "10X98765432"
_BLANKS = {"", "无", "-", "/", "—", "－", "暂无", "空", "null", "none", "n/a", "#n/a", "na"}
_CJK = re.compile(r"[㐀-鿿]")
_EMAIL = re.compile(r"^[A-Za-z0-9._%+]+@[A-Za-z0-9.]+\.[A-Za-z]{2,}$")
_LANDLINE = re.compile(r"^0[1-9]\d{8,10}$")
_DATEISH = re.compile(r"^(1[89]|20)\d{2}\s*[-/.年]\s*\d{1,2}")


def _text(v):
    """格子的值 → 文字（3.0 当 3，错误格当空）。"""
    v = key_value(v)
    return "" if v is None or isinstance(v, bool) else soft_clean(v)


def _norm(v):
    """拿来校验号码的写法：去掉空格和连字符。日期不当号码。"""
    if isinstance(v, (datetime.date, datetime.time)):
        return ""
    return _text(v).replace(" ", "").replace("-", "")


def _digits(s):
    return s.isascii() and s.isdigit()


def _filled(v):
    return _text(v).lower() not in _BLANKS


def _valid_date(y, m, d):
    try:
        datetime.date(y, m, d)
    except ValueError:
        return False
    return 1900 <= y <= datetime.date.today().year


def _id_check_digit(s17):
    return _ID_CHECK[sum(int(c) * w for c, w in zip(s17, _ID_W)) % 11]


def is_id_card(s):
    """18 位居民身份证号：出生日期合法、校验码对得上。"""
    if len(s) != 18 or not _digits(s[:17]):
        return False
    if not _valid_date(int(s[6:10]), int(s[10:12]), int(s[12:14])):
        return False
    return _id_check_digit(s[:17]) == s[17].upper()


def _is_id15(s):
    """15 位老身份证（没有校验码，只看出生日期），列名对上时才认。"""
    return (len(s) == 15 and _digits(s) and s[0] in "1234568"
            and _valid_date(1900 + int(s[6:8]), int(s[8:10]), int(s[10:12])))


def _strip86(s):
    if s.startswith("+86"):
        return s[3:]
    if s.startswith("86") and len(s) == 13:
        return s[2:]
    return s


def is_mobile(s):
    s = _strip86(s)
    return len(s) == 11 and _digits(s) and s[0] == "1" and s[1] in "3456789"


def _luhn(s):
    total = 0
    for i, c in enumerate(reversed(s)):
        d = int(c)
        if i % 2:
            d = d * 2 - 9 if d > 4 else d * 2
        total += d
    return total % 10 == 0


def is_iccid(s):
    """SIM 卡的 ICCID：19～20 位，8986 开头（中国）。"""
    return len(s) in (19, 20) and s.isascii() and s.isalnum() and s.startswith("8986")


def is_bank_card(s):
    """16～19 位卡号，Luhn 校验；身份证号、ICCID 不算。"""
    return 16 <= len(s) <= 19 and _digits(s) and not is_id_card(s) and not is_iccid(s) and _luhn(s)


def _bank_loose(s):
    """列名对上「银行卡 / 账号」时：12～19 位数字就算（对公账号不一定过 Luhn）。"""
    return 12 <= len(s) <= 19 and _digits(s) and not is_id_card(s)


def is_imsi(s):
    return len(s) == 15 and _digits(s) and s.startswith("460")


def is_imei(s):
    return len(s) == 15 and _digits(s) and not s.startswith("460") and _luhn(s)


def is_landline(s):
    return bool(_LANDLINE.match(s)) and not s.startswith("00")


def is_email(s):
    return bool(_EMAIL.match(s))


def is_name(s):
    """2～4 个汉字的人名；带「·」的少数民族名放宽到 15 个字。"""
    if "·" in s or "•" in s:
        core = s.replace("·", "").replace("•", "")
        return 2 <= len(core) <= 15 and all(_CJK.match(ch) for ch in core)
    return 2 <= len(s) <= 4 and all(_CJK.match(ch) for ch in s)


def is_address(s):
    return len(s) >= 6 and len(_CJK.findall(s)) >= 4


def _is_dateish(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return True
    return bool(_DATEISH.match(_text(v)))


def _is_coord(v):
    v = key_value(v)
    if isinstance(v, str):
        try:
            v = float(v.strip())
        except ValueError:
            return False
    return isinstance(v, (int, float)) and not isinstance(v, bool) and -180 <= v <= 180 and v != int(v)


def _on_norm(f):
    return lambda v: f(_norm(v))


def _on_text(f):
    return lambda v: f(_text(v))


# (类型, 列名关键词, 强校验：单看值就能认, 宽校验：列名对上后用, 宽校验最低占比)
# 顺序有讲究：先认校验最严的号码，再认靠列名的；ICCID 在银行卡前，手机号在固话前。
_KINDS = (
    ("身份证号", ("身份证", "证件号", "身份号码", "公民身份"), _on_norm(is_id_card),
     _on_norm(lambda s: is_id_card(s) or _is_id15(s)), 0.3),
    ("ICCID", ("iccid", "sim卡号", "卡序列号"), _on_norm(is_iccid), _on_norm(is_iccid), 0.3),
    ("银行卡号", ("银行卡", "银行账号", "银行帐号", "托收账号", "托收帐号", "卡号"), _on_norm(is_bank_card),
     _on_norm(_bank_loose), 0.3),
    ("手机号", ("手机", "移动电话", "联系电话", "服务号码", "计费号码", "用户号码", "电话", "号码", "联系方式"),
     _on_norm(is_mobile), _on_norm(is_mobile), 0.3),
    ("IMSI", ("imsi",), None, _on_norm(is_imsi), 0.3),
    ("IMEI", ("imei", "串号"), None, _on_norm(is_imei), 0.3),
    ("固定电话", ("固话", "固定电话", "座机", "办公电话", "电话"), None, _on_norm(is_landline), 0.3),
    ("邮箱", ("邮箱", "邮件", "email", "e-mail", "mail"), _on_text(is_email), _on_text(is_email), 0.3),
    ("姓名", ("姓名", "名字", "户名", "联系人", "机主", "收件人", "收货人", "客户名称", "用户名称"), None,
     _on_text(is_name), 0.5),
    ("地址", ("地址", "住址", "居住地", "户籍"), None, _on_text(is_address), 0.5),
    ("出生日期", ("出生日期", "出生年月", "生日"), None, _is_dateish, 0.5),
    ("经纬度", ("经度", "纬度", "经纬度", "坐标"), None, _is_coord, 0.5),
    ("账号编码", ("宽带账号", "宽带帐号", "客户编码", "客户编号", "用户编码", "用户编号", "账户编码", "账户编号",
              "帐户编码", "合同号", "用户id", "客户id"), None, None, 0),
)
# 列名里带这些字的，是在说「关于这个信息的统计 / 属性」，不是信息本身（手机号码归属地、证件类型…）
_NOT_HEAD = ("是否", "类型", "状态", "归属", "数量", "个数", "次数", "人数", "率", "等级", "有效期", "验证")


def detect_column(name, values):
    """一列是什么个人信息：name 列名，values 样本值。认不出返回 None。只在本机看值，结果里不带值。"""
    low = soft_clean(name).lower()
    head_ok = bool(low) and not any(w in low for w in _NOT_HEAD)
    vals = [v for v in values if _filled(v)]
    for kind, keys, strong, loose, need in _KINDS:
        if strong and vals and sum(1 for v in vals if strong(v)) / len(vals) >= 0.6:
            return kind
        if head_ok and any(k in low for k in keys):
            if loose is None:
                return kind
            if vals and sum(1 for v in vals if loose(v)) / len(vals) >= need:
                return kind
    return None


# =====================================================
# 处理：打码 / 代号 / 泛化 / 清空
# =====================================================

_COMPOUND = ("欧阳", "司马", "上官", "诸葛", "东方", "皇甫", "尉迟", "公孙", "慕容", "长孙", "宇文", "司徒",
             "夏侯", "轩辕", "令狐", "钟离", "端木", "百里", "呼延", "南宫", "独孤", "西门", "万俟", "闻人")
_TOWN = re.compile(r"^.{2,40}?(?:街道|镇|乡)")
_COUNTY = re.compile(r"^.{2,30}?(?:区|县|旗)")
_CITY = re.compile(r"^.{2,20}?市")
_PREFIX = {"手机号": "TEL", "固定电话": "TEL", "身份证号": "IDN", "银行卡号": "BNK", "姓名": "NAM", "地址": "ADR",
           "邮箱": "EML", "ICCID": "SIM", "IMSI": "SIM", "IMEI": "DEV", "账号编码": "ACC"}


def _keep_ends(s, head, tail):
    """留前 head 位、后 tail 位，中间打星；太短的只留后半截。"""
    n = len(s)
    if n == 0:
        return s
    if n <= head + tail:
        t = min(tail, n // 2)
        return "*" * (n - t) + (s[-t:] if t else "")
    return s[:head] + "*" * (n - head - tail) + (s[-tail:] if tail else "")


def _addr_cut(s, town):
    """地址留到镇街（town=True）或区县，后面的门牌不要。认不出行政区划返回 None。"""
    for rx in ((_TOWN, _COUNTY, _CITY) if town else (_COUNTY, _CITY)):
        m = rx.match(s)
        if m:
            return m.group(0)
    return None


def _mask_name(s):
    if len(s) <= 1:
        return "*"
    k = 2 if s[:2] in _COMPOUND and len(s) > 2 else 1
    return s[:k] + "".join(ch if ch in "·•" else "*" for ch in s[k:])


def _mask(kind, v):
    t = _text(v)
    n = _norm(v)
    if kind == "手机号":
        s = _strip86(n)
        return s[:3] + "****" + s[-4:] if len(s) == 11 else _keep_ends(s, 3, 4)
    if kind == "身份证号":
        return _keep_ends(n, 6, 4)
    if kind in ("银行卡号", "ICCID"):
        return _keep_ends(n, 6, 4)
    if kind == "IMSI":
        return _keep_ends(n, 5, 4)
    if kind == "固定电话":
        if n.startswith("0") and len(n) > 7:
            area = 3 if n[1] in "12" else 4
            return n[:area] + "-" + "*" * (len(n) - area - 4) + n[-4:]
        return _keep_ends(n, 0, 4)
    if kind == "邮箱":
        local, at, dom = t.partition("@")
        return (local[:1] + "***" + at + dom) if at else _keep_ends(t, 1, 0)
    if kind == "姓名":
        return _mask_name(t)
    if kind == "地址":
        cut = _addr_cut(t, town=True)
        return cut + "****" if cut and cut != t else (t[:6] + "****" if len(t) > 6 else "*" * len(t))
    if kind in ("出生日期", "经纬度"):
        return _generalize(kind, v)
    if kind in ("IMEI", "账号编码"):
        return _keep_ends(t, 0, 4)
    # 自定义列：数字留后 4 位，文字留第一个字
    return _keep_ends(t, 0, 4) if _digits(n) else (t[:1] + "*" * (len(t) - 1) if len(t) > 1 else "*")


def _year(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return f"{v.year}年"
    m = re.search(r"(1[89]|20)\d{2}", _text(v))
    return f"{m.group(0)}年" if m else None


def _generalize(kind, v):
    """泛化：地址到区县、出生日期到年、身份证号只留地区和出生年、坐标留两位小数；
    自定义列的数字按最高位分段（5234 → 5000-6000）、日期到年。其余类型泛化不了，退回打码。"""
    if kind == "地址":
        cut = _addr_cut(_text(v), town=False)
        return cut if cut else _mask(kind, v)
    if kind == "出生日期":
        return _year(v) or "*"
    if kind == "身份证号":
        n = _norm(v)
        return f"{n[:6]}·{n[6:10]}年" if len(n) == 18 else _mask(kind, v)
    if kind not in ("经纬度", "自定义"):
        return _mask(kind, v)
    if isinstance(v, (datetime.date, datetime.datetime)):
        return _year(v)
    x = key_value(v)
    if isinstance(x, str):
        try:
            x = float(x.replace(",", "").strip())
        except ValueError:
            return _mask(kind, v)
    if not isinstance(x, (int, float)) or isinstance(x, bool):
        return _mask(kind, v)
    if kind == "经纬度":
        return round(float(x), 2)
    w = 10 ** max(0, len(str(int(abs(x)))) - 1)
    lo = int(x // w * w)
    return f"{lo}-{lo + w}"


def _code(key, kind, v):
    n = _strip86(_norm(v)) if kind in ("手机号", "身份证号", "银行卡号", "ICCID", "IMSI", "IMEI", "固定电话") \
        else _text(v).upper()
    d = hmac.new(key, f"{kind}|{n}".encode(), hashlib.sha256).digest()
    return f"{_PREFIX.get(kind, 'X')}-{base64.b32encode(d).decode()[:10]}"


def apply_value(v, kind, method, key=b""):
    """按处理方式给一个值脱敏，返回新值（clear 返回 None）。空的、「无」之类的占位原样返回。"""
    if method == "keep" or not _filled(v):
        return v
    if method == "clear":
        return None
    if method == "pseudo":
        return _code(key, kind, v)
    if method == "generalize":
        return _generalize(kind, v)
    return _mask(kind, v)


_RX_ID = re.compile(r"(?<!\d)\d{17}[\dXx](?!\d)")
_RX_CARD = re.compile(r"(?<!\d)\d{16,19}(?!\d)")
# 前面紧挨着 +86 / 86 的也算（+8613800001234）
_RX_MOBILE = re.compile(r"(?:(?<=86)|(?<!\d))(1[3-9]\d)([- ]?)(\d{4})([- ]?)(\d{4})(?!\d)")


def mask_text(s, kinds=EMBED_KINDS):
    """一段文字里夹带的身份证号、银行卡号、手机号就地打码（都带校验，不乱遮）。返回 (新文字, 遮了几处)。"""
    n = 0

    def _id(m):
        nonlocal n
        if is_id_card(m.group(0)):
            n += 1
            return _keep_ends(m.group(0), 6, 4)
        return m.group(0)

    def _card(m):
        nonlocal n
        if is_bank_card(m.group(0)):
            n += 1
            return _keep_ends(m.group(0), 6, 4)
        return m.group(0)

    def _mob(m):
        nonlocal n
        n += 1
        return m.group(1) + m.group(2) + "****" + m.group(4) + m.group(5)

    if "身份证号" in kinds:
        s = _RX_ID.sub(_id, s)
    if "银行卡号" in kinds:
        s = _RX_CARD.sub(_card, s)
    if "手机号" in kinds:
        s = _RX_MOBILE.sub(_mob, s)
    return s, n


def find_kinds(v, kinds=EMBED_KINDS):
    """复检用：这个值里还有没有完整的号码，有哪几类。"""
    v = key_value(v)
    if v is None or isinstance(v, (bool, datetime.date, datetime.time)):
        return set()
    if isinstance(v, (int, float)):
        s = str(int(v)) if v == int(v) else ""
        return {"手机号"} if "手机号" in kinds and is_mobile(s) else set()
    s = str(v)
    found = set()
    if "身份证号" in kinds and any(is_id_card(m.group(0)) for m in _RX_ID.finditer(s)):
        found.add("身份证号")
    if "银行卡号" in kinds and any(is_bank_card(m.group(0)) for m in _RX_CARD.finditer(s)):
        found.add("银行卡号")
    if "手机号" in kinds and _RX_MOBILE.search(s):
        found.add("手机号")
    return found


# =====================================================
# 一次脱敏任务
# =====================================================

def _scene_method(scene, kind):
    if scene == "内部下发":
        return "mask" if kind in ("身份证号", "银行卡号") else "keep"
    if scene == "发群通报":
        return "mask"
    if scene == "交给AI分析":
        return "generalize" if kind in ("地址", "出生日期", "经纬度") else "pseudo"
    if scene == "对外报送":
        return "pseudo" if kind == "账号编码" else "clear"
    return "keep"            # 自定义：没指定的列照发


def _fake(kind):
    """报告里的示例用的虚构值，不来自任何表格。"""
    if kind == "身份证号":
        s17 = "11010519900307002"
        return s17 + _id_check_digit(s17)
    return {"手机号": "13800001234", "银行卡号": "6222000000001234", "姓名": "张三", "邮箱": "zhangsan@example.com",
            "地址": "某某省某某市某某区某某街道某某路 1 号", "固定电话": "07608888123", "ICCID": "89860000000000001234",
            "IMSI": "460000000001234", "IMEI": "860000000001234", "出生日期": datetime.date(1990, 3, 7),
            "经纬度": 113.392411, "账号编码": "GD0000001234"}.get(kind, "示例 12345678")


class Job:
    """一次脱敏：同一批里各文件、各 sheet 共用一份规则、一份代号种子，统计攒在一起写报告。"""

    def __init__(self, scene, columns=None, notes=True, seed=""):
        self.scene = scene
        self.columns = {soft_clean(k): m for k, m in (columns or {}).items() if m in METHODS}
        self.notes = notes
        try:
            self.key = bytes.fromhex(seed) if seed else b""
        except ValueError:
            self.key = b""
        self.fixed_seed = bool(self.key)
        if not self.key:
            self.key = os.urandom(16)        # 没给种子：代号只在这一批里一致
        if scene == "内部下发":
            self.active = {"身份证号", "银行卡号"}
        else:
            self.active = set(EMBED_KINDS)
        self.kept = set()          # 用户明确要照发的、本来要处理的列（复检时跳过）
        self.details = []          # [{file, sheet, column, kind, method, cells}]
        self.embedded = 0          # 文字里夹带、就地打码的处数
        self.hits = []             # 复检残留
        self.files = []            # [(相对路径, 绝对路径)]：写指纹用

    @classmethod
    def from_config(cls, d):
        """config['desens'] → Job；没开返回 None。d 可以是场景名，也可以是
        {"scene": 场景, "columns": {列名: mask|pseudo|generalize|clear|keep}, "notes": True, "seed": 十六进制}。"""
        if not d:
            return None
        if isinstance(d, str):
            d = {"scene": d}
        if not isinstance(d, dict):
            return None
        scene = d.get("scene") or ""
        cols = d.get("columns") or {}
        if scene not in SCENES:
            if not cols:
                return None
            scene = "自定义"
        return cls(scene, cols, d.get("notes", True) is not False, str(d.get("seed") or ""))

    # ---------- 分段做：存进度 / 读回 ----------
    def to_state(self):
        """存成能写进 JSON 的样子（技能分几条命令做完一批时用）。含代号种子：做完后调用方应把它去掉。"""
        return {"scene": self.scene, "columns": self.columns, "notes": self.notes, "key": self.key.hex(),
                "fixed": self.fixed_seed, "kept": sorted(self.kept), "details": self.details,
                "embedded": self.embedded, "hits": self.hits, "files": [list(f) for f in self.files]}

    @classmethod
    def from_state(cls, d):
        job = cls(d.get("scene") or "自定义", d.get("columns"), d.get("notes", True), d.get("key") or "")
        job.fixed_seed = bool(d.get("fixed"))
        job.kept = set(d.get("kept") or [])
        job.details = list(d.get("details") or [])
        job.embedded = int(d.get("embedded") or 0)
        job.hits = list(d.get("hits") or [])
        job.files = [tuple(f) for f in d.get("files") or []]
        return job

    # ---------- 识别 ----------
    def rules(self, names, samples):
        """names：列名（0 基）；samples：{0 基列号: [样本值]}。返回 {1 基列号: (类型, 处理方式)}。"""
        out = {}
        width = max([len(names)] + [c + 1 for c in samples])
        for c in range(width):
            name = soft_clean(names[c]) if c < len(names) else ""
            forced = self.columns.get(name) if name else None
            kind = detect_column(name, samples.get(c, []))
            if kind is None and forced is None:
                continue
            kind = kind or "自定义"
            method = forced or _scene_method(self.scene, kind)
            if method == "keep" and kind in self.active and name:
                self.kept.add(name)          # 有意照发的列，复检时跳过
            out[c + 1] = (kind, method)      # 照发的也记一笔，报告里写清楚
        return out

    # ---------- 处理一个 sheet ----------
    def sheet(self, ws, h, last, names, where=("", ""), log_fn=None):
        """在内存里给 ws 第 h+1 ~ last 行脱敏（h=0 表示没有表头）。表头和上面的标题行不动。返回处理的格数。"""
        if last <= h:
            return 0
        samples = {}
        for row in ws.iter_rows(min_row=h + 1, max_row=min(last, h + SAMPLE_ROWS), values_only=True):
            for c, v in enumerate(row):
                if v is not None:
                    samples.setdefault(c, []).append(v)
        rules = self.rules(names, samples)
        counts = {}
        emb = 0
        for (r, c), cell in list(ws._cells.items()):
            if r <= h or r > last or isinstance(cell, MergedCell) or cell.value is None:
                continue
            rule = rules.get(c)
            if rule:
                kind, method = rule
                if method == "keep" or not _filled(cell.value):
                    continue
                put(cell, apply_value(cell.value, kind, method, self.key))
                counts[c] = counts.get(c, 0) + 1
            elif self.notes:
                v = cell.value
                if isinstance(v, str):
                    nv, k = mask_text(v, self.active)
                    if k:
                        put(cell, nv)
                        emb += k
                elif "手机号" in self.active and find_kinds(v, {"手机号"}):
                    put(cell, _mask("手机号", v))
                    emb += 1
        self.embedded += emb
        done = []
        for c, (kind, method) in sorted(rules.items()):
            name = soft_clean(names[c - 1]) if c - 1 < len(names) else ""
            self.details.append({"file": where[0], "sheet": where[1], "column": name or get_column_letter(c),
                                 "kind": kind, "method": method, "cells": counts.get(c, 0)})
            if method != "keep":
                done.append(f"「{name or get_column_letter(c)}」{kind}{METHODS[method]} {counts.get(c, 0)} 格")
        if log_fn and (done or emb):
            log_fn(f"  🛡 sheet「{where[1]}」脱敏：" + "；".join(done)
                   + (f"{'；' if done else ''}文字里夹带的号码打码 {emb} 处" if emb else ""))
        return sum(counts.values()) + emb

    # ---------- 复检 ----------
    def _scan_book(self, wb, label):
        hits = {}
        for ws in wb.worksheets:
            if hasattr(ws, "reset_dimensions"):
                ws.reset_dimensions()
            rows = list(ws.iter_rows(values_only=True))
            head = [tuple(key_value(v) for v in r) for r in rows[:HEAD_SCAN]]
            h = resolve_header_row(head, {}) if head else -1
            names = col_names(head, h) if h > 0 else []
            skip = {c for c, n in enumerate(names) if n in self.kept}
            for r in rows:
                for c, v in enumerate(r):
                    if v is None or c in skip:
                        continue
                    for kind in find_kinds(v, self.active):
                        col = names[c] if c < len(names) and names[c] else get_column_letter(c + 1)
                        k = (label, ws.title, col, kind)
                        hits[k] = hits.get(k, 0) + 1
        return [{"file": f, "sheet": s, "column": col, "kind": k, "cells": n} for (f, s, col, k), n in hits.items()]

    def scan_files(self, items):
        """items：[(相对路径, 绝对路径)]。落盘后重新只读扫一遍，返回残留 [{file, sheet, column, kind, cells}]（不含值）。"""
        found = []
        for rel, path in items:
            try:
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            except Exception:
                continue
            try:
                found += self._scan_book(wb, rel)
            finally:
                wb.close()
        self.hits += found
        return found

    def scan_bytes(self, data, label):
        """还没落盘的工作簿（字节）先扫一遍：不过就不写出去。"""
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        try:
            found = self._scan_book(wb, label)
        finally:
            wb.close()
        self.hits += found
        return found

    # ---------- 报告 ----------
    def summary(self):
        done = [d for d in self.details if d["method"] != "keep"]
        return {"scene": self.scene, "columns": len(done), "cells": sum(d["cells"] for d in done),
                "embedded": self.embedded, "residual": sum(h["cells"] for h in self.hits), "ok": not self.hits}

    def write_report(self, path, files=None):
        """脱敏报告.xlsx：概要 / 处理明细 / 复检 / 指纹。只有列名、类型、格数和虚构示例。"""
        files = self.files if files is None else files
        s = self.summary()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "概要"
        rows = [("用途", f"{self.scene}：{SCENE_TEXT.get(self.scene, '')}"),
                ("处理时间", time.strftime("%Y-%m-%d %H:%M")),
                ("处理了", f"{s['columns']} 列、{s['cells']} 个格子；文字里夹带的号码 {s['embedded']} 处"),
                ("复检", "通过：处理过的类型里没有发现完整的号码" if s["ok"] else
                 f"没通过：还有 {s['residual']} 个格子有完整的号码，这批结果不要发出去（见「复检」页）"),
                ("代号", "固定种子：同一份种子下，各批的代号一致" if self.fixed_seed else
                 "本次随机：代号只在这一批里一致，种子没有保存，推不回原值"),
                ("说明", "识别靠列名和号码校验，不保证找全；表头上方的标题、sheet 名、文件名、批注里的内容不处理。"
                       "处理方式参照 GB/T 37964-2019（打码＝部分屏蔽、代号＝假名化、泛化、清空＝抑制）。"
                       "本报告不含任何表格里的真实值，示例都是虚构的。")]
        for r in rows:
            ws.append(list(r))
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 100
        det = wb.create_sheet("处理明细")
        det.append(["文件", "sheet", "列", "识别为", "处理方式", "处理的格数", "示例（虚构）"])
        for d in self.details:
            fake = _fake(d["kind"])
            after = apply_value(fake, d["kind"], d["method"], self.key)
            det.append([d["file"], d["sheet"], d["column"], d["kind"], METHODS[d["method"]], d["cells"],
                        f"{fake} → {'（空）' if after is None else after}"])
        for col, w in zip("ABCDEFG", (30, 14, 18, 10, 10, 10, 46)):
            det.column_dimensions[col].width = w
        chk = wb.create_sheet("复检")
        chk.append(["文件", "sheet", "列", "还有完整的", "格数"])
        if self.hits:
            for h in self.hits:
                chk.append([h["file"], h["sheet"], h["column"], h["kind"], h["cells"]])
        else:
            chk.append(["（全部通过）", "", "", "、".join(sorted(self.active)), 0])
        chk.column_dimensions["A"].width = 30
        fp = wb.create_sheet("指纹")
        fp.append(["文件", "SHA-256"])
        for rel, p in files:
            try:
                fp.append([rel, _sha256(p)])
            except OSError:
                pass
        fp.column_dimensions["A"].width = 40
        fp.column_dimensions["B"].width = 70
        wb.save(path)
        return path


def _sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def write_fail_note(folder, hits):
    """复检没过：在结果文件夹里写一张醒目的说明（不删任何文件）。"""
    p = os.path.join(folder, FAIL_NOTE)
    lines = ["这批结果没通过脱敏复检，请不要发出去。", "", "下面这些格子里还有完整的号码（只列位置，不列内容）："]
    lines += [f"  {h['file']} · {h['sheet']} · 「{h['column']}」：{h['kind']} {h['cells']} 格" for h in hits]
    lines += ["", "处理办法：在源表里把这些列的列名写清楚（如「身份证号」「手机号」），或者在脱敏设置里指定这些列，再跑一次。"]
    with open(p, "w", encoding="utf-8-sig") as f:
        f.write("\n".join(lines) + "\n")
    return p


# =====================================================
# 单独脱敏：一个表格或一个文件夹
# =====================================================

def _last_row(ws):
    last = 0
    for (r, _c), cell in ws._cells.items():
        if r > last and cell.value is not None and cell.value != "":
            last = r
    return last


def _clear_cols(ws, start, cols):
    """隐藏列从 start 行起清空（跨到可见列的合并格不动）。与拆分的 _clear_hidden 同一规则。"""
    keep = {(mr.min_row, mr.min_col) for mr in ws.merged_cells.ranges
            if any(c not in cols for c in range(mr.min_col, mr.max_col + 1))}
    for (r, c), cell in list(ws._cells.items()):
        if r >= start and c in cols and not isinstance(cell, MergedCell) and (r, c) not in keep:
            cell.value = None


def desens_book(job, src, label, config=None, log_fn=None):
    """给一个表格脱敏，返回 (新工作簿的字节, 处理了几个 sheet)。源文件不动。

    隐藏 sheet 不留、隐藏列清空（keep_hidden 打开时照留）；批注、图片、图表、透视表缓存、文档属性都去掉——
    这些地方也可能带着原始数据，又不在单元格里，脱不了就不带。
    """
    config = config or {}
    keep_hidden = config.get("keep_hidden", False)
    wb, _ = open_book(src)
    hid = hidden_info(src)
    n = 0
    for ws in list(wb.worksheets):
        state, hcols = hid.get(ws.title, ("visible", set()))
        if state != "visible" and not keep_hidden:
            if len(wb.worksheets) > 1:
                wb.remove(ws)
                if log_fn:
                    log_fn(f"  🙈 sheet「{ws.title}」是隐藏的，没有带进结果")
                continue
        last = _last_row(ws)
        head = [tuple(key_value(v) for v in r) for r in ws.iter_rows(min_row=1, max_row=min(last, HEAD_SCAN),
                                                                     values_only=True)] if last else []
        h = resolve_header_row(head, config) if head else -1
        h = h if h > 0 else 0
        names = col_names(head, h) if h else []
        if hcols and not keep_hidden:
            _clear_cols(ws, max(h, 1), {c for c in hcols if c <= ws.max_column})
        job.sheet(ws, h, last, names, (label, ws.title), log_fn)
        for cell in ws._cells.values():
            if cell.comment is not None:
                cell.comment = None
        ws._images = []
        ws._charts = []
        ws._pivots = []
        n += 1
    wb.properties = DocumentProperties()
    _no_formulas(wb)
    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    return buf.getvalue(), n


def desens_copy(path, setting, log_fn=None):
    """汇总后另存一份脱敏版（默认「对外报送」用途）：{原名}_脱敏版.xlsx + {原名}_脱敏报告.xlsx，原汇总表不动。

    先在内存里处理、复检，过了才写出去。返回 (脱敏版路径，没写出来是 "", Job)。
    """
    log = log_fn or (lambda m: None)
    job = Job.from_config(setting or "对外报送") or Job("对外报送")
    stem = os.path.splitext(path)[0]
    name = os.path.basename(stem) + "_脱敏版.xlsx"
    data, _n = desens_book(job, path, name, None, log)
    folder = os.path.dirname(path)
    if job.scan_bytes(data, name):
        write_fail_note(folder, job.hits)
        log(f"  ❌ 脱敏版没通过复检，没有写出来；见「{FAIL_NOTE}」")
        out = ""
    else:
        out = stem + "_脱敏版.xlsx"
        with open(out, "wb") as f:
            f.write(data)
        job.files = [(name, out)]
        s = job.summary()
        log(f"🛡 脱敏版（{job.scene}）：{out}（处理 {s['columns']} 列、{s['cells']} 格）")
    try:
        job.write_report(stem + "_脱敏报告.xlsx")
    except Exception as e:
        log(f"  ⚠️ 脱敏报告没写成：{e}")
    return out, job


def inspect(path, setting=None, config=None):
    """只读识别一个表格：每个 sheet 认出哪些列、按用途怎么处理。界面和技能在开始前给用户确认用。

    返回 [{sheet, hidden, header_row, rows, cols, columns: [{column, kind, method}]}]——只有列名、类型和计数，
    没有任何格子的值。隐藏的 sheet 也列出来（hidden 为 True），结果里默认不带它。
    """
    config = config or {}
    job = Job.from_config(setting or "发群通报") or Job("发群通报")
    hid = hidden_info(path)
    out = []
    for title, rows in _iter_sheets_ro(path):
        head, samples, n, width = [], {}, 0, 0
        for i, row in enumerate(rows):
            if i < HEAD_SCAN:
                head.append(tuple(key_value(v) for v in row))
            if any(v is not None and v != "" for v in row):
                n = i + 1
                width = max(width, len(row))
            if i < HEAD_SCAN + SAMPLE_ROWS:
                for c, v in enumerate(row):
                    if v is not None:
                        samples.setdefault(i, {})[c] = v
        h = resolve_header_row(head, config) if head else -1
        h = h if h > 0 else 0
        names = col_names(head, h) if h else []
        by_col = {}
        for i, cells in samples.items():
            if h < i + 1 <= h + SAMPLE_ROWS:
                for c, v in cells.items():
                    by_col.setdefault(c, []).append(v)
        cols = []
        for c, (kind, method) in sorted(job.rules(names, by_col).items()):
            name = soft_clean(names[c - 1]) if c - 1 < len(names) else ""
            cols.append({"column": name or get_column_letter(c), "kind": kind, "method": method})
        state = hid.get(title, ("visible", set()))[0]
        out.append({"sheet": title, "hidden": state != "visible", "header_row": h, "rows": max(0, n - h),
                    "cols": width, "columns": cols})
    return out


def _out_dir(base):
    stamp = datetime.datetime.now().strftime("%m%d%H%M")
    p = os.path.join(base, RESULT_DIR, f"{stamp}结果")
    k = 2
    while os.path.exists(p):
        p = os.path.join(base, RESULT_DIR, f"{stamp}结果_{k}")
        k += 1
    os.makedirs(p, exist_ok=True)
    return p


def _tasks(input_path):
    if os.path.isfile(input_path):
        return [(input_path, os.path.basename(input_path))]
    tasks = []
    skip = os.path.abspath(os.path.join(input_path, RESULT_DIR))
    for root, dirs, files in os.walk(input_path):
        if os.path.abspath(root).startswith(skip):
            continue
        dirs.sort()
        for f in sorted(files):
            lf = f.lower()
            if lf.endswith((".xlsx", ".xlsm", ".xls")) and not f.startswith("~$"):
                p = os.path.join(root, f)
                tasks.append((p, os.path.relpath(p, input_path)))
    return tasks


def run_desens(config, log_fn=None, progress_fn=None, stop_flag=None):
    """单独脱敏入口，签名同 run_split。config：input_path（文件或文件夹）、desens（场景或设置，默认「发群通报」）、
    可选 keep_hidden / header_mode / header_row。结果在输入旁「脱敏结果\\{月日时分}结果」，文件名加「_脱敏」。

    每个文件先在内存里处理、复检，过了才写出去；没过的不写，并写「未通过复检_请勿发出.txt」。返回结果文件夹。
    """
    log = log_fn or (lambda m: None)
    job = Job.from_config(config.get("desens") or "发群通报") or Job("发群通报")
    input_path = config["input_path"]
    base = os.path.dirname(input_path) if os.path.isfile(input_path) else input_path
    out = _out_dir(base)
    tasks = _tasks(input_path)
    log(f"共找到 {len(tasks)} 个表格，用途：{job.scene}（{SCENE_TEXT.get(job.scene, '')}）")
    written, failed, blocked, stopped = [], 0, 0, False
    for i, (fp, rel) in enumerate(tasks):
        if stop_flag and stop_flag():
            stopped = True
            log("⛔ 已停止")
            break
        log(f"[{i + 1}/{len(tasks)}] {rel}")
        stem, _ext = os.path.splitext(rel)
        rel_out = stem + "_脱敏.xlsx"
        try:
            data, _n = desens_book(job, fp, rel_out, config, log)
            hits = job.scan_bytes(data, rel_out)
            if hits:
                blocked += 1
                log(f"  ❌ 复检没过：还有 {sum(h['cells'] for h in hits)} 个格子有完整的号码，这个文件没有写出来")
            else:
                dst = os.path.join(out, *[safe_filename(p) for p in os.path.normpath(rel_out).split(os.sep)])
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                with open(dst, "wb") as f:
                    f.write(data)
                written.append((rel_out, dst))
        except Exception as e:
            failed += 1
            log(f"  ❌ 处理失败：{e}")
        if progress_fn:
            progress_fn((i + 1) / max(1, len(tasks)) * 0.95)
    job.files = written
    report = os.path.join(out, REPORT_NAME)
    try:
        job.write_report(report)
    except Exception as e:
        report = ""
        log(f"  ⚠️ 脱敏报告没写成：{e}")
    if job.hits:
        write_fail_note(out, job.hits)
        log(f"⚠️ {blocked} 个文件没通过复检，没有写出来；见「{FAIL_NOTE}」")
    s = job.summary()
    log(f"\n✅ 完成：{len(written)} 个文件，处理 {s['columns']} 列、{s['cells']} 格，文字里夹带的号码 {s['embedded']} 处")
    log(f"📁 结果：{out}")
    if progress_fn:
        progress_fn(1.0)
    summary = dict(s, mode="desens", files=len(written), blocked=blocked, failed_files=failed, stopped=stopped,
                   output=out, report=report)
    log(f"[SUMMARY] {json.dumps(summary, ensure_ascii=False)}")
    return out
