"""短密钥的本机封存：Windows 自带的数据保护接口（DPAPI，系统接口 crypt32.dll）。

封好的字符串只有同一台电脑上的同一个 Windows 账户解得开：进度文件被拷走、被别的账户读到，都还原不出原来的字节。
技能分几条命令做完一批脱敏时，用它把代号种子存进进度文件（core.desens.Job.to_state / from_state）。
不是 Windows、或系统接口调用失败时，available() 为 False，调用方就不存种子（这一批只能在一次调用里做完）。

Copyright (c) 2026 AbeLin · MIT License
"""
import base64
import ctypes
import sys

_UI_FORBIDDEN = 0x1      # CRYPTPROTECT_UI_FORBIDDEN：不弹任何对话框


class _Blob(ctypes.Structure):
    _fields_ = [("cbData", ctypes.c_uint32), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _call(name, data):
    if sys.platform != "win32":
        raise OSError("只在 Windows 上可用")
    crypt32 = ctypes.WinDLL("crypt32.dll", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32.dll")
    fn = getattr(crypt32, name)
    fn.argtypes = [ctypes.POINTER(_Blob), ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p,
                   ctypes.c_uint32, ctypes.POINTER(_Blob)]
    fn.restype = ctypes.c_int
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    buf = ctypes.create_string_buffer(data, len(data))
    src = _Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    out = _Blob()
    if not fn(ctypes.byref(src), None, None, None, None, _UI_FORBIDDEN, ctypes.byref(out)):
        raise OSError(f"{name} 失败（错误码 {ctypes.get_last_error()}）")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(out.pbData, ctypes.c_void_p))


def seal(data):
    """字节 → 封好的字符串（base64），只有当前 Windows 账户解得开。"""
    return base64.b64encode(_call("CryptProtectData", bytes(data))).decode("ascii")


def unseal(text):
    """seal 的反过程；换了电脑、换了账户或内容被改过时抛 ValueError。"""
    try:
        return _call("CryptUnprotectData", base64.b64decode(text.encode("ascii"), validate=True))
    except (OSError, ValueError, TypeError, AttributeError) as e:
        raise ValueError("解不开：不是这台电脑上这个 Windows 账户封的，或者内容被改过") from e


def available():
    """这台电脑能不能用：试封、试解一次。"""
    try:
        return unseal(seal(b"ok")) == b"ok"
    except (OSError, ValueError):
        return False
