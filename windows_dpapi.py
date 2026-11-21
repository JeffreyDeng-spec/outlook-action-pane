"""Windows account-bound DPAPI store primitive; tested with dummy values only."""
from __future__ import annotations
import ctypes as C
from ctypes import wintypes as W
import os

class _Blob(C.Structure):
    _fields_ = (("cbData", W.DWORD), ("pbData", C.POINTER(C.c_ubyte)))

def _call(data: bytes, encrypt: bool) -> bytes:
    if os.name != "nt":
        raise OSError("DPAPI requires Windows")
    if not isinstance(data, bytes) or not data:
        raise ValueError("Non-empty bytes are required")
    crypt = C.WinDLL("crypt32.dll", use_last_error=True)
    kernel = C.WinDLL("kernel32.dll", use_last_error=True)
    kernel.LocalFree.argtypes = (C.c_void_p,)
    kernel.LocalFree.restype = C.c_void_p
    buffer = (C.c_ubyte * len(data)).from_buffer_copy(data)
    src = _Blob(len(data), buffer)
    out = _Blob()
    try:
        method = crypt.CryptProtectData if encrypt else crypt.CryptUnprotectData
        method.argtypes = (C.POINTER(_Blob), C.c_void_p, C.c_void_p, C.c_void_p, C.c_void_p, W.DWORD, C.POINTER(_Blob))
        method.restype = W.BOOL
        if not method(C.byref(src), None, None, None, None, 1, C.byref(out)):
            raise C.WinError(C.get_last_error())
        return C.string_at(out.pbData, out.cbData)
    finally:
        C.memset(buffer, 0, len(data))
        if out.pbData:
            C.memset(out.pbData, 0, out.cbData)
            kernel.LocalFree(C.cast(out.pbData, C.c_void_p))

def protect(data: bytes) -> bytes:
    return _call(data, True)

def unprotect(encrypted: bytes) -> bytes:
    return _call(encrypted, False)
