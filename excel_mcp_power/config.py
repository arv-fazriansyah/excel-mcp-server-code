"""Configuration from environment variables."""
from __future__ import annotations

import os
import platform

BACKEND = os.environ.get("EXCEL_MCP_BACKEND", "auto").lower()
USE_COM = BACKEND == "com" or (BACKEND == "auto" and platform.system() == "Windows")

ALLOW_MACROS = os.environ.get("EXCEL_MCP_ALLOW_MACROS", "0") == "1"
READONLY = os.environ.get("EXCEL_MCP_READONLY", "0") == "1"
ALLOW_VBA_WRITE = os.environ.get("EXCEL_MCP_ALLOW_VBA_WRITE", "0") == "1"
ALLOW_UNSAFE_FORMULAS = os.environ.get("EXCEL_MCP_ALLOW_UNSAFE_FORMULAS", "0") == "1"
VISIBLE = os.environ.get("EXCEL_MCP_VISIBLE", "0") == "1"
MAX_CELLS = int(os.environ.get("EXCEL_MCP_MAX_CELLS", "20000"))

def _documents_dir() -> str:
    if platform.system() == "Windows":
        try:
            import ctypes
            buf = ctypes.create_unicode_buffer(260)
            if ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf) == 0 and buf.value:
                return buf.value
        except Exception:
            pass
    return os.path.join(os.path.expanduser("~"), "Documents")

ROOT = os.environ.get("EXCEL_MCP_ROOT") or os.path.join(_documents_dir(), "excel_files")
os.makedirs(ROOT, exist_ok=True)

XL_FORMATS = {".xlsx": 51, ".xlsm": 52, ".xlsb": 50, ".xls": 56, ".csv": 6}
MACRO_EXTS = {".xlsm", ".xlsb", ".xls"}

DANGEROUS_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")
# formulas that can execute code / exfiltrate
DANGEROUS_FORMULA_TOKENS = ("cmd|", "powershell", "http://", "https://", "mshta", "wscript", "cscript")

def is_dangerous_formula(value) -> bool:
    if not isinstance(value, str):
        return False
    s = value.strip()
    if not s.startswith("="):
        return False
    low = s.lower()
    if any(t in low for t in DANGEROUS_FORMULA_TOKENS):
        return True
    # leading = followed by non-formula function? keep simple: block external links
    if "!" in s and "[" in s:
        return True
    return False
