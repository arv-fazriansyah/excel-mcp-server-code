"""Consolidated Excel Engine (COM-based, .xlsb + VBA + Ranges + Sheets + Tables + Named Ranges)."""
from __future__ import annotations

import os
import time
import uuid
import re
import hashlib
import pythoncom
import win32com.client as win32
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

# --- CONFIG ---
BACKEND = os.environ.get("EXCEL_MCP_BACKEND", "auto").lower()
ALLOW_MACROS = os.environ.get("EXCEL_MCP_ALLOW_MACROS", "0") == "1"
READONLY = os.environ.get("EXCEL_MCP_READONLY", "0") == "1"
ALLOW_VBA_WRITE = os.environ.get("EXCEL_MCP_ALLOW_VBA_WRITE", "0") == "1"
ALLOW_UNSAFE_FORMULAS = os.environ.get("EXCEL_MCP_ALLOW_UNSAFE_FORMULAS", "0") == "1"
VISIBLE = os.environ.get("EXCEL_MCP_VISIBLE", "0") == "1"
MAX_CELLS = int(os.environ.get("EXCEL_MCP_MAX_CELLS", "20000"))

def _documents_dir() -> str:
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
VB_TYPE = {1: "Module", 2: "Class", 3: "Form", 100: "Document"}
VB_EXT = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}

def abspath(path: str) -> str:
    return str(Path(path).expanduser().resolve())

def check_path(path: str) -> str:
    full = abspath(path)
    try:
        Path(full).relative_to(Path(ROOT).resolve())
    except ValueError:
        raise PermissionError(f"Path di luar EXCEL_MCP_ROOT ({ROOT}): {full}")
    return full

def cell_ref(row: int, col: int) -> str:
    letters = ""
    c = col
    while c:
        c, rem = divmod(c - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"

def sanitize(value):
    if value is None: return None
    if isinstance(value, (bool, int, float, str)):
        if isinstance(value, float) and (value != value or value in (float("inf"), float("-inf"))):
            return None
        return value
    return str(value)

def normalize_grid(raw):
    if not isinstance(raw, tuple): return [[raw]]
    return [list(r) if isinstance(r, tuple) else [r] for r in raw]

# --- COM EXECUTOR (Thread-local STA) ---
_local = threading.local() if False else __import__("threading").local()

def get_excel_app(visible: bool = False):
    if not getattr(_local, "initialized", False):
        pythoncom.CoInitialize()
        _local.initialized = True
    app = getattr(_local, "app", None)
    if app is None:
        try:
            app = win32.DispatchEx("Excel.Application")
        except Exception:
            app = win32.Dispatch("Excel.Application")
        app.Visible = visible
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 3
        _local.app = app
    return app

def run_com(fn, *args, **kwargs):
    if not getattr(_local, "initialized", False):
        pythoncom.CoInitialize()
        _local.initialized = True
    return fn(*args, **kwargs)

# --- SESSIONS ---
@dataclass
class Session:
    id: str
    workbook: Any
    app: Any
    path: str
    read_only: bool

_sessions: dict[str, Session] = {}

def session_open(path: str, read_only: bool = False, password: str | None = None) -> dict:
    check_path(path)
    app = get_excel_app(VISIBLE)
    full = abspath(path)
    wb = run_com(lambda: app.Workbooks.Open(full, 0, bool(read_only or READONLY), None, password))
    sid = str(uuid.uuid4())
    _sessions[sid] = Session(id=sid, workbook=wb, app=app, path=full, read_only=bool(wb.ReadOnly))
    return {"sessionId": sid, "path": full, "readOnly": bool(wb.ReadOnly), "hasVba": bool(wb.HasVBProject)}

def session_create(path: str, sheets: str = "", force: bool = False) -> dict:
    check_path(path)
    full = abspath(path)
    if os.path.exists(full) and not force:
        raise FileExistsError(f"File sudah ada: {full}")
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    if os.path.exists(full): os.remove(full)
    ext = os.path.splitext(full)[1].lower()
    if ext not in XL_FORMATS: raise ValueError(f"Format tidak didukung: {ext}")
    app = get_excel_app(VISIBLE)
    names = [n.strip() for n in sheets.split(",") if n.strip()]
    def _create():
        wb = app.Workbooks.Add()
        while wb.Worksheets.Count > max(1, len(names)):
            wb.Worksheets(wb.Worksheets.Count).Delete()
        for i, name in enumerate(names):
            if i < wb.Worksheets.Count: wb.Worksheets(i + 1).Name = name
            else:
                ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
                ws.Name = name
        wb.SaveAs(full, XL_FORMATS[ext])
        return wb
    wb = run_com(_create)
    sid = str(uuid.uuid4())
    _sessions[sid] = Session(id=sid, workbook=wb, app=app, path=full, read_only=False)
    return {"sessionId": sid, "path": full, "sheets": names or ["Sheet1"]}

def session_attach() -> dict:
    app = get_excel_app(VISIBLE)
    return {"workbooks": run_com(lambda: [{"name": wb.Name, "path": wb.FullName, "readOnly": bool(wb.ReadOnly)} for wb in app.Workbooks])}

def get_session(sid: str) -> Session:
    if sid not in _sessions: raise KeyError(f"Session tidak ditemukan: {sid}")
    return _sessions[sid]

def session_info(sid: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    def _info():
        sheets = [{"name": ws.Name, "index": ws.Index, "visible": int(ws.Visible),
                   "usedRange": ws.UsedRange.Address.replace("$", ""),
                   "rows": ws.UsedRange.Rows.Count, "cols": ws.UsedRange.Columns.Count} for ws in wb.Worksheets]
        return {"path": s.path, "readOnly": bool(wb.ReadOnly), "format": int(wb.FileFormat),
                "hasVba": bool(wb.HasVBProject), "sheets": sheets, "namedRanges": [n.Name for n in wb.Names]}
    return run_com(_info)

def session_save(sid: str) -> dict:
    s = get_session(sid)
    run_com(lambda: s.workbook.Save())
    return {"ok": True, "path": s.path}

def session_close(sid: str, save: bool = False) -> dict:
    s = get_session(sid)
    if save: run_com(lambda: s.workbook.Save())
    run_com(lambda: s.workbook.Close(bool(save)))
    del _sessions[sid]
    return {"ok": True, "saved": save}

def session_list() -> dict:
    return {"sessions": [{"id": sid, "path": s.path, "readOnly": s.read_only} for sid, s in _sessions.items()]}

# --- TOOLS: SHEETS ---
def sheet_list(sid: str) -> dict:
    s = get_session(sid)
    return run_com(lambda: {"sheets": [{"name": ws.Name, "index": ws.Index, "visible": int(ws.Visible)} for ws in s.workbook.Worksheets]})

def sheet_add(sid: str, name: str | None = None, after: str | None = None) -> dict:
    s = get_session(sid)
    wb = s.workbook
    def _run():
        anchor = wb.Worksheets(after) if after else wb.Worksheets(wb.Worksheets.Count)
        ws = wb.Worksheets.Add(None, anchor)
        if name: ws.Name = name
        return {"name": ws.Name, "index": ws.Index}
    return run_com(_run)

def sheet_rename(sid: str, old_name: str, new_name: str) -> dict:
    s = get_session(sid)
    run_com(lambda: setattr(s.workbook.Worksheets(old_name), "Name", new_name))
    return {"from": old_name, "to": new_name}

def sheet_delete(sid: str, name: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    def _run():
        if wb.Worksheets.Count <= 1: raise ValueError("Tidak bisa menghapus satu-satunya sheet")
        wb.Worksheets(name).Delete()
        return {"deleted": name}
    return run_com(_run)

# --- TOOLS: RANGES ---
def range_get_values(sid: str, sheet: str, range_address: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    def _run():
        ws = wb.Worksheets(sheet) if sheet else wb.ActiveSheet
        rng = ws.Range(range_address)
        grid = normalize_grid(rng.Value)
        return {"sheet": ws.Name, "range": rng.Address.replace("$", ""),
                "values": [[sanitize(v) for v in row] for row in grid],
                "rows": len(grid), "cols": len(grid[0]) if grid else 0}
    return run_com(_run)

def range_set_values(sid: str, sheet: str, range_address: str, values: Any) -> dict:
    s = get_session(sid)
    wb = s.workbook
    if not isinstance(values, list): values = [[values]]
    elif values and not isinstance(values[0], list): values = [values]
    rows = [list(r) if isinstance(r, list) else [r] for r in values]
    width = max((len(r) for r in rows), default=0)
    rows = [r + [None] * (width - len(r)) for r in rows]
    def _run():
        ws = wb.Worksheets(sheet) if sheet else wb.ActiveSheet
        rng = ws.Range(range_address)
        first_r, first_c = rng.Row, rng.Column
        last_r = first_r + len(rows) - 1
        last_c = first_c + width - 1
        addr = cell_ref(first_r, first_c)
        if len(rows) > 1 or width > 1: addr += ":" + cell_ref(last_r, last_c)
        target = ws.Range(addr)
        block = []
        for row in rows:
            out = []
            for v in row:
                out.append(pythoncom.Empty if v is None else v)
            block.append(tuple(out))
        target.Value = tuple(block)
        return {"sheet": ws.Name, "writtenRange": addr, "rows": len(rows), "cols": width}
    return run_com(_run)

def range_get_formulas(sid: str, sheet: str, range_address: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    def _run():
        ws = wb.Worksheets(sheet) if sheet else wb.ActiveSheet
        rng = ws.Range(range_address)
        grid = normalize_grid(rng.Formula)
        return {"sheet": ws.Name, "range": rng.Address.replace("$", ""), "formulas": [[sanitize(v) for v in row] for row in grid]}
    return run_com(_run)

def range_set_formula(sid: str, sheet: str, range_address: str, formula: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    def _run():
        ws = wb.Worksheets(sheet) if sheet else wb.ActiveSheet
        ws.Range(range_address).Formula = formula
        return {"sheet": ws.Name, "cell": range_address, "formula": formula}
    return run_com(_run)

# --- TOOLS: TABLES & NAMED RANGES ---
def table_list(sid: str, sheet: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    def _run():
        ws = wb.Worksheets(sheet)
        return {"tables": [{"name": lo.Name, "range": lo.Range.Address.replace("$", "")} for lo in ws.ListObjects]}
    return run_com(_run)

def table_add(sid: str, sheet: str, range_address: str, name: str | None = None, has_header: bool = True) -> dict:
    s = get_session(sid)
    wb = s.workbook
    def _run():
        ws = wb.Worksheets(sheet)
        lo = ws.ListObjects.Add(1, ws.Range(range_address), None, 1 if has_header else 2)
        if name: lo.Name = name
        return {"name": lo.Name, "range": lo.Range.Address.replace("$", "")}
    return run_com(_run)

def named_range_list(sid: str) -> dict:
    s = get_session(sid)
    return run_com(lambda: {"namedRanges": [{"name": n.Name, "refersTo": n.RefersTo} for n in s.workbook.Names]})

def named_range_add(sid: str, name: str, refers_to: str) -> dict:
    s = get_session(sid)
    run_com(lambda: s.workbook.Names.Add(Name=name, RefersTo=refers_to))
    return {"name": name, "refersTo": refers_to}

# --- TOOLS: VBA ---
def _ensure_vba(path: str):
    if os.path.splitext(path)[1].lower() not in MACRO_EXTS:
        raise ValueError("Operasi VBA memerlukan workbook macro-enabled (.xlsm/.xlsb/.xls)")

def vba_status(sid: str) -> dict:
    s = get_session(sid)
    _ensure_vba(s.path)
    def _run():
        try:
            vp = s.workbook.VBProject
            return {"projectAccess": True, "protection": "Locked" if vp.Protection == 1 else "None"}
        except Exception as exc:
            return {"projectAccess": False, "error": str(exc)}
    return run_com(_run)

def vba_list_modules(sid: str) -> dict:
    s = get_session(sid)
    _ensure_vba(s.path)
    def _run():
        vp = s.workbook.VBProject
        out = []
        for comp in vp.VBComponents:
            cm = comp.CodeModule
            out.append({"name": comp.Name, "type": VB_TYPE.get(comp.Type, str(comp.Type)), "lineCount": cm.CountOfLines})
        return {"modules": out}
    return run_com(_run)

def vba_read_module(sid: str, module_name: str) -> dict:
    s = get_session(sid)
    _ensure_vba(s.path)
    def _run():
        comp = s.workbook.VBProject.VBComponents(module_name)
        cm = comp.CodeModule
        return {"module": module_name, "source": cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""}
    return run_com(_run)

def vba_run_macro(sid: str, macro: str, args: list | None = None) -> dict:
    if not ALLOW_MACROS: raise PermissionError("Macro dinonaktifkan. Set EXCEL_MCP_ALLOW_MACROS=1")
    s = get_session(sid)
    _ensure_vba(s.path)
    target = macro if "!" in macro else f"'{s.workbook.Name}'!{macro}"
    return {"macro": target, "result": run_com(lambda: s.app.Run(target, *(args or [])))}
