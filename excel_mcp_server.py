#!/usr/bin/env python3
"""
excel-mcp: MCP server untuk Excel (.xlsx .xlsm .xlsb .xls) + VBA.

Backend:
  COM  (Windows + Microsoft Excel + pywin32)
       Fitur lengkap: baca/tulis .xlsb, VBA (list/baca/tulis/hapus/impor/ekspor/jalankan),
       format, chart, tabel, konversi antar format.
  FILE (lintas platform, tanpa Excel: openpyxl / pyxlsb / oletools)
       Terbatas: .xlsx/.xlsm baca-tulis, .xlsb baca saja, VBA baca saja.

Env:
  EXCEL_MCP_BACKEND       auto | com | file         (default auto)
  EXCEL_MCP_ALLOW_MACROS  1 untuk mengizinkan eksekusi macro (default 0)
  EXCEL_MCP_ROOT          batasi akses file ke folder ini
  EXCEL_MCP_VISIBLE       1 untuk menampilkan jendela Excel
  EXCEL_MCP_MAX_CELLS     batas sel per pembacaan (default 20000)
  EXCEL_MCP_ALLOW_EVAL    1 untuk mengaktifkan excel_eval (akses penuh, default 0)
  EXCEL_MCP_READONLY      1 = blokir semua tool yang mengubah data/file (default 0)
  EXCEL_MCP_ALLOW_VBA_WRITE  1 = izinkan tulis/hapus/impor VBA dan UserForm (default 0)
  EXCEL_MCP_ALLOW_UNSAFE_FORMULAS  1 = nonaktifkan filter formula berbahaya (default 0)
"""
from __future__ import annotations

import atexit
import functools
import json
import os
import platform
import queue
import re
import threading
import time
import zipfile
from typing import Any

try:
    from mcp.server import MCPServer as FastMCP  # mcp 2.x (panduan 2026-07-28)
except ImportError:
    from mcp.server.fastmcp import FastMCP  # mcp 1.x

mcp = FastMCP("excel")

BACKEND = os.environ.get("EXCEL_MCP_BACKEND", "auto").lower()
USE_COM = BACKEND == "com" or (BACKEND == "auto" and platform.system() == "Windows")
ALLOW_MACROS = os.environ.get("EXCEL_MCP_ALLOW_MACROS", "0") == "1"


def _documents_dir() -> str:
    """Folder Documents milik user (menghormati redirection/OneDrive di Windows)."""
    if platform.system() == "Windows":
        try:
            import ctypes
            buf = ctypes.create_unicode_buffer(260)
            if ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf) == 0 and buf.value:
                return buf.value
        except Exception:
            pass
    return os.path.join(os.path.expanduser("~"), "Documents")


def _default_root() -> str:
    """Folder kerja default: <Documents>/excel_files (dibuat otomatis)."""
    return os.path.join(_documents_dir(), "excel_files")


ROOT = os.environ.get("EXCEL_MCP_ROOT") or _default_root()
try:
    os.makedirs(ROOT, exist_ok=True)
except OSError:
    pass
MAX_CELLS = int(os.environ.get("EXCEL_MCP_MAX_CELLS", "20000"))

READONLY = os.environ.get("EXCEL_MCP_READONLY", "0") == "1"
ALLOW_VBA_WRITE = os.environ.get("EXCEL_MCP_ALLOW_VBA_WRITE", "0") == "1"
ALLOW_UNSAFE_FORMULAS = os.environ.get("EXCEL_MCP_ALLOW_UNSAFE_FORMULAS", "0") == "1"

WRITE_TOOLS = {
    "excel_create", "excel_save", "excel_save_as", "excel_convert", "excel_export_pdf", "refresh_all",
    "protect_workbook", "sheet_add", "sheet_rename", "sheet_delete", "sheet_copy", "sheet_move",
    "sheet_visibility", "protect_sheet", "page_setup", "range_write", "range_set_formula", "range_clear",
    "range_replace", "range_format", "range_size", "range_merge", "range_copy", "range_structure",
    "range_sort", "autofilter", "remove_duplicates", "freeze_panes", "csv_import", "data_validation",
    "conditional_format", "table_add", "table_edit", "chart_add", "chart_edit", "pivot_add", "pivot_refresh",
    "query_add", "name_define", "name_delete", "comment_set", "hyperlink_add", "image_add", "button_add",
    "shape_delete", "vba_export_module", "vba_run_macro", "excel_eval", "goal_seek", "slicer_add",
    "slicer_delete", "sparklines_add", "sparklines_delete",
}
VBA_WRITE_TOOLS = {"vba_write_module", "vba_delete_module", "vba_import_module",
                   "vba_userform_create", "vba_userform_add_control"}


def _check_write(name: str, vba: bool | None = None) -> None:
    if READONLY:
        raise PermissionError(f"Server dalam mode read-only (EXCEL_MCP_READONLY=1): tool '{name}' diblokir.")
    if vba is None:
        vba = name in VBA_WRITE_TOOLS
    if vba and not ALLOW_VBA_WRITE:
        raise PermissionError("Tulis VBA dinonaktifkan. Set env EXCEL_MCP_ALLOW_VBA_WRITE=1 lalu restart server.")


# --------------------------------------------------------------------------- COM executor
class _ComExecutor:
    """Jalankan semua panggilan COM di satu thread STA agar aman dan tidak crash.

    pywin32 membuat objek COM yang terikat ke apartment/thread pembuatnya. MCP
    dapat memanggil tool dari thread worker yang berbeda-beda; tanpa penjagaan
    ini Excel kadang menutup koneksi atau meninggalkan proses yatim.
    """

    RETRY_HRESULTS = {
        -2147418111,  # RPC_E_CALL_REJECTED
        -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
        -2147023174,  # RPC_E_SERVER_UNAVAILABLE
    }

    def __init__(self) -> None:
        self._q: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._tid: int | None = None
        self._lock = threading.Lock()

    def _ensure(self) -> None:
        with self._lock:
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(target=self._loop, name="excel-com", daemon=True)
                self._thread.start()

    def _loop(self) -> None:
        import pythoncom

        pythoncom.CoInitialize()
        self._tid = threading.get_ident()
        try:
            while True:
                job = self._q.get()
                if job is None:
                    break
                fn, args, kwargs, ev, box = job
                for attempt in range(3):
                    try:
                        box["result"] = fn(*args, **kwargs)
                        break
                    except BaseException as e:  # diteruskan ke pemanggil
                        if attempt < 2 and self._retryable(e):
                            time.sleep(0.4 * (attempt + 1))
                            continue
                        box["error"] = e
                        break
                ev.set()
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    @classmethod
    def _retryable(cls, e: BaseException) -> bool:
        hr = getattr(e, "hresult", None)
        if hr is None and getattr(e, "args", None) and isinstance(e.args[0], int):
            hr = e.args[0]
        return hr in cls.RETRY_HRESULTS

    def run(self, fn, *args, **kwargs):
        if self._tid == threading.get_ident():
            return fn(*args, **kwargs)
        self._ensure()
        ev = threading.Event()
        box: dict = {}
        self._q.put((fn, args, kwargs, ev, box))
        ev.wait()
        if "error" in box:
            raise box["error"]
        return box["result"]

    def stop(self) -> None:
        if self._thread is not None and self._thread.is_alive():
            self._q.put(None)
            self._thread.join(timeout=5)


_COM = _ComExecutor()

_real_tool = mcp.tool


def _tool(*a, **kw):
    deco = _real_tool(*a, **kw)

    def apply(fn):
        write = fn.__name__ in WRITE_TOOLS or fn.__name__ in VBA_WRITE_TOOLS

        @functools.wraps(fn)
        def guarded(*args, **kwargs):
            if write:
                _check_write(fn.__name__)
            return _COM.run(fn, *args, **kwargs)

        return deco(guarded)

    return apply


mcp.tool = _tool

_UNSAFE_FORMULA = re.compile(
    r"(?i)(\bWEBSERVICE\s*\(|\bFILTERXML\s*\(|\bCALL\s*\(|\bREGISTER(\.ID)?\s*\(|\bEXEC\s*\("
    r"|\|\s*['\"]?[^!]*!|\[[^\]]+\.(xls[xmb]?|xlam)\]"
    r"|\b(cmd|powershell|mshta|wscript|cscript|regsvr32|rundll32)\s*\|)"
)


def _check_formulas(values: list[list[Any]]) -> None:
    """Tolak formula berbahaya (DDE, WEBSERVICE, link eksternal, fungsi XLM)."""
    if ALLOW_UNSAFE_FORMULAS:
        return
    for row in values:
        for v in row:
            if isinstance(v, str) and v and v[0] in "=+-@" and _UNSAFE_FORMULA.search(v):
                raise ValueError(f"Formula ditolak karena berpotensi berbahaya: {v[:60]!r}. "
                                 "Set EXCEL_MCP_ALLOW_UNSAFE_FORMULAS=1 jika memang disengaja.")

EXT_FMT = {".xlsx": 51, ".xlsm": 52, ".xlsb": 50, ".xls": 56, ".csv": 6}
MACRO_EXT = {".xlsm", ".xlsb", ".xls"}
ERRORS = {
    -2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?",
    -2146826288: "#NULL!", -2146826252: "#NUM!", -2146826265: "#REF!",
    -2146826273: "#VALUE!",
}
CHART_TYPES = {
    "column": 51, "column_stacked": 52, "bar": 57, "line": 4, "line_markers": 65,
    "pie": 5, "scatter": -4169, "area": 1,
}
VBA_TYPES = {1: "standard", 2: "class", 3: "userform", 100: "document"}


# --------------------------------------------------------------------------- util
def _long_path(path: str) -> str:
    """Normalisasi path short-name (8.3) menjadi long path agar validasi ROOT akurat."""
    if platform.system() != "Windows":
        return path
    try:
        import ctypes

        target = path if os.path.exists(path) else os.path.dirname(path)
        if not target:
            return path
        buf = ctypes.create_unicode_buffer(32768)
        n = ctypes.windll.kernel32.GetLongPathNameW(target, buf, 32768)
        if 0 < n < 32768:
            long = buf.value
            return long if target == path else os.path.join(long, os.path.basename(path))
    except Exception:
        pass
    return path


def _p(path: str) -> str:
    ap = _long_path(os.path.abspath(os.path.expanduser(path)))
    if ROOT:
        r = _long_path(os.path.abspath(os.path.expanduser(ROOT)))
        try:
            inside = os.path.normcase(os.path.commonpath([os.path.normcase(ap), os.path.normcase(r)])) == os.path.normcase(r)
        except ValueError:
            inside = False
        if not inside:
            raise PermissionError(f"Path di luar EXCEL_MCP_ROOT: {ap}")
    return ap


def _ext(path: str) -> str:
    return os.path.splitext(path)[1].lower()


def _clean(v: Any) -> Any:
    if isinstance(v, int) and not isinstance(v, bool) and v in ERRORS:
        return ERRORS[v]
    if hasattr(v, "isoformat"):
        return v.isoformat()
    return v


def _grid(v: Any) -> list[list[Any]]:
    if not isinstance(v, tuple):
        return [[_clean(v)]]
    return [[_clean(c) for c in row] for row in v]


def _rgb(h: str) -> int:
    h = h.lstrip("#")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


def _pad(values: list[list[Any]]) -> tuple:
    if not values or not isinstance(values[0], (list, tuple)):
        raise ValueError("values harus array 2D, contoh [[1,2],[3,4]]")
    w = max(len(r) for r in values)
    return tuple(tuple(list(r) + [None] * (w - len(r))) for r in values)


def _range_at(ws, row: int, col: int, nrows: int = 1, ncols: int = 1):
    """Bangun Range dari koordinat tanpa memakai Range.Offset/Resize (tidak andal di pywin32 3.14)."""
    return ws.Range(ws.Cells(row, col), ws.Cells(row + nrows - 1, col + ncols - 1))


def _need_com() -> None:
    if not USE_COM:
        raise RuntimeError(
            "Tool ini butuh backend COM (Windows + Microsoft Excel + pywin32). "
            "Backend FILE hanya mendukung: sheet_list, range_read, range_write, excel_create (.xlsx), "
            "vba_list_modules, vba_read_module, vba_inspect_file."
        )


# --------------------------------------------------------------------------- COM core
class _Com:
    app = None

    def _drop(self) -> None:
        """Buang instance Excel yang sudah tidak sehat, tanpa meninggalkan proses yatim."""
        app, self.app = self.app, None
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass

    def application(self):
        import pythoncom
        import win32com.client as wc

        if self.app is not None:
            try:
                _ = self.app.Version
            except Exception:
                self._drop()
        if self.app is None:
            pythoncom.CoInitialize()
            app = wc.DispatchEx("Excel.Application")
            app.Visible = os.environ.get("EXCEL_MCP_VISIBLE", "0") == "1"
            app.DisplayAlerts = False
            app.AutomationSecurity = 1 if ALLOW_MACROS else 3  # 3 = macro dinonaktifkan saat buka
            self.app = app
        return self.app

    def wb(self, path: str | None = None):
        app = self.application()
        if path:
            ap = _p(path)
            for w in app.Workbooks:
                if os.path.normcase(w.FullName) == os.path.normcase(ap):
                    return w
            if not os.path.exists(ap):
                raise FileNotFoundError(ap)
            return app.Workbooks.Open(Filename=ap, UpdateLinks=0, ReadOnly=READONLY)
        if app.Workbooks.Count == 0:
            raise RuntimeError("Tidak ada workbook terbuka. Panggil excel_open atau excel_create.")
        return app.ActiveWorkbook

    def quit(self):
        app, self.app = self.app, None
        if app is None:
            return
        try:
            for w in list(app.Workbooks):
                try:
                    w.Close(False)
                except Exception:
                    pass
        finally:
            try:
                app.Quit()
            except Exception:
                pass


com = _Com()


def _shutdown() -> None:
    try:
        _COM.run(com.quit)
    except Exception:
        pass
    finally:
        _COM.stop()


atexit.register(_shutdown)


def _ws(wb, sheet: str | None):
    if sheet is None:
        return wb.ActiveSheet
    try:
        return wb.Worksheets(sheet)
    except Exception:
        if str(sheet).isdigit():
            return wb.Worksheets(int(sheet))
        raise ValueError(f"Sheet '{sheet}' tidak ditemukan")


def _has_vba(wb) -> bool | None:
    try:
        for c in wb.VBProject.VBComponents:
            if c.CodeModule.CountOfLines > 0:
                return True
        return False
    except Exception:
        return None


def _vbproj(wb):
    try:
        return wb.VBProject
    except Exception as e:
        raise RuntimeError(
            "Akses VBA ditolak. Aktifkan: Excel > File > Options > Trust Center > Trust Center Settings > "
            f"Macro Settings > 'Trust access to the VBA project object model'. Detail: {e}"
        )


def _comp(proj, name: str):
    for c in proj.VBComponents:
        if c.Name.lower() == name.lower():
            return c
    return None


def _crlf(code: str) -> str:
    return code.replace("\r\n", "\n").replace("\n", "\r\n")


# --------------------------------------------------------------------------- VBA dari file (tanpa Excel)
def _vba_from_file(path: str) -> list[dict]:
    from oletools.olevba import VBA_Parser

    ap = _p(path)
    ext = _ext(ap)
    parsers = []
    if ext in (".xlsm", ".xlsb"):
        with zipfile.ZipFile(ap) as z:
            for n in z.namelist():
                if n.lower().endswith("vbaproject.bin"):
                    parsers.append(VBA_Parser(n, data=z.read(n)))
    else:
        parsers.append(VBA_Parser(ap))
    out: list[dict] = []
    for vp in parsers:
        try:
            if vp.detect_vba_macros():
                for (_f, stream, vname, code) in vp.extract_macros():
                    out.append({"module": vname, "stream": stream, "code": code})
        finally:
            vp.close()
    return out


# --------------------------------------------------------------------------- FILE backend (openpyxl / pyxlsb)
def _bounds(address: str):
    from openpyxl.utils.cell import range_boundaries

    return range_boundaries(address)  # min_col, min_row, max_col, max_row


def _file_sheets(path: str) -> list[str]:
    ap = _p(path)
    if _ext(ap) == ".xlsb":
        from pyxlsb import open_workbook

        with open_workbook(ap) as wb:
            return list(wb.sheets)
    import openpyxl

    wb = openpyxl.load_workbook(ap, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _file_read(path: str, sheet: str | None, address: str | None, formulas: bool,
               offset: int = 0, limit: int | None = None) -> dict:
    ap = _p(path)
    offset = max(0, offset)
    if _ext(ap) == ".xlsb":
        from pyxlsb import open_workbook

        with open_workbook(ap) as wb:
            name = sheet or wb.sheets[0]
            cells: dict[tuple[int, int], Any] = {}
            with wb.get_sheet(name) as ws:
                for row in ws.rows(sparse=True):
                    for c in row:
                        if c.v is not None:
                            cells[(c.r + 1, c.c + 1)] = c.v
        if not cells and not address:
            return {"sheet": name, "values": [], "total_rows": 0, "next_offset": None, "truncated": False}
        if address:
            c0, r0, c1, r1 = _bounds(address)
        else:
            r0 = min(r for r, _ in cells); r1 = max(r for r, _ in cells)
            c0 = min(c for _, c in cells); c1 = max(c for _, c in cells)
        width = c1 - c0 + 1
        cap = max(1, MAX_CELLS // width)
        per = max(1, min(limit, cap)) if limit else cap
        start = r0 + offset
        end = min(r1, start + per - 1)
        grid = [[cells.get((r, c)) for c in range(c0, c1 + 1)] for r in range(start, end + 1)]
        nxt = offset + len(grid) if end < r1 else None
        return {"sheet": name, "values": grid, "total_rows": r1 - r0 + 1, "next_offset": nxt,
                "truncated": nxt is not None, "note": "Tanggal di .xlsb berupa serial number."}
    import openpyxl

    wb = openpyxl.load_workbook(ap, data_only=not formulas, read_only=True)
    try:
        ws = wb[sheet] if sheet else wb.active
        if address:
            c0, r0, c1, r1 = _bounds(address)
            it = ws.iter_rows(min_row=r0 + offset, max_row=r1, min_col=c0, max_col=c1, values_only=True)
            total = r1 - r0 + 1
        else:
            it = ws.iter_rows(min_row=1 + offset, values_only=True)
            total = ws.max_row
        grid: list[list[Any]] = []
        count, nxt = 0, None
        for row in it:
            if (limit and len(grid) >= limit) or (grid and count + len(row) > MAX_CELLS):
                nxt = offset + len(grid)
                break
            count += len(row)
            grid.append([_clean(v) for v in row])
        return {"sheet": ws.title, "values": grid, "total_rows": total, "next_offset": nxt,
                "truncated": nxt is not None}
    finally:
        wb.close()


def _file_write(path: str, sheet: str | None, start: str, values: list[list[Any]]) -> dict:
    import openpyxl
    from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

    ap = _p(path)
    ext = _ext(ap)
    if ext not in (".xlsx", ".xlsm"):
        raise RuntimeError("Backend FILE hanya bisa menulis .xlsx/.xlsm. Untuk .xlsb/.xls gunakan backend COM.")
    wb = openpyxl.load_workbook(ap, keep_vba=(ext == ".xlsm"))
    ws = wb[sheet] if sheet else wb.active
    col, row0 = coordinate_from_string(start)
    c0 = column_index_from_string(col)
    for i, row in enumerate(values):
        for j, v in enumerate(row):
            ws.cell(row=row0 + i, column=c0 + j, value=v)
    wb.save(ap)
    return {"sheet": ws.title, "written_rows": len(values), "saved": True}


# =========================================================================== TOOLS: workbook
@mcp.tool()
def excel_status() -> dict:
    """Info backend, versi Excel, dan workbook yang sedang terbuka."""
    info: dict[str, Any] = {"backend": "com" if USE_COM else "file", "macros_allowed": ALLOW_MACROS,
                            "root": ROOT, "platform": platform.system(), "readonly": READONLY,
                            "vba_write_allowed": ALLOW_VBA_WRITE,
                            "eval_allowed": os.environ.get("EXCEL_MCP_ALLOW_EVAL", "0") == "1"}
    if USE_COM and com.app is not None:
        try:
            info["excel_version"] = com.app.Version
            info["open_workbooks"] = [w.FullName for w in com.app.Workbooks]
        except Exception:
            info["excel_running"] = False
    return info


@mcp.tool()
def excel_open(path: str, read_only: bool = False, password: str | None = None) -> dict:
    """Buka workbook (.xlsx/.xlsm/.xlsb/.xls). Backend FILE: hanya memverifikasi file & daftar sheet."""
    ap = _p(path)
    if not os.path.exists(ap):
        raise FileNotFoundError(ap)
    if not USE_COM:
        return {"path": ap, "sheets": _file_sheets(ap)}
    app = com.application()
    for w in app.Workbooks:
        if os.path.normcase(w.FullName) == os.path.normcase(ap):
            return {"path": ap, "already_open": True, "read_only": bool(read_only or READONLY),
                    "sheets": [s.Name for s in w.Worksheets], "has_vba": _has_vba(w)}
    kw: dict[str, Any] = {"Filename": ap, "UpdateLinks": 0, "ReadOnly": read_only or READONLY}
    if password:
        kw["Password"] = password
    w = app.Workbooks.Open(**kw)
    return {"path": ap, "sheets": [s.Name for s in w.Worksheets], "has_vba": _has_vba(w)}


@mcp.tool()
def excel_create(path: str, sheet_names: list[str] | None = None) -> dict:
    """Buat workbook baru; format ditentukan ekstensi (.xlsx .xlsm .xlsb .xls). Backend FILE: hanya .xlsx."""
    ap = _p(path)
    ext = _ext(ap)
    if ext not in EXT_FMT or ext == ".csv":
        raise ValueError("Ekstensi harus .xlsx/.xlsm/.xlsb/.xls")
    if not USE_COM:
        if ext != ".xlsx":
            raise RuntimeError("Backend FILE hanya bisa membuat .xlsx.")
        import openpyxl

        wb = openpyxl.Workbook()
        names = sheet_names or ["Sheet1"]
        wb.active.title = names[0]
        for n in names[1:]:
            wb.create_sheet(n)
        wb.save(ap)
        return {"path": ap, "sheets": names}
    app = com.application()
    wb = app.Workbooks.Add()
    if sheet_names:
        while wb.Worksheets.Count < len(sheet_names):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, n in enumerate(sheet_names, 1):
            wb.Worksheets(i).Name = n
    wb.SaveAs(Filename=ap, FileFormat=EXT_FMT[ext])
    return {"path": ap, "sheets": [s.Name for s in wb.Worksheets]}


@mcp.tool()
def excel_save(path: str | None = None) -> dict:
    """Simpan workbook (default: workbook aktif)."""
    _need_com()
    wb = com.wb(path)
    wb.Save()
    return {"saved": wb.FullName}


@mcp.tool()
def excel_save_as(new_path: str, path: str | None = None, allow_macro_loss: bool = False) -> dict:
    """Simpan sebagai format lain (.xlsx .xlsm .xlsb .xls .csv). Menolak jika VBA akan hilang kecuali allow_macro_loss."""
    _need_com()
    wb = com.wb(path)
    np_ = _p(new_path)
    ext = _ext(np_)
    if ext not in EXT_FMT:
        raise ValueError("Ekstensi tidak didukung")
    if ext not in MACRO_EXT:
        has = _has_vba(wb)
        if has is None:
            has = _ext(wb.FullName) in MACRO_EXT
        if has and not allow_macro_loss:
            raise RuntimeError(f"Workbook berisi VBA; menyimpan sebagai {ext} akan MENGHAPUS macro. "
                               "Pakai .xlsm/.xlsb atau set allow_macro_loss=true.")
    wb.SaveAs(Filename=np_, FileFormat=EXT_FMT[ext])
    return {"saved": np_}


@mcp.tool()
def excel_convert(src: str, dst: str, allow_macro_loss: bool = False) -> dict:
    """Konversi file antar format, mis. .xlsx -> .xlsb atau .xlsb -> .xlsm. VBA ikut terbawa jika format tujuan mendukung."""
    _need_com()
    app = com.application()
    s, d = _p(src), _p(dst)
    if _ext(d) not in EXT_FMT:
        raise ValueError("Ekstensi tujuan tidak didukung")
    wb = app.Workbooks.Open(Filename=s, UpdateLinks=0, ReadOnly=True)
    try:
        if _ext(d) not in MACRO_EXT and not allow_macro_loss:
            has = _has_vba(wb)
            if has or (has is None and _ext(s) in MACRO_EXT):
                raise RuntimeError("Sumber berisi VBA; format tujuan tidak menyimpan macro. Set allow_macro_loss=true untuk lanjut.")
        wb.SaveAs(Filename=d, FileFormat=EXT_FMT[_ext(d)])
    finally:
        wb.Close(False)
    return {"converted": d, "size_bytes": os.path.getsize(d)}


@mcp.tool()
def excel_export_pdf(out_path: str, path: str | None = None, sheet: str | None = None) -> dict:
    """Ekspor workbook (atau satu sheet) ke PDF."""
    _need_com()
    wb = com.wb(path)
    target = _ws(wb, sheet) if sheet else wb
    op = _p(out_path)
    target.ExportAsFixedFormat(0, op)
    return {"exported": op}


@mcp.tool()
def excel_close(path: str | None = None, save: bool = False) -> dict:
    """Tutup workbook yang SEDANG terbuka. Tidak membuka file baru jika path belum terbuka."""
    _need_com()
    if save:
        _check_write("excel_close")
    app = com.application()
    if path:
        ap = _p(path)
        wb = next((w for w in app.Workbooks if os.path.normcase(w.FullName) == os.path.normcase(ap)), None)
        if wb is None:
            return {"closed": None, "saved": False, "reason": "not_open", "path": ap}
    else:
        if app.Workbooks.Count == 0:
            raise RuntimeError("Tidak ada workbook terbuka.")
        wb = app.ActiveWorkbook
    name = wb.FullName
    wb.Close(bool(save))
    return {"closed": name, "saved": save}


@mcp.tool()
def excel_quit() -> dict:
    """Tutup semua workbook tanpa menyimpan dan keluar dari Excel."""
    _need_com()
    com.quit()
    return {"quit": True}


@mcp.tool()
def excel_calculate(full: bool = False, path: str | None = None) -> dict:
    """Hitung ulang formula."""
    _need_com()
    app = com.application()
    com.wb(path)
    app.CalculateFull() if full else app.Calculate()
    return {"calculated": True}


# =========================================================================== TOOLS: sheet
@mcp.tool()
def sheet_list(path: str | None = None) -> dict:
    """Daftar sheet beserta used range."""
    if not USE_COM:
        if not path:
            raise ValueError("path wajib pada backend FILE")
        return {"sheets": _file_sheets(path)}
    wb = com.wb(path)
    return {"sheets": [{"name": s.Name, "visible": s.Visible == -1, "used_range": s.UsedRange.GetAddress(False, False)}
                       for s in wb.Worksheets]}


@mcp.tool()
def sheet_add(name: str, after: str | None = None, path: str | None = None) -> dict:
    """Tambah sheet (default di akhir)."""
    _need_com()
    wb = com.wb(path)
    if any(s.Name.lower() == name.lower() for s in wb.Worksheets):
        raise ValueError(f"Sheet '{name}' sudah ada")
    anchor = _ws(wb, after) if after else wb.Worksheets(wb.Worksheets.Count)
    ws = wb.Worksheets.Add(After=anchor)
    ws.Name = name
    return {"added": name}


@mcp.tool()
def sheet_rename(old: str, new: str, path: str | None = None) -> dict:
    """Ganti nama sheet."""
    _need_com()
    _ws(com.wb(path), old).Name = new
    return {"renamed": new}


@mcp.tool()
def sheet_delete(name: str, path: str | None = None) -> dict:
    """Hapus sheet."""
    _need_com()
    _ws(com.wb(path), name).Delete()
    return {"deleted": name}


# =========================================================================== TOOLS: range
@mcp.tool()
def range_read(sheet: str | None = None, address: str | None = None, path: str | None = None,
               formulas: bool = False, offset: int = 0, limit: int | None = None) -> dict:
    """Baca range sebagai array 2D. address kosong = used range. formulas=true mengembalikan formula.
    Paging: offset (baris awal, 0-based relatif ke range) dan limit (jumlah baris). Jika masih ada sisa,
    hasil memuat next_offset; panggil lagi dengan offset=next_offset sampai null."""
    if not USE_COM:
        if not path:
            raise ValueError("path wajib pada backend FILE")
        return _file_read(path, sheet, address, formulas, offset, limit)
    offset = max(0, offset)
    wb = com.wb(path)
    ws = _ws(wb, sheet)
    rng = ws.Range(address) if address else ws.UsedRange
    total, cols = rng.Rows.Count, rng.Columns.Count
    cap = max(1, MAX_CELLS // cols)
    n = max(1, min(limit, cap)) if limit else cap
    if offset >= total:
        return {"sheet": ws.Name, "address": None, "values": [], "total_rows": total,
                "next_offset": None, "truncated": False}
    n = min(n, total - offset)
    sub = _range_at(ws, rng.Row + offset, rng.Column, n, cols)
    data = sub.Formula if formulas else sub.Value
    nxt = offset + n if offset + n < total else None
    return {"sheet": ws.Name, "address": sub.GetAddress(False, False), "values": _grid(data),
            "total_rows": total, "next_offset": nxt, "truncated": nxt is not None}


@mcp.tool()
def range_write(values: list[list[Any]], sheet: str | None = None, start: str = "A1",
                path: str | None = None, save: bool = False) -> dict:
    """Tulis array 2D mulai dari sel `start`. String berawalan '=' dianggap formula."""
    _check_formulas(values)
    if not USE_COM:
        if not path:
            raise ValueError("path wajib pada backend FILE")
        return _file_write(path, sheet, start, values)
    wb = com.wb(path)
    ws = _ws(wb, sheet)
    data = _pad(values)
    _start = ws.Range(start)
    rng = _range_at(ws, _start.Row, _start.Column, len(data), len(data[0]))
    rng.Value = data
    if save:
        wb.Save()
    return {"sheet": ws.Name, "address": rng.GetAddress(False, False), "saved": save}


@mcp.tool()
def range_set_formula(address: str, formula: str, sheet: str | None = None, path: str | None = None) -> dict:
    """Isi seluruh range dengan satu formula (referensi relatif menyesuaikan), mis. address='C2:C100', formula='=A2*B2'."""
    _need_com()
    _check_formulas([[formula]])
    ws = _ws(com.wb(path), sheet)
    ws.Range(address).Formula = formula
    return {"address": address, "formula": formula}


@mcp.tool()
def range_clear(address: str, sheet: str | None = None, path: str | None = None, contents_only: bool = True) -> dict:
    """Hapus isi (atau isi+format) range."""
    _need_com()
    rng = _ws(com.wb(path), sheet).Range(address)
    rng.ClearContents() if contents_only else rng.Clear()
    return {"cleared": address}


@mcp.tool()
def range_find(text: str, sheet: str | None = None, path: str | None = None,
               match_case: bool = False, whole_cell: bool = False, limit: int = 200) -> dict:
    """Cari teks/nilai di sheet; kembalikan alamat sel."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    rng = ws.UsedRange
    found, first = [], None
    c = rng.Find(What=text, LookIn=-4163, LookAt=1 if whole_cell else 2, MatchCase=match_case)
    while c is not None and len(found) < limit:
        addr = c.GetAddress(False, False)
        if first is None:
            first = addr
        elif addr == first:
            break
        found.append({"cell": addr, "value": _clean(c.Value)})
        c = rng.FindNext(c)
    return {"sheet": ws.Name, "matches": found}


@mcp.tool()
def range_format(address: str, sheet: str | None = None, path: str | None = None,
                 bold: bool | None = None, italic: bool | None = None, font_size: float | None = None,
                 font_color: str | None = None, fill_color: str | None = None, number_format: str | None = None,
                 h_align: str | None = None, wrap: bool | None = None, borders: bool | None = None,
                 autofit: bool = False) -> dict:
    """Format range. Warna hex '#RRGGBB'. h_align: left|center|right. number_format mis. '#,##0.00' atau 'dd/mm/yyyy'."""
    _need_com()
    rng = _ws(com.wb(path), sheet).Range(address)
    if bold is not None: rng.Font.Bold = bold
    if italic is not None: rng.Font.Italic = italic
    if font_size is not None: rng.Font.Size = font_size
    if font_color: rng.Font.Color = _rgb(font_color)
    if fill_color: rng.Interior.Color = _rgb(fill_color)
    if number_format: rng.NumberFormat = number_format
    if h_align: rng.HorizontalAlignment = {"left": -4131, "center": -4108, "right": -4152}[h_align.lower()]
    if wrap is not None: rng.WrapText = wrap
    if borders is not None: rng.Borders.LineStyle = 1 if borders else -4142
    if autofit: rng.Columns.AutoFit()
    return {"formatted": address}


@mcp.tool()
def table_add(address: str, name: str, sheet: str | None = None, path: str | None = None,
              style: str = "TableStyleMedium2", has_headers: bool = True) -> dict:
    """Ubah range menjadi Excel Table (ListObject)."""
    _need_com()
    wb = com.wb(path)
    for s in wb.Worksheets:
        for lo in s.ListObjects:
            if lo.Name.lower() == name.lower():
                raise ValueError(f"Tabel '{name}' sudah ada di sheet '{s.Name}'")
    ws = _ws(wb, sheet)
    lo = ws.ListObjects.Add(1, ws.Range(address), None, 1 if has_headers else 2)
    lo.Name = name
    lo.TableStyle = style
    return {"table": name, "range": lo.Range.GetAddress(False, False)}


@mcp.tool()
def chart_add(source: str, chart_type: str = "column", title: str | None = None, sheet: str | None = None,
              path: str | None = None, left: float | None = None, top: float | None = None,
              width: float = 400, height: float = 250) -> dict:
    """Tambah chart dari range sumber (termasuk header). Tipe: column, column_stacked, bar, line, line_markers, pie, scatter, area."""
    _need_com()
    if chart_type not in CHART_TYPES:
        raise ValueError(f"chart_type harus salah satu dari {list(CHART_TYPES)}")
    ws = _ws(com.wb(path), sheet)
    parts = [p.strip() for p in source.split(",") if p.strip()]
    if len(parts) == 1:
        src = ws.Range(parts[0])
    else:
        app = com.application()
        src = ws.Range(parts[0])
        for p in parts[1:]:
            src = app.Union(src, ws.Range(p))
    shp = ws.Shapes.AddChart2(-1, CHART_TYPES[chart_type],
                              left if left is not None else src.Left + src.Width + 20,
                              top if top is not None else src.Top, width, height)
    ch = shp.Chart
    ch.SetSourceData(src)
    if title:
        ch.HasTitle = True
        ch.ChartTitle.Text = title
    return {"chart": shp.Name}


@mcp.tool()
def name_define(name: str, refers_to: str, path: str | None = None) -> dict:
    """Definisikan named range, mis. refers_to='=Sheet1!$A$1:$B$10'."""
    _need_com()
    wb = com.wb(path)
    for n in wb.Names:
        if n.Name.lower() == name.lower():
            raise ValueError(f"Named range '{name}' sudah ada (refers_to={n.RefersTo})")
    wb.Names.Add(Name=name, RefersTo=refers_to)
    return {"name": name, "refers_to": refers_to}


# =========================================================================== TOOLS: VBA
@mcp.tool()
def vba_inspect_file(path: str) -> dict:
    """Baca VBA langsung dari file (.xlsm/.xlsb/.xls) TANPA membuka Excel dan tanpa menjalankan apa pun. Aman & lintas platform."""
    mods = _vba_from_file(path)
    return {"has_vba": bool(mods), "modules": [{"module": m["module"], "stream": m["stream"],
                                                  "lines": len(m["code"].splitlines())} for m in mods]}


@mcp.tool()
def vba_list_modules(path: str | None = None) -> dict:
    """Daftar modul VBA (nama, tipe, jumlah baris)."""
    if not USE_COM:
        if not path:
            raise ValueError("path wajib pada backend FILE")
        return {"modules": [{"name": m["module"], "lines": len(m["code"].splitlines())} for m in _vba_from_file(path)]}
    proj = _vbproj(com.wb(path))
    return {"modules": [{"name": c.Name, "type": VBA_TYPES.get(c.Type, c.Type), "lines": c.CodeModule.CountOfLines}
                        for c in proj.VBComponents]}


@mcp.tool()
def vba_read_module(name: str, path: str | None = None) -> dict:
    """Baca kode satu modul VBA."""
    if not USE_COM:
        if not path:
            raise ValueError("path wajib pada backend FILE")
        for m in _vba_from_file(path):
            if m["module"].lower() == name.lower():
                return {"name": m["module"], "code": m["code"]}
        raise ValueError(f"Modul '{name}' tidak ditemukan")
    proj = _vbproj(com.wb(path))
    c = _comp(proj, name)
    if c is None:
        raise ValueError(f"Modul '{name}' tidak ditemukan")
    cm = c.CodeModule
    n = cm.CountOfLines
    return {"name": c.Name, "type": VBA_TYPES.get(c.Type, c.Type), "lines": n,
            "code": cm.Lines(1, n).replace("\r\n", "\n") if n else ""}


@mcp.tool()
def vba_write_module(name: str, code: str, module_type: str = "standard", mode: str = "replace",
                     line: int | None = None, path: str | None = None, save: bool = False) -> dict:
    """Tulis kode VBA. module_type: standard|class (diabaikan jika modul sudah ada, mis. ThisWorkbook/Sheet1).
    mode: replace (ganti seluruh isi) | append | insert (butuh `line`). save=true menyimpan workbook
    (harus .xlsm/.xlsb agar macro tersimpan)."""
    _need_com()
    if mode not in ("replace", "append", "insert"):
        raise ValueError("mode harus replace|append|insert")
    wb = com.wb(path)
    proj = _vbproj(wb)
    c = _comp(proj, name)
    created = False
    if c is None:
        if module_type not in ("standard", "class"):
            raise ValueError("module_type harus standard|class")
        c = proj.VBComponents.Add(1 if module_type == "standard" else 2)
        c.Name = name
        created = True
    cm = c.CodeModule
    body = _crlf(code)
    if mode == "replace":
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(body)
    elif mode == "append":
        cm.AddFromString("\r\n" + body)
    else:
        if line is None:
            raise ValueError("mode insert butuh `line`")
        cm.InsertLines(line, body)
    if save:
        if _ext(wb.FullName) not in MACRO_EXT:
            raise RuntimeError("Workbook bukan .xlsm/.xlsb/.xls; macro tidak akan tersimpan. Gunakan excel_save_as dulu.")
        wb.Save()
    return {"module": c.Name, "created": created, "lines": cm.CountOfLines, "saved": save}


@mcp.tool()
def vba_delete_module(name: str, path: str | None = None) -> dict:
    """Hapus modul standard/class/userform (modul dokumen seperti ThisWorkbook hanya bisa dikosongkan)."""
    _need_com()
    proj = _vbproj(com.wb(path))
    c = _comp(proj, name)
    if c is None:
        raise ValueError(f"Modul '{name}' tidak ditemukan")
    if c.Type == 100:
        if c.CodeModule.CountOfLines:
            c.CodeModule.DeleteLines(1, c.CodeModule.CountOfLines)
        return {"cleared": name}
    proj.VBComponents.Remove(c)
    return {"deleted": name}


@mcp.tool()
def vba_export_module(name: str, out_path: str, path: str | None = None) -> dict:
    """Ekspor modul ke file (.bas / .cls / .frm)."""
    _need_com()
    c = _comp(_vbproj(com.wb(path)), name)
    if c is None:
        raise ValueError(f"Modul '{name}' tidak ditemukan")
    op = _p(out_path)
    c.Export(op)
    return {"exported": op}


@mcp.tool()
def vba_import_module(file_path: str, path: str | None = None) -> dict:
    """Impor file .bas / .cls / .frm ke project VBA."""
    _need_com()
    fp = _p(file_path)
    c = _vbproj(com.wb(path)).VBComponents.Import(fp)
    return {"imported": c.Name}


@mcp.tool()
def vba_search(pattern: str, regex: bool = False, path: str | None = None) -> dict:
    """Cari teks/regex di seluruh modul VBA; kembalikan modul, nomor baris, isi baris."""
    _need_com()
    rx = re.compile(pattern if regex else re.escape(pattern), re.I)
    hits = []
    for c in _vbproj(com.wb(path)).VBComponents:
        cm = c.CodeModule
        if not cm.CountOfLines:
            continue
        for i, ln in enumerate(cm.Lines(1, cm.CountOfLines).split("\r\n"), 1):
            if rx.search(ln):
                hits.append({"module": c.Name, "line": i, "text": ln.strip()})
                if len(hits) >= 500:
                    return {"matches": hits, "truncated": True}
    return {"matches": hits, "truncated": False}


@mcp.tool()
def vba_run_macro(macro: str, args: list[Any] | None = None, path: str | None = None) -> dict:
    """Jalankan Sub/Function VBA, mis. macro='Module1.Hitung' dengan args. Butuh EXCEL_MCP_ALLOW_MACROS=1.
    Hanya jalankan macro dari file yang Anda percaya."""
    _need_com()
    if not ALLOW_MACROS:
        raise PermissionError("Eksekusi macro dinonaktifkan. Set env EXCEL_MCP_ALLOW_MACROS=1 lalu restart server.")
    wb = com.wb(path)
    app = com.application()
    res = app.Run(f"'{wb.Name}'!{macro}", *(args or []))
    return {"macro": macro, "result": _clean(res) if not isinstance(res, tuple) else _grid(res)}


# =========================================================================== TOOLS: struktur & tampilan (COM)
PASTE_MODES = {"all": -4104, "values": -4163, "formulas": -4123, "formats": -4122}


def _rng(sheet: str | None, address: str, path: str | None):
    return _ws(com.wb(path), sheet).Range(address)


@mcp.tool()
def range_structure(address: str, action: str, sheet: str | None = None, path: str | None = None) -> dict:
    """Sisip/hapus/sembunyikan/tampilkan. action: insert|delete|hide|unhide.
    address: baris '5:7', kolom 'C:D', atau range 'A1:C3' (insert/delete pada range)."""
    _need_com()
    r = _rng(sheet, address, path)
    a = action.lower()
    if a == "insert":
        r.Insert()
    elif a == "delete":
        r.Delete()
    elif a in ("hide", "unhide"):
        if re.fullmatch(r"\$?\d+:\$?\d+", address):
            tgt = r.EntireRow
        elif re.fullmatch(r"\$?[A-Za-z]+:\$?[A-Za-z]+", address):
            tgt = r.EntireColumn
        else:
            raise ValueError("hide/unhide butuh alamat baris '5:7' atau kolom 'C:D'")
        tgt.Hidden = a == "hide"
    else:
        raise ValueError("action harus insert|delete|hide|unhide")
    return {"address": address, "action": a}


@mcp.tool()
def range_size(address: str, column_width: float | None = None, row_height: float | None = None,
               autofit: bool = False, sheet: str | None = None, path: str | None = None) -> dict:
    """Atur lebar kolom / tinggi baris, atau autofit."""
    _need_com()
    r = _rng(sheet, address, path)
    if column_width is not None:
        r.ColumnWidth = column_width
    if row_height is not None:
        r.RowHeight = row_height
    if autofit:
        r.Columns.AutoFit()
        r.Rows.AutoFit()
    return {"address": address}


@mcp.tool()
def range_merge(address: str, merge: bool = True, across: bool = False,
                sheet: str | None = None, path: str | None = None) -> dict:
    """Gabung (merge) atau pisahkan (unmerge) sel."""
    _need_com()
    r = _rng(sheet, address, path)
    r.Merge(across) if merge else r.UnMerge()
    return {"address": address, "merged": merge}


@mcp.tool()
def freeze_panes(cell: str | None = None, sheet: str | None = None, path: str | None = None) -> dict:
    """Bekukan baris/kolom di atas-kiri `cell` (mis. 'B2' membekukan baris 1 dan kolom A). cell kosong = lepas."""
    _need_com()
    wb = com.wb(path)
    ws = _ws(wb, sheet)
    wb.Activate()
    ws.Activate()
    w = com.application().ActiveWindow
    w.FreezePanes = False
    if cell:
        c = ws.Range(cell)
        w.ScrollRow = 1
        w.ScrollColumn = 1
        w.SplitRow = c.Row - 1
        w.SplitColumn = c.Column - 1
        w.FreezePanes = True
    return {"frozen_at": cell}


@mcp.tool()
def range_copy(src: str, dst: str, sheet: str | None = None, dst_sheet: str | None = None,
               paste: str = "all", move: bool = False, path: str | None = None) -> dict:
    """Salin atau pindahkan range. paste: all|values|formulas|formats. move=true memotong lalu menempel."""
    _need_com()
    if paste not in PASTE_MODES:
        raise ValueError(f"paste harus salah satu dari {list(PASTE_MODES)}")
    wb = com.wb(path)
    s = _ws(wb, sheet)
    d = _ws(wb, dst_sheet) if dst_sheet else s
    sr, dr = s.Range(src), d.Range(dst)
    if move:
        sr.Cut(Destination=dr)
    elif paste == "all":
        sr.Copy(Destination=dr)
    else:
        sr.Copy()
        dr.PasteSpecial(Paste=PASTE_MODES[paste])
        com.application().CutCopyMode = False
    return {"from": src, "to": dst, "mode": "move" if move else paste}


@mcp.tool()
def range_replace(find: str, replace: str, sheet: str | None = None, path: str | None = None,
                  match_case: bool = False, whole_cell: bool = False) -> dict:
    """Cari dan ganti teks di seluruh used range sheet."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    ws.UsedRange.Replace(What=find, Replacement=replace, LookAt=1 if whole_cell else 2, MatchCase=match_case)
    return {"sheet": ws.Name, "replaced": find, "with": replace}


@mcp.tool()
def sheet_copy(name: str, new_name: str | None = None, after: str | None = None, path: str | None = None) -> dict:
    """Duplikasi sheet."""
    _need_com()
    wb = com.wb(path)
    ws = _ws(wb, name)
    anchor = _ws(wb, after) if after else wb.Worksheets(wb.Worksheets.Count)
    ws.Copy(After=anchor)
    new = wb.ActiveSheet
    if new_name:
        new.Name = new_name
    return {"copied": new.Name}


@mcp.tool()
def sheet_move(name: str, after: str | None = None, before: str | None = None, path: str | None = None) -> dict:
    """Pindahkan urutan sheet."""
    _need_com()
    wb = com.wb(path)
    ws = _ws(wb, name)
    if before:
        ws.Move(Before=_ws(wb, before))
    elif after:
        ws.Move(After=_ws(wb, after))
    else:
        raise ValueError("isi after atau before")
    return {"moved": name}


@mcp.tool()
def sheet_visibility(name: str, state: str, path: str | None = None) -> dict:
    """state: visible | hidden | veryhidden."""
    _need_com()
    m = {"visible": -1, "hidden": 0, "veryhidden": 2}
    if state not in m:
        raise ValueError("state harus visible|hidden|veryhidden")
    _ws(com.wb(path), name).Visible = m[state]
    return {"sheet": name, "state": state}


@mcp.tool()
def protect_sheet(name: str, password: str | None = None, unprotect: bool = False, path: str | None = None) -> dict:
    """Kunci atau buka kunci sheet."""
    _need_com()
    ws = _ws(com.wb(path), name)
    if unprotect:
        ws.Unprotect(password) if password else ws.Unprotect()
    else:
        ws.Protect(Password=password) if password else ws.Protect()
    return {"sheet": name, "protected": not unprotect}


@mcp.tool()
def protect_workbook(password: str | None = None, unprotect: bool = False, path: str | None = None) -> dict:
    """Kunci/buka struktur workbook (tambah/hapus/pindah sheet)."""
    _need_com()
    wb = com.wb(path)
    if unprotect:
        wb.Unprotect(password) if password else wb.Unprotect()
    else:
        wb.Protect(Password=password, Structure=True) if password else wb.Protect(Structure=True)
    return {"structure_protected": not unprotect}


@mcp.tool()
def page_setup(sheet: str | None = None, path: str | None = None, orientation: str | None = None,
               print_area: str | None = None, fit_wide: int | None = None, fit_tall: int | None = None,
               center_header: str | None = None, center_footer: str | None = None) -> dict:
    """Pengaturan cetak. orientation: portrait|landscape. print_area '' = hapus. fit_wide/fit_tall = jumlah halaman."""
    _need_com()
    ps = _ws(com.wb(path), sheet).PageSetup
    if orientation:
        ps.Orientation = 2 if orientation.lower() == "landscape" else 1
    if print_area is not None:
        ps.PrintArea = print_area
    if fit_wide is not None or fit_tall is not None:
        ps.Zoom = False
        ps.FitToPagesWide = fit_wide if fit_wide else False
        ps.FitToPagesTall = fit_tall if fit_tall else False
    if center_header is not None:
        ps.CenterHeader = center_header
    if center_footer is not None:
        ps.CenterFooter = center_footer
    return {"page_setup": "ok"}


@mcp.tool()
def calculation_mode(mode: str) -> dict:
    """mode: auto | manual | semiauto."""
    _need_com()
    m = {"auto": -4105, "manual": -4135, "semiauto": 2}
    if mode not in m:
        raise ValueError("mode harus auto|manual|semiauto")
    com.application().Calculation = m[mode]
    return {"calculation": mode}


# =========================================================================== TOOLS: data (COM)
@mcp.tool()
def range_sort(address: str, keys: list[dict], has_header: bool = True,
               sheet: str | None = None, path: str | None = None) -> dict:
    """Urutkan range. keys: [{"column":"B","order":"asc"|"desc"}, ...] (huruf kolom di dalam range)."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    rng = ws.Range(address)
    first, last = rng.Row, rng.Row + rng.Rows.Count - 1
    srt = ws.Sort
    srt.SortFields.Clear()
    for k in keys:
        col = k["column"]
        srt.SortFields.Add(Key=ws.Range(f"{col}{first}:{col}{last}"), SortOn=0,
                           Order=2 if str(k.get("order", "asc")).lower() == "desc" else 1, DataOption=0)
    srt.SetRange(rng)
    srt.Header = 1 if has_header else 2
    srt.Apply()
    return {"sorted": address, "keys": keys}


@mcp.tool()
def autofilter(address: str | None = None, field: int | None = None, criteria: str | None = None,
               clear: bool = False, sheet: str | None = None, path: str | None = None) -> dict:
    """Pasang AutoFilter pada range (field = nomor kolom 1-based, criteria mis. '>100' atau 'Jakarta'). clear=true menghapus filter."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    if clear:
        ws.AutoFilterMode = False
        return {"filter": "cleared"}
    if not address:
        raise ValueError("address wajib")
    rng = ws.Range(address)
    if field is None:
        rng.AutoFilter()
    else:
        rng.AutoFilter(Field=field, Criteria1=criteria)
    return {"filter": address, "field": field, "criteria": criteria}


@mcp.tool()
def remove_duplicates(address: str, columns: list[int], has_header: bool = True,
                      sheet: str | None = None, path: str | None = None) -> dict:
    """Hapus baris duplikat berdasarkan kolom (nomor 1-based di dalam range)."""
    _need_com()
    rng = _rng(sheet, address, path)
    rng.RemoveDuplicates(Columns=tuple(columns), Header=1 if has_header else 2)
    return {"deduplicated": address}


@mcp.tool()
def data_validation(address: str, kind: str, formula1: str | None = None, formula2: str | None = None,
                    operator: str = "between", error_message: str | None = None, clear: bool = False,
                    sheet: str | None = None, path: str | None = None) -> dict:
    """Validasi data. kind: list|whole|decimal|date|time|textlength|custom.
    list: formula1='Ya,Tidak' atau '=Sheet1!$A$1:$A$5'. operator: between|notbetween|equal|notequal|greater|less|gte|lte."""
    _need_com()
    rng = _rng(sheet, address, path)
    v = rng.Validation
    v.Delete()
    if clear:
        return {"validation": "cleared"}
    kinds = {"whole": 1, "decimal": 2, "list": 3, "date": 4, "time": 5, "textlength": 6, "custom": 7}
    ops = {"between": 1, "notbetween": 2, "equal": 3, "notequal": 4, "greater": 5, "less": 6, "gte": 7, "lte": 8}
    if kind not in kinds:
        raise ValueError(f"kind harus salah satu dari {list(kinds)}")
    args: dict[str, Any] = {"Type": kinds[kind], "AlertStyle": 1}
    if kind not in ("list", "custom"):
        args["Operator"] = ops[operator]
    if formula1 is not None:
        args["Formula1"] = formula1
    if formula2 is not None:
        args["Formula2"] = formula2
    v.Add(**args)
    if error_message:
        v.ErrorMessage = error_message
    return {"validation": kind, "address": address}


@mcp.tool()
def conditional_format(address: str, kind: str, formula: str | None = None, operator: str = "greater",
                       value2: str | None = None, fill_color: str | None = None, font_color: str | None = None,
                       bold: bool | None = None, clear: bool = False,
                       sheet: str | None = None, path: str | None = None) -> dict:
    """Conditional formatting. kind: cell_value (pakai operator + formula sebagai nilai, mis. '100') |
    expression (formula mis. '=$B2>100') | duplicates | color_scale | data_bar. clear=true menghapus semua aturan di range."""
    _need_com()
    rng = _rng(sheet, address, path)
    if clear:
        rng.FormatConditions.Delete()
        return {"conditional_format": "cleared"}
    ops = {"between": 1, "notbetween": 2, "equal": 3, "notequal": 4, "greater": 5, "less": 6, "gte": 7, "lte": 8}
    if kind == "color_scale":
        rng.FormatConditions.AddColorScale(ColorScaleType=3)
        return {"conditional_format": kind}
    if kind == "data_bar":
        rng.FormatConditions.AddDatabar()
        return {"conditional_format": kind}
    if kind == "cell_value":
        kw: dict[str, Any] = {"Type": 1, "Operator": ops[operator], "Formula1": formula}
        if value2 is not None:
            kw["Formula2"] = value2
        fc = rng.FormatConditions.Add(**kw)
    elif kind == "expression":
        fc = rng.FormatConditions.Add(Type=2, Formula1=formula)
    elif kind == "duplicates":
        fc = rng.FormatConditions.Add(Type=8)
        fc.DupeUnique = 0
    else:
        raise ValueError("kind harus cell_value|expression|duplicates|color_scale|data_bar")
    if fill_color:
        fc.Interior.Color = _rgb(fill_color)
    if font_color:
        fc.Font.Color = _rgb(font_color)
    if bold is not None:
        fc.Font.Bold = bold
    return {"conditional_format": kind, "address": address}


@mcp.tool()
def csv_import(csv_path: str, sheet: str | None = None, start: str = "A1", path: str | None = None,
               delimiter: str = ",", encoding: str = "utf-8-sig") -> dict:
    """Impor file CSV ke sheet mulai dari `start`."""
    import csv

    with open(_p(csv_path), newline="", encoding=encoding) as f:
        rows = [list(r) for r in csv.reader(f, delimiter=delimiter)]
    if not rows:
        return {"imported_rows": 0}
    res = range_write(rows, sheet=sheet, start=start, path=path)
    res["imported_rows"] = len(rows)
    return res


# =========================================================================== TOOLS: tabel, chart, pivot, query (COM)
@mcp.tool()
def table_list(path: str | None = None) -> dict:
    """Daftar Excel Table di seluruh sheet."""
    _need_com()
    out = []
    for ws in com.wb(path).Worksheets:
        for lo in ws.ListObjects:
            out.append({"sheet": ws.Name, "name": lo.Name, "range": lo.Range.GetAddress(False, False)})
    return {"tables": out}


@mcp.tool()
def table_edit(name: str, resize: str | None = None, style: str | None = None,
               convert_to_range: bool = False, path: str | None = None) -> dict:
    """Ubah ukuran/gaya tabel, atau ubah kembali menjadi range biasa."""
    _need_com()
    for ws in com.wb(path).Worksheets:
        for lo in ws.ListObjects:
            if lo.Name.lower() == name.lower():
                if resize:
                    lo.Resize(ws.Range(resize))
                if style:
                    lo.TableStyle = style
                if convert_to_range:
                    lo.Unlist()
                return {"table": name, "range": None if convert_to_range else lo.Range.GetAddress(False, False)}
    raise ValueError(f"Tabel '{name}' tidak ditemukan")


@mcp.tool()
def chart_list(sheet: str | None = None, path: str | None = None) -> dict:
    """Daftar chart pada sheet."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    return {"charts": [{"name": c.Name, "title": c.Chart.ChartTitle.Text if c.Chart.HasTitle else None}
                       for c in ws.ChartObjects()]}


@mcp.tool()
def chart_edit(name: str, title: str | None = None, chart_type: str | None = None,
               delete: bool = False, sheet: str | None = None, path: str | None = None) -> dict:
    """Ubah judul/tipe chart, atau hapus."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    co = ws.ChartObjects(name)
    if delete:
        co.Delete()
        return {"deleted": name}
    if title is not None:
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = title
    if chart_type:
        if chart_type not in CHART_TYPES:
            raise ValueError(f"chart_type harus salah satu dari {list(CHART_TYPES)}")
        co.Chart.ChartType = CHART_TYPES[chart_type]
    return {"chart": name}


@mcp.tool()
def pivot_add(source: str, dest_cell: str, name: str, rows: list[str] | None = None,
              columns: list[str] | None = None, values: list[dict] | None = None,
              filters: list[str] | None = None, sheet: str | None = None,
              dest_sheet: str | None = None, path: str | None = None) -> dict:
    """Buat PivotTable. source: range termasuk header (mis. 'A1:E500'). dest_cell di dest_sheet (default sheet yang sama).
    values: [{"field":"Penjualan","function":"sum|count|average|max|min","caption":"Total"}]. Nama field = teks header."""
    _need_com()
    wb = com.wb(path)
    s = _ws(wb, sheet)
    d = _ws(wb, dest_sheet) if dest_sheet else s
    src = f"'{s.Name}'!{s.Range(source).GetAddress(True, True, 1)}"
    cache = wb.PivotCaches().Create(SourceType=1, SourceData=src)
    pt = cache.CreatePivotTable(TableDestination=d.Range(dest_cell), TableName=name)
    for f in rows or []:
        pt.PivotFields(f).Orientation = 1
    for f in columns or []:
        pt.PivotFields(f).Orientation = 2
    for f in filters or []:
        pt.PivotFields(f).Orientation = 3
    fn = {"sum": -4157, "count": -4112, "average": -4106, "max": -4136, "min": -4139}
    for v in values or []:
        func = v.get("function", "sum").lower()
        if func not in fn:
            raise ValueError(f"function harus salah satu dari {list(fn)}")
        pt.AddDataField(pt.PivotFields(v["field"]), v.get("caption", f"{func} {v['field']}"), fn[func])
    return {"pivot": name, "range": pt.TableRange2.GetAddress(False, False)}


@mcp.tool()
def pivot_refresh(name: str | None = None, path: str | None = None) -> dict:
    """Segarkan satu PivotTable (nama) atau semuanya."""
    _need_com()
    n = 0
    for ws in com.wb(path).Worksheets:
        for pt in ws.PivotTables():
            if name is None or pt.Name.lower() == name.lower():
                pt.RefreshTable()
                n += 1
    return {"refreshed": n}


@mcp.tool()
def queries_list(path: str | None = None) -> dict:
    """Daftar Power Query beserta kode M-nya (Excel 2016+)."""
    _need_com()
    return {"queries": [{"name": q.Name, "formula": q.Formula} for q in com.wb(path).Queries]}


@mcp.tool()
def query_add(name: str, m_code: str, load_sheet: str | None = None, load_cell: str = "A1",
              path: str | None = None) -> dict:
    """Tambah Power Query (kode M). Jika load_sheet diisi, hasilnya dimuat ke tabel di sheet tersebut."""
    _need_com()
    wb = com.wb(path)
    for q in wb.Queries:
        if q.Name.lower() == name.lower():
            raise ValueError(f"Query '{name}' sudah ada")
    wb.Queries.Add(Name=name, Formula=m_code)
    if load_sheet:
        ws = _ws(wb, load_sheet)
        conn = f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""'
        lo = ws.ListObjects.Add(SourceType=0, Source=conn, Destination=ws.Range(load_cell))
        lo.QueryTable.CommandType = 2
        lo.QueryTable.CommandText = f"SELECT * FROM [{name}]"
        lo.QueryTable.Refresh(False)
    return {"query": name, "loaded_to": load_sheet}


@mcp.tool()
def refresh_all(path: str | None = None) -> dict:
    """Segarkan semua koneksi data, query, dan pivot."""
    _need_com()
    com.wb(path).RefreshAll()
    com.application().CalculateUntilAsyncQueriesDone()
    return {"refreshed": True}


# =========================================================================== TOOLS: konten (COM)
@mcp.tool()
def names_list(path: str | None = None) -> dict:
    """Daftar named range."""
    _need_com()
    return {"names": [{"name": n.Name, "refers_to": n.RefersTo} for n in com.wb(path).Names]}


@mcp.tool()
def name_delete(name: str, path: str | None = None) -> dict:
    """Hapus named range."""
    _need_com()
    com.wb(path).Names(name).Delete()
    return {"deleted": name}


@mcp.tool()
def comment_set(cell: str, text: str | None = None, sheet: str | None = None, path: str | None = None) -> dict:
    """Setel catatan (note) pada sel; text kosong menghapus."""
    _need_com()
    r = _rng(sheet, cell, path)
    r.ClearComments()
    if text:
        r.AddComment(text)
    return {"cell": cell, "comment": bool(text)}


@mcp.tool()
def hyperlink_add(cell: str, url: str, text: str | None = None, sheet: str | None = None,
                  path: str | None = None) -> dict:
    """Tambah hyperlink (URL, atau lokasi internal seperti 'Sheet2!A1' dengan url diawali '#')."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    if url.startswith("#"):
        ws.Hyperlinks.Add(Anchor=ws.Range(cell), Address="", SubAddress=url[1:], TextToDisplay=text or url[1:])
    else:
        ws.Hyperlinks.Add(Anchor=ws.Range(cell), Address=url, TextToDisplay=text or url)
    return {"cell": cell, "url": url}


@mcp.tool()
def image_add(file_path: str, cell: str, width: float | None = None, height: float | None = None,
              sheet: str | None = None, path: str | None = None) -> dict:
    """Sisipkan gambar di posisi sel."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    c = ws.Range(cell)
    shp = ws.Shapes.AddPicture(_p(file_path), 0, -1, c.Left, c.Top, width or -1, height or -1)
    return {"shape": shp.Name}


@mcp.tool()
def button_add(cell: str, caption: str, macro: str, width: float = 100, height: float = 24,
               sheet: str | None = None, path: str | None = None) -> dict:
    """Tambah tombol di sheet yang menjalankan macro VBA (mis. macro='Module1.Hitung')."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    c = ws.Range(cell)
    b = ws.Buttons().Add(c.Left, c.Top, width, height)
    b.Caption = caption
    b.OnAction = macro
    return {"button": b.Name, "macro": macro}


@mcp.tool()
def shape_list(sheet: str | None = None, path: str | None = None) -> dict:
    """Daftar shape/gambar/tombol/chart pada sheet."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    return {"shapes": [{"name": s.Name, "type": s.Type, "left": s.Left, "top": s.Top,
                        "width": s.Width, "height": s.Height} for s in ws.Shapes]}


@mcp.tool()
def shape_delete(name: str, sheet: str | None = None, path: str | None = None) -> dict:
    """Hapus shape berdasarkan nama."""
    _need_com()
    _ws(com.wb(path), sheet).Shapes(name).Delete()
    return {"deleted": name}


# =========================================================================== TOOLS: VBA tambahan (COM)
@mcp.tool()
def vba_references(action: str = "list", guid: str | None = None, major: int = 0, minor: int = 0,
                   file_path: str | None = None, name: str | None = None, path: str | None = None) -> dict:
    """Kelola reference library VBA. action: list | add (guid+major+minor atau file_path) | remove (name)."""
    _need_com()
    if action in ("add", "remove"):
        _check_write("vba_references", vba=True)
    refs = _vbproj(com.wb(path)).References
    if action == "add":
        if guid:
            refs.AddFromGuid(guid, major, minor)
        elif file_path:
            refs.AddFromFile(_p(file_path))
        else:
            raise ValueError("add butuh guid atau file_path")
    elif action == "remove":
        for r in refs:
            if r.Name.lower() == (name or "").lower():
                refs.Remove(r)
                break
        else:
            raise ValueError(f"Reference '{name}' tidak ditemukan")
    elif action != "list":
        raise ValueError("action harus list|add|remove")
    return {"references": [{"name": r.Name, "guid": r.Guid, "description": r.Description,
                            "builtin": r.BuiltIn} for r in refs]}


@mcp.tool()
def vba_userform_create(name: str, caption: str | None = None, width: float | None = None,
                        height: float | None = None, path: str | None = None) -> dict:
    """Buat UserForm baru."""
    _need_com()
    c = _vbproj(com.wb(path)).VBComponents.Add(3)
    c.Name = name
    if caption:
        c.Properties("Caption").Value = caption
    if width:
        c.Properties("Width").Value = width
    if height:
        c.Properties("Height").Value = height
    return {"userform": c.Name}


@mcp.tool()
def vba_userform_add_control(form: str, control: str, name: str, left: float | None = None, top: float | None = None,
                             width: float = 80, height: float = 24, caption: str | None = None,
                             path: str | None = None) -> dict:
    """Tambah kontrol ke UserForm. control: CommandButton|Label|TextBox|ComboBox|ListBox|CheckBox|OptionButton|Frame|Image."""
    _need_com()
    c = _comp(_vbproj(com.wb(path)), form)
    if c is None:
        raise ValueError(f"UserForm '{form}' tidak ditemukan")
    if left is None or top is None:
        # Tata otomatis berbaris 2 kolom, setiap baris = 1 kontrol
        existing = list(c.Designer.Controls)
        idx = len(existing)
        col, row = idx % 2, idx // 2
        left = (10 + col * (width + 10)) if left is None else left
        top = (10 + row * (height + 8)) if top is None else top
    ctl = c.Designer.Controls.Add(f"Forms.{control}.1", name, True)
    ctl.Left, ctl.Top, ctl.Width, ctl.Height = left, top, width, height
    if caption is not None:
        ctl.Caption = caption
    return {"form": form, "control": name}


@mcp.tool()
def vba_userform_controls(form: str, path: str | None = None) -> dict:
    """Daftar kontrol pada UserForm."""
    _need_com()
    c = _comp(_vbproj(com.wb(path)), form)
    if c is None:
        raise ValueError(f"UserForm '{form}' tidak ditemukan")
    return {"controls": [{"name": x.Name, "left": x.Left, "top": x.Top, "width": x.Width, "height": x.Height}
                         for x in c.Designer.Controls]}


# =========================================================================== TOOL: akses penuh (COM)
@mcp.tool()
def excel_eval(code: str, path: str | None = None, sheet: str | None = None) -> dict:
    """AKSES PENUH ke object model Excel lewat Python (pywin32). Variabel tersedia: app (Excel.Application),
    wb (workbook), ws (sheet), os. Simpan keluaran di variabel `result`. Untuk apa pun yang belum punya tool khusus.
    Mati secara default; butuh EXCEL_MCP_ALLOW_EVAL=1 (kode berjalan dengan hak akses Anda)."""
    _need_com()
    if os.environ.get("EXCEL_MCP_ALLOW_EVAL", "0") != "1":
        raise PermissionError("excel_eval dinonaktifkan. Set env EXCEL_MCP_ALLOW_EVAL=1 lalu restart server.")
    wb = com.wb(path)
    ns: dict[str, Any] = {"app": com.application(), "wb": wb, "ws": _ws(wb, sheet), "os": os, "result": None}
    exec(code, ns)
    r = ns.get("result")
    if isinstance(r, tuple):
        r = _grid(r)
    else:
        r = _clean(r)
    try:
        json.dumps(r)
    except TypeError:
        r = str(r)
    return {"result": r}


# =========================================================================== TOOLS: ringkasan file
def _profile(grid: list[list[Any]], header: bool = True) -> list[dict]:
    if not grid or not grid[0]:
        return []
    heads = grid[0] if header else [None] * len(grid[0])
    body = grid[1:] if header else grid
    err = set(ERRORS.values())
    out = []
    for j in range(len(grid[0])):
        vals = [r[j] for r in body if j < len(r)]
        filled = [v for v in vals if v not in (None, "")]
        kinds: dict[str, int] = {}
        for v in filled:
            if isinstance(v, bool):
                k = "bool"
            elif isinstance(v, (int, float)):
                k = "number"
            elif isinstance(v, str) and v in err:
                k = "error"
            elif isinstance(v, str) and re.match(r"^\d{4}-\d{2}-\d{2}T", v):
                k = "date"
            else:
                k = "text"
            kinds[k] = kinds.get(k, 0) + 1
        out.append({"column": j + 1, "header": heads[j], "filled": len(filled),
                    "empty": len(vals) - len(filled), "types": kinds, "example": filled[0] if filled else None})
    return out


@mcp.tool()
def describe_workbook(path: str | None = None) -> dict:
    """Ringkasan cepat workbook: format, sheet (ukuran, baris header, tabel, chart, pivot), named range, ada/tidaknya VBA."""
    if not USE_COM:
        if not path:
            raise ValueError("path wajib pada backend FILE")
        ap = _p(path)
        sheets = []
        for n in _file_sheets(ap):
            first = _file_read(ap, n, "A1:AD1", False).get("values", [[]])
            hdr = list(first[0]) if first else []
            while hdr and hdr[-1] is None:
                hdr.pop()
            sheets.append({"name": n, "header_row": hdr})
        try:
            has_vba = bool(_vba_from_file(ap)) if _ext(ap) in MACRO_EXT else False
        except Exception:
            has_vba = None
        return {"path": ap, "format": _ext(ap), "sheets": sheets, "has_vba": has_vba}
    wb = com.wb(path)
    sheets = []
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        ncols = ur.Columns.Count
        first = _grid(_range_at(ws, ur.Row, ur.Column, 1, min(ncols, 30)).Value)[0]
        sheets.append({"name": ws.Name, "visible": ws.Visible == -1, "used_range": ur.GetAddress(False, False),
                       "rows": ur.Rows.Count, "columns": ncols, "header_row": first,
                       "tables": [lo.Name for lo in ws.ListObjects], "charts": ws.ChartObjects().Count,
                       "pivots": ws.PivotTables().Count, "protected": bool(ws.ProtectContents)})
    info: dict[str, Any] = {"path": wb.FullName, "format": _ext(wb.FullName), "sheets": sheets,
                            "named_ranges": wb.Names.Count, "has_vba": _has_vba(wb)}
    try:
        info["power_queries"] = wb.Queries.Count
    except Exception:
        pass
    return info


@mcp.tool()
def describe_sheet(sheet: str | None = None, path: str | None = None, sample_rows: int = 5,
                   profile_rows: int = 1000, max_columns: int = 30) -> dict:
    """Ringkasan satu sheet: ukuran data, profil tiap kolom (header, terisi/kosong, tipe, contoh), sampel baris,
    jumlah sel berformula, tabel, chart, pivot. Profil dihitung dari `profile_rows` baris pertama."""
    if not USE_COM:
        if not path:
            raise ValueError("path wajib pada backend FILE")
        res = _file_read(path, sheet, None, False, 0, profile_rows)
        grid = [r[:max_columns] for r in res.get("values", [])]
        return {"sheet": res.get("sheet"), "total_rows": res.get("total_rows"),
                "columns_profile": _profile(grid), "sample": grid[:1 + sample_rows]}
    ws = _ws(com.wb(path), sheet)
    ur = ws.UsedRange
    rows, cols = ur.Rows.Count, ur.Columns.Count
    n_rows, n_cols = min(rows, profile_rows), min(cols, max_columns)
    grid = _grid(_range_at(ws, ur.Row, ur.Column, n_rows, n_cols).Value)
    try:
        formulas = ur.SpecialCells(-4123).Count
    except Exception:
        formulas = 0
    return {"sheet": ws.Name, "used_range": ur.GetAddress(False, False), "rows": rows, "columns": cols,
            "formula_cells": formulas, "has_merged_cells": ur.MergeCells is not False,
            "profiled_rows": n_rows, "profiled_columns": n_cols, "columns_profile": _profile(grid),
            "sample": grid[:1 + sample_rows], "tables": [lo.Name for lo in ws.ListObjects],
            "charts": ws.ChartObjects().Count, "pivots": ws.PivotTables().Count,
            "shapes": ws.Shapes.Count, "protected": bool(ws.ProtectContents)}


# =========================================================================== TOOLS: formula (COM)
def _areas(rng) -> list[str]:
    return [a.GetAddress(False, False) for a in rng.Areas]


@mcp.tool()
def trace_formula(cell: str, sheet: str | None = None, path: str | None = None) -> dict:
    """Telusuri formula sebuah sel: isi formula, precedents (sel yang dipakai) dan dependents (sel yang memakainya).
    Keterbatasan Excel: hanya mencakup sel pada sheet yang sama."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    c = ws.Range(cell)
    out: dict[str, Any] = {"cell": cell, "formula": c.Formula, "value": _clean(c.Value),
                           "precedents": [], "dependents": []}
    try:
        out["precedents"] = _areas(c.DirectPrecedents)
    except Exception:
        pass
    try:
        out["dependents"] = _areas(c.DirectDependents)
    except Exception:
        pass
    return out


@mcp.tool()
def goal_seek(target_cell: str, goal: float, changing_cell: str, sheet: str | None = None,
              path: str | None = None) -> dict:
    """Goal Seek: ubah `changing_cell` agar formula di `target_cell` bernilai `goal` (keduanya harus di sheet yang sama).
    Mengubah nilai di workbook."""
    _need_com()
    ws = _ws(com.wb(path), sheet)
    t, ch = ws.Range(target_cell), ws.Range(changing_cell)
    ok = bool(t.GoalSeek(Goal=goal, ChangingCell=ch))
    return {"success": ok, "changing_cell_value": _clean(ch.Value), "target_cell_value": _clean(t.Value)}


# =========================================================================== TOOLS: slicer & sparklines (COM)
def _find_slicer_source(wb, name):
    """Cari tabel/pivot yang bernama `name` di seluruh workbook."""
    for s in wb.Worksheets:
        for lo in s.ListObjects:
            if lo.Name.lower() == name.lower():
                return lo
    for s in wb.Worksheets:
        for pt in s.PivotTables():
            if pt.Name.lower() == name.lower():
                return pt
    return None


@mcp.tool()
def slicer_add(source: str, field: str, dest_cell: str, caption: str | None = None, timeline: bool = False,
               width: float = 150, height: float = 150, sheet: str | None = None, path: str | None = None) -> dict:
    """Tambah slicer (atau timeline untuk field tanggal) untuk Excel Table atau PivotTable bernama `source`."""
    _need_com()
    wb = com.wb(path)
    ws = _ws(wb, sheet)
    src = _find_slicer_source(wb, source)
    if src is None:
        raise ValueError(f"Tabel/PivotTable '{source}' tidak ditemukan")
    kw: dict[str, Any] = {"Source": src, "SourceField": field}
    if timeline:
        kw["SlicerCacheType"] = 1
    cache = wb.SlicerCaches.Add2(**kw)
    c = ws.Range(dest_cell)
    sl = cache.Slicers.Add(SlicerDestination=ws, Caption=caption or field,
                           Top=c.Top, Left=c.Left, Width=width, Height=height)
    return {"slicer": sl.Name, "cache": cache.Name}


@mcp.tool()
def slicers_list(path: str | None = None) -> dict:
    """Daftar slicer/timeline di workbook."""
    _need_com()
    out = []
    for cache in com.wb(path).SlicerCaches:
        for sl in cache.Slicers:
            out.append({"name": sl.Name, "caption": sl.Caption, "cache": cache.Name})
    return {"slicers": out}


@mcp.tool()
def slicer_delete(name: str, path: str | None = None) -> dict:
    """Hapus slicer/timeline berdasarkan nama (beserta cache-nya)."""
    _need_com()
    for cache in com.wb(path).SlicerCaches:
        for sl in cache.Slicers:
            if sl.Name.lower() == name.lower():
                cache.Delete()
                return {"deleted": name}
    raise ValueError(f"Slicer '{name}' tidak ditemukan")


@mcp.tool()
def sparklines_add(data_range: str, location_range: str, kind: str = "line", sheet: str | None = None,
                   path: str | None = None) -> dict:
    """Tambah sparklines. kind: line|column|winloss. location_range = sel tujuan, data_range = data sumber."""
    _need_com()
    k = {"line": 1, "column": 2, "winloss": 3}
    if kind not in k:
        raise ValueError("kind harus line|column|winloss")
    ws = _ws(com.wb(path), sheet)
    ws.Range(location_range).SparklineGroups.Add(Type=k[kind], SourceData=f"'{ws.Name}'!{data_range}")
    return {"sparklines": location_range, "kind": kind}


@mcp.tool()
def sparklines_delete(location_range: str, sheet: str | None = None, path: str | None = None) -> dict:
    """Hapus sparklines pada range tujuan."""
    _need_com()
    _ws(com.wb(path), sheet).Range(location_range).SparklineGroups.Clear()
    return {"cleared": location_range}


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
