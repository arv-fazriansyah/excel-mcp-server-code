"""Workbook session management."""
from __future__ import annotations

import os
import time
import uuid
from dataclasses import dataclass, field
from typing import Any

from .config import XL_FORMATS, ROOT, VISIBLE
from .utils import abspath, check_path
from .com_executor import get_excel_app, run_com


@dataclass
class Session:
    id: str
    workbook: Any
    app: Any
    path: str
    read_only: bool
    created_at: float = field(default_factory=time.time)
    dirty: bool = False


_sessions: dict[str, Session] = {}


def session_open(path: str, read_only: bool = False, password: str | None = None) -> dict:
    check_path(path)
    app = get_excel_app(VISIBLE)
    full = abspath(path)
    def _do():
        return app.Workbooks.Open(full, 0, bool(read_only), None, password)
    wb = run_com(_do)
    sid = str(uuid.uuid4())
    _sessions[sid] = Session(id=sid, workbook=wb, app=app, path=full, read_only=bool(wb.ReadOnly))
    return {"sessionId": sid, "path": full, "readOnly": bool(wb.ReadOnly), "hasVba": bool(wb.HasVBProject)}


def session_create(path: str, sheets: str = "", force: bool = False) -> dict:
    check_path(path)
    full = abspath(path)
    if os.path.exists(full) and not force:
        raise FileExistsError(f"File sudah ada: {full} (pakai force=true)")
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    if os.path.exists(full):
        os.remove(full)
    ext = os.path.splitext(full)[1].lower()
    if ext not in XL_FORMATS:
        raise ValueError(f"Ekstensi tidak didukung: {ext}")
    app = get_excel_app(VISIBLE)
    names = [n.strip() for n in sheets.split(",") if n.strip()]
    def _create():
        wb = app.Workbooks.Add()
        while wb.Worksheets.Count > max(1, len(names)):
            wb.Worksheets(wb.Worksheets.Count).Delete()
        for i, name in enumerate(names):
            if i < wb.Worksheets.Count:
                wb.Worksheets(i + 1).Name = name
            else:
                ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
                ws.Name = name
        wb.SaveAs(full, XL_FORMATS[ext])
        return wb
    wb = run_com(_create)
    sid = str(uuid.uuid4())
    _sessions[sid] = Session(id=sid, workbook=wb, app=app, path=full, read_only=False)
    return {"sessionId": sid, "path": full, "sheets": names or ["Sheet1"]}


def session_attach(workbook_id: str | None = None) -> dict:
    app = get_excel_app(VISIBLE)
    def _list():
        out = []
        for wb in app.Workbooks:
            out.append({"name": wb.Name, "path": wb.FullName, "readOnly": bool(wb.ReadOnly)})
        return out
    return {"workbooks": run_com(_list)}


def _get(sid: str) -> Session:
    if sid not in _sessions:
        raise KeyError(f"Session tidak ditemukan: {sid}")
    return _sessions[sid]


def session_info(sid: str) -> dict:
    s = _get(sid)
    wb = s.workbook
    def _info():
        sheets = []
        for ws in wb.Worksheets:
            used = ws.UsedRange
            sheets.append({"name": ws.Name, "index": ws.Index, "visible": int(ws.Visible),
                           "usedRange": used.Address.replace("$", ""), "rows": used.Rows.Count,
                           "cols": used.Columns.Count})
        return {"path": s.path, "readOnly": bool(wb.ReadOnly), "format": int(wb.FileFormat),
                "hasVba": bool(wb.HasVBProject), "sheets": sheets,
                "namedRanges": [n.Name for n in wb.Names]}
    return run_com(_info)


def session_save(sid: str) -> dict:
    s = _get(sid)
    run_com(lambda: s.workbook.Save())
    s.dirty = False
    return {"ok": True, "path": s.path}


def session_save_as(sid: str, target_path: str, fmt: str = "auto", overwrite: bool = False) -> dict:
    check_path(target_path)
    s = _get(sid)
    full = abspath(target_path)
    if os.path.exists(full) and not overwrite:
        raise FileExistsError(f"File sudah ada: {full}")
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    ext = os.path.splitext(full)[1].lower()
    if fmt == "auto":
        if ext not in XL_FORMATS:
            raise ValueError(f"Format tidak didukung: {ext}")
        file_format = XL_FORMATS[ext]
    else:
        file_format = XL_FORMATS.get("." + fmt.lower().lstrip("."), XL_FORMATS.get(ext))
    if file_format is None:
        raise ValueError(f"Format tidak dikenal: {fmt}")
    run_com(lambda: s.workbook.SaveAs(full, file_format))
    s.path = full
    return {"ok": True, "path": full}


def session_close(sid: str, save: bool = False) -> dict:
    s = _get(sid)
    if save:
        run_com(lambda: s.workbook.Save())
    run_com(lambda: s.workbook.Close(bool(save)))
    del _sessions[sid]
    return {"ok": True, "saved": save}


def session_list() -> dict:
    return {"sessions": [{"id": sid, "path": s.path, "readOnly": s.read_only} for sid, s in _sessions.items()]}


def get_session(sid: str) -> Session:
    return _get(sid)
