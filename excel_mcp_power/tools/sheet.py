"""Worksheet operations."""
from __future__ import annotations

from ..session import get_session
from ..com_executor import get_executor


def sheet_list(sid: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        out = []
        for ws in wb.Worksheets:
            used = ws.UsedRange
            out.append({"name": ws.Name, "index": ws.Index, "visible": int(ws.Visible),
                        "usedRange": used.Address.replace("$", ""), "rows": used.Rows.Count,
                        "cols": used.Columns.Count})
        return {"sheets": out}
    return ex.call(_run)


def sheet_add(sid: str, name: str | None = None, after: str | None = None) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        if after:
            try:
                anchor = wb.Worksheets(after)
                ws = wb.Worksheets.Add(None, anchor)
            except Exception:
                ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
        else:
            ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
        if name:
            ws.Name = name
        return {"name": ws.Name, "index": ws.Index}
    return ex.call(_run)


def sheet_rename(sid: str, old_name: str, new_name: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        wb.Worksheets(old_name).Name = new_name
        return {"from": old_name, "to": new_name}
    return ex.call(_run)


def sheet_delete(sid: str, name: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        if wb.Worksheets.Count <= 1:
            raise ValueError("Tidak bisa menghapus satu-satunya sheet")
        wb.Worksheets(name).Delete()
        return {"deleted": name}
    return ex.call(_run)


def sheet_copy(sid: str, source: str, target_workbook: str | None = None, new_name: str | None = None) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        ws = wb.Worksheets(source)
        if target_workbook:
            dst = ex.call(lambda: None)  # placeholder not used
        ws.Copy(None, wb.Worksheets(wb.Worksheets.Count))
        new_ws = wb.Worksheets(wb.Worksheets.Count)
        if new_name:
            new_ws.Name = new_name
        return {"copied": source, "newName": new_ws.Name}
    return ex.call(_run)


def sheet_move(sid: str, name: str, after: str | None = None, before: str | None = None) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        ws = wb.Worksheets(name)
        if after:
            ws.Move(None, wb.Worksheets(after))
        elif before:
            ws.Move(wb.Worksheets(before), None)
        else:
            ws.Move(None, wb.Worksheets(wb.Worksheets.Count))
        return {"moved": name}
    return ex.call(_run)


def sheet_visibility(sid: str, name: str, visible: int) -> dict:
    # visible: -1 visible, 0 hidden, 2 very hidden
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        wb.Worksheets(name).Visible = visible
        return {"name": name, "visible": visible}
    return ex.call(_run)


def sheet_protect(sid: str, name: str, protect: bool = True, password: str | None = None) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        ws = wb.Worksheets(name)
        if protect:
            ws.Protect(Password=password) if password else ws.Protect()
        else:
            ws.Unprotect(Password=password) if password else ws.Unprotect()
        return {"name": name, "protected": protect}
    return ex.call(_run)
