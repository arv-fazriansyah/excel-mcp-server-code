"""Named range operations."""
from __future__ import annotations

from ..session import get_session
from ..com_executor import get_executor


def named_range_list(sid: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        out = []
        for n in wb.Names:
            try:
                out.append({"name": n.Name, "refersTo": n.RefersTo, "value": str(n.Value)})
            except Exception:
                out.append({"name": n.Name, "refersTo": n.RefersTo})
        return {"namedRanges": out}
    return ex.call(_run)


def named_range_add(sid: str, name: str, refers_to: str, sheet: str | None = None) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        if sheet:
            ws = wb.Worksheets(sheet)
            ws.Names.Add(Name=name, RefersTo=refers_to)
        else:
            wb.Names.Add(Name=name, RefersTo=refers_to)
        return {"name": name, "refersTo": refers_to}
    return ex.call(_run)


def named_range_delete(sid: str, name: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        wb.Names(name).Delete()
        return {"deleted": name}
    return ex.call(_run)
