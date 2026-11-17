"""Excel Table (ListObject) operations."""
from __future__ import annotations

from session import get_session
from com_executor import get_executor


def table_list(sid: str, sheet: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        ws = wb.Worksheets(sheet)
        out = []
        for lo in ws.ListObjects:
            out.append({"name": lo.Name, "range": lo.Range.Address.replace("$", ""),
                        "headerRowRange": lo.HeaderRowRange.Address.replace("$", "") if lo.HeaderRowRange else None,
                        "dataRange": lo.DataBodyRange.Address.replace("$", "") if lo.DataBodyRange else None})
        return {"tables": out}
    return ex.call(_run)


def table_add(sid: str, sheet: str, range_address: str, name: str | None = None, has_header: bool = True) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    header_const = 1 if has_header else 2  # xlYes=1, xlNo=2
    def _run():
        ws = wb.Worksheets(sheet)
        rng = ws.Range(range_address)
        lo = ws.ListObjects.Add(1, rng, None, header_const)  # xlSrcRange=1
        if name:
            lo.Name = name
        return {"name": lo.Name, "range": lo.Range.Address.replace("$", "")}
    return ex.call(_run)


def table_delete(sid: str, sheet: str, name: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        ws = wb.Worksheets(sheet)
        ws.ListObjects(name).Delete()
        return {"deleted": name}
    return ex.call(_run)
