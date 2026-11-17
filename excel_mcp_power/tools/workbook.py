"""Workbook-level operations."""
from __future__ import annotations

from ..session import get_session
from ..com_executor import get_executor




def workbook_get_info(sid: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        props = {}
        try:
            for p in wb.BuiltinDocumentProperties:
                try:
                    props[p.Name] = str(p.Value)
                except Exception:
                    pass
        except Exception:
            pass
        return {"path": s.path, "name": wb.Name, "readOnly": bool(wb.ReadOnly),
                "format": int(wb.FileFormat), "hasVba": bool(wb.HasVBProject),
                "sheets": [ws.Name for ws in wb.Worksheets],
                "properties": props}
    return ex.call(_run)


def workbook_set_autosave(sid: str, enabled: bool) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        wb.AutoSaveOn = bool(enabled)
        return {"autoSave": bool(wb.AutoSaveOn)}
    return ex.call(_run)


def workbook_export_pdf(sid: str, target_path: str, from_page: int | None = None, to_page: int | None = None,
                        quality: str = "Standard", ignore_print_areas: bool = False) -> dict:
    from ..utils import check_path
    check_path(target_path)
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    import os
    os.makedirs(os.path.dirname(target_path) or ".", exist_ok=True)
    q = 0 if quality.lower().startswith("min") else 1  # xlQualityMinimum=1? actually 1=xlQualityStandard, 2=xlQualityMinimum
    # Excel XlFixedFormatQuality: xlQualityStandard=0, xlQualityMinimum=1
    qv = 1 if quality.lower().startswith("min") else 0
    def _run():
        wb.ExportAsFixedFormat(0, target_path, qv, True, ignore_print_areas,
                               from_page, to_page, False)
        return {"path": target_path}
    return ex.call(_run)


def workbook_protect(sid: str, protect: bool, password: str | None = None) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        if protect:
            wb.Protect(Password=password) if password else wb.Protect()
        else:
            wb.Unprotect(Password=password) if password else wb.Unprotect()
        return {"protected": protect}
    return ex.call(_run)


def workbook_refresh_all(sid: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        wb.RefreshAll()
        return {"refreshed": True}
    return ex.call(_run)


def workbook_external_links(sid: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    def _run():
        out = []
        try:
            for link in wb.LinkSources(1):  # xlExcelLinks=1
                out.append(str(link))
        except Exception:
            pass
        return {"links": out}
    return ex.call(_run)


def workbook_calculate(sid: str, scope: str = "workbook") -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    app = s.app
    def _run():
        if scope == "application":
            app.Calculate()
        elif scope == "sheet":
            pass
        else:
            wb.Calculate()
        return {"calculated": scope}
    return ex.call(_run)
