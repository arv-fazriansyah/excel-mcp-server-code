"""Range operations."""
from __future__ import annotations

import json
from typing import Any

from session import get_session
from com_executor import get_executor
from utils import normalize_grid, sanitize, looks_like_error, check_grid_size, check_formula_values


def _ws(sid, sheet):
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    ws = ex.call(lambda: wb.Worksheets(sheet) if sheet else wb.ActiveSheet)
    return s, wb, ws, ex


def range_get_values(sid: str, sheet: str, range_address: str) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(range_address)
        raw = rng.Value
        grid = normalize_grid(raw)
        # map error cells via per-cell Text
        base_r, base_c = rng.Row, rng.Column
        if any(looks_like_error(v) for row in grid for v in row):
            for r, row in enumerate(grid):
                for c, v in enumerate(row):
                    if looks_like_error(v):
                        try:
                            grid[r][c] = str(ws.Cells(base_r + r, base_c + c).Text)
                        except Exception:
                            grid[r][c] = "#ERROR"
        check_grid_size(grid)
        return {"sheet": ws.Name, "range": rng.Address.replace("$", ""),
                "values": [[sanitize(v) for v in row] for row in grid],
                "rows": len(grid), "cols": len(grid[0]) if grid else 0}
    return ex.call(_run)


def range_get_formulas(sid: str, sheet: str, range_address: str) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(range_address)
        grid = normalize_grid(rng.Formula)
        return {"sheet": ws.Name, "range": rng.Address.replace("$", ""),
                "formulas": [[sanitize(v) for v in row] for row in grid]}
    return ex.call(_run)


def range_set_values(sid: str, sheet: str, range_address: str, values: Any) -> dict:
    # values: 2D list or scalar
    s, wb, ws, ex = _ws(sid, sheet)
    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]
    rows = [list(r) if isinstance(r, list) else [r] for r in values]
    # pad
    width = max(len(r) for r in rows)
    rows = [r + [None] * (width - len(r)) for r in rows]
    check_grid_size(rows)
    check_formula_values(rows)
    def _run():
        rng = ws.Range(range_address)
        anchor = rng
        first_r, first_c = anchor.Row, anchor.Column
        last_r = first_r + len(rows) - 1
        last_c = first_c + width - 1
        # build address
        from utils import cell_ref
        addr = cell_ref(first_r, first_c)
        if len(rows) > 1 or width > 1:
            addr += ":" + cell_ref(last_r, last_c)
        target = ws.Range(addr)
        # convert None -> Empty
        import pythoncom
        block = []
        for row in rows:
            out = []
            for v in row:
                if v is None:
                    out.append(pythoncom.Empty)
                elif isinstance(v, (dict, list)):
                    out.append(__import__("json").dumps(v, ensure_ascii=False))
                else:
                    out.append(v)
            block.append(tuple(out))
        target.Value = tuple(block)
        return {"sheet": ws.Name, "writtenRange": addr, "rows": len(rows), "cols": width}
    return ex.call(_run)


def range_set_formula(sid: str, sheet: str, range_address: str, formula: str) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    from config import ALLOW_UNSAFE_FORMULAS, is_dangerous_formula
    if not ALLOW_UNSAFE_FORMULAS and is_dangerous_formula(formula):
        raise ValueError(f"Formula berbahaya: {formula}")
    def _run():
        ws.Range(range_address).Formula = formula
        return {"sheet": ws.Name, "cell": range_address, "formula": formula}
    return ex.call(_run)


def range_clear(sid: str, sheet: str, range_address: str, contents_only: bool = False) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(range_address)
        if contents_only:
            rng.ClearContents()
        else:
            rng.Clear()
        return {"sheet": ws.Name, "cleared": range_address}
    return ex.call(_run)


def range_find(sid: str, sheet: str, range_address: str, what: str, match_case: bool = False, whole_cell: bool = False) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(range_address)
        found = rng.Find(what, None, None, 1 if whole_cell else 2, 1, 1, match_case)
        # LookAt: 1=xlWhole, 2=xlPart
        if found is None:
            return {"found": False}
        return {"found": True, "address": found.Address.replace("$", ""), "value": sanitize(found.Value)}
    return ex.call(_run)


def range_replace(sid: str, sheet: str, range_address: str, what: str, replacement: str, match_case: bool = False) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(range_address)
        ok = rng.Replace(what, replacement, 2, 1, match_case)
        return {"replaced": bool(ok)}
    return ex.call(_run)


def range_copy(sid: str, source_sheet: str, source_range: str, dest_sheet: str, dest_range: str) -> dict:
    s, wb, ws, ex = _ws(sid, source_sheet)
    s2, wb2, ws2, ex2 = _ws(sid, dest_sheet)
    def _run():
        src = ws.Range(source_range)
        dst = ws2.Range(dest_range)
        src.Copy(dst)
        return {"copied": f"{source_sheet}!{source_range}", "to": f"{dest_sheet}!{dest_range}"}
    return ex.call(_run)


def range_insert(sid: str, sheet: str, range_address: str, shift: str = "down") -> dict:
    # shift: down/right
    s, wb, ws, ex = _ws(sid, sheet)
    shift_const = -4121 if shift == "down" else -4161  # xlDown=-4121, xlToRight=-4161
    def _run():
        ws.Range(range_address).Insert(shift_const)
        return {"inserted": range_address, "shift": shift}
    return ex.call(_run)


def range_delete(sid: str, sheet: str, range_address: str, shift: str = "up") -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    shift_const = -4162 if shift == "up" else -4159  # xlUp=-4162, xlToLeft=-4159
    def _run():
        ws.Range(range_address).Delete(shift_const)
        return {"deleted": range_address, "shift": shift}
    return ex.call(_run)


def range_merge(sid: str, sheet: str, range_address: str, across: bool = False) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        ws.Range(range_address).Merge(across)
        return {"merged": range_address}
    return ex.call(_run)


def range_unmerge(sid: str, sheet: str, range_address: str) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        ws.Range(range_address).UnMerge()
        return {"unmerged": range_address}
    return ex.call(_run)


def range_number_format(sid: str, sheet: str, range_address: str, number_format: str) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        ws.Range(range_address).NumberFormat = number_format
        return {"range": range_address, "numberFormat": number_format}
    return ex.call(_run)


def range_format(sid: str, sheet: str, range_address: str, fmt: dict) -> dict:
    """fmt keys: fontName, fontSize, bold, italic, fontColor (hex), fillColor (hex),
    wrapText, horizontal, vertical, borders (bool)."""
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(range_address)
        if "fontName" in fmt or "fontSize" in fmt or "bold" in fmt or "italic" in fmt or "fontColor" in fmt:
            f = rng.Font
            if "fontName" in fmt: f.Name = fmt["fontName"]
            if "fontSize" in fmt: f.Size = fmt["fontSize"]
            if "bold" in fmt: f.Bold = bool(fmt["bold"])
            if "italic" in fmt: f.Italic = bool(fmt["italic"])
            if "fontColor" in fmt:
                f.Color = _hex_to_bgr(fmt["fontColor"])
        if "fillColor" in fmt:
            rng.Interior.Color = _hex_to_bgr(fmt["fillColor"])
        if "wrapText" in fmt:
            rng.WrapText = bool(fmt["wrapText"])
        if "horizontal" in fmt:
            rng.HorizontalAlignment = _halign(fmt["horizontal"])
        if "vertical" in fmt:
            rng.VerticalAlignment = _valign(fmt["vertical"])
        if fmt.get("borders"):
            for edge in (7, 8, 9, 10, 11, 12):  # xlEdgeLeft..xlInsideHorizontal
                try:
                    rng.Borders(edge).LineStyle = 1
                except Exception:
                    pass
        return {"formatted": range_address}
    return ex.call(_run)


def _hex_to_bgr(hexstr: str) -> int:
    h = hexstr.lstrip("#")
    if len(h) == 3:
        h = "".join(ch * 2 for ch in h)
    r = int(h[0:2], 16); g = int(h[2:4], 16); b = int(h[4:6], 16)
    return b << 16 | g << 8 | r


def _halign(v):
    return {"left": -4131, "center": -4108, "right": -4152, "justify": -4130}.get(str(v).lower(), -4131)


def _valign(v):
    return {"top": -4160, "center": -4108, "bottom": -4107}.get(str(v).lower(), -4107)


def range_sort(sid: str, sheet: str, range_address: str, key_column: int = 1, ascending: bool = True, has_header: bool = True) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    order = 1 if ascending else 2  # xlAscending=1, xlDescending=2
    header = 1 if has_header else 0  # xlYes=1, xlNo=0
    def _run():
        rng = ws.Range(range_address)
        rng.Sort(rng.Columns(key_column), order, Header=header)
        return {"sorted": range_address, "keyColumn": key_column, "ascending": ascending}
    return ex.call(_run)


def range_autofilter(sid: str, sheet: str, range_address: str, field: int | None = None, criteria1: str | None = None, operator: int | None = None, criteria2: str | None = None) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(range_address)
        if field is None:
            rng.AutoFilter()
        else:
            args = [field]
            if criteria1 is not None: args.append(criteria1)
            if operator is not None: args.append(operator)
            if criteria2 is not None: args.append(criteria2)
            rng.AutoFilter(*args)
        return {"filtered": range_address}
    return ex.call(_run)


def range_remove_duplicates(sid: str, sheet: str, range_address: str, columns: list[int] | None = None, has_header: bool = True) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    cols = columns or [1]
    def _run():
        rng = ws.Range(range_address)
        rng.RemoveDuplicates(cols, 1 if has_header else 0)
        return {"deduped": range_address}
    return ex.call(_run)


def range_used(sid: str, sheet: str) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        used = ws.UsedRange
        return {"sheet": ws.Name, "range": used.Address.replace("$", ""), "rows": used.Rows.Count, "cols": used.Columns.Count}
    return ex.call(_run)


def range_current_region(sid: str, sheet: str, anchor: str) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        cr = ws.Range(anchor).CurrentRegion
        return {"range": cr.Address.replace("$", ""), "rows": cr.Rows.Count, "cols": cr.Columns.Count}
    return ex.call(_run)


def range_special_cells(sid: str, sheet: str, range_address: str, cell_type: int, value: int | None = None) -> dict:
    # cell_type: 2=constants, 4=formulas, 5=blanks, etc. (xlCellType*)
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(range_address)
        sc = rng.SpecialCells(cell_type, value) if value is not None else rng.SpecialCells(cell_type)
        return {"range": sc.Address.replace("$", "")}
    return ex.call(_run)


def range_hyperlink(sid: str, sheet: str, cell: str, address: str, text: str | None = None, sub_address: str | None = None, screen_tip: str | None = None) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(cell)
        ws.Hyperlinks.Add(rng, address, sub_address, screen_tip, text)
        return {"cell": cell, "address": address}
    return ex.call(_run)


def range_comment(sid: str, sheet: str, cell: str, text: str, replace: bool = True) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(cell)
        if rng.Comment is not None:
            if replace:
                rng.Comment.Delete()
            else:
                rng.Comment.Text(rng.Comment.Text() + "\n" + text)
                return {"cell": cell}
        rng.AddComment(text)
        return {"cell": cell}
    return ex.call(_run)


def range_validation(sid: str, sheet: str, range_address: str, validation_type: int, formula1: str | None = None, formula2: str | None = None, alert_style: int = 1, operator: int | None = None) -> dict:
    # validation_type: 1=whole number, 2=decimal, 3=list, 4=date, 5=time, 6=text length, 7=custom
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(range_address)
        dv = rng.Validation
        dv.Delete()
        args = [validation_type]
        if operator is not None: args.append(operator)
        # Add(Type, AlertStyle, Operator, Formula1, Formula2)
        dv.Add(validation_type, alert_style, operator if operator is not None else 1, formula1, formula2)
        return {"range": range_address, "type": validation_type}
    return ex.call(_run)


def range_trace(sid: str, sheet: str, cell: str, dependents: bool = False) -> dict:
    s, wb, ws, ex = _ws(sid, sheet)
    def _run():
        rng = ws.Range(cell)
        target = rng.Dependents if dependents else rng.Precedents
        try:
            addr = target.Address.replace("$", "")
        except Exception:
            addr = None
        return {"cell": cell, "dependents": dependents, "relatedRange": addr}
    return ex.call(_run)
