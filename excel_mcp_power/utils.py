"""Shared helpers."""
from __future__ import annotations

import datetime
import json
import os
from pathlib import Path

from .config import ROOT, MAX_CELLS, is_dangerous_formula, ALLOW_UNSAFE_FORMULAS


def abspath(path: str) -> str:
    return str(Path(path).expanduser().resolve())


def check_path(path: str, vba: bool = False) -> str:
    full = abspath(path)
    # sandbox
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
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value != value or value in (float("inf"), float("-inf")):
            return None
        return value
    if isinstance(value, datetime.datetime):
        if value.tzinfo is not None:
            value = value.replace(tzinfo=None)
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, (datetime.date, datetime.time)):
        return value.isoformat()
    if isinstance(value, str):
        return value
    if isinstance(value, bytes):
        return value.decode("utf-8", "replace")
    return str(value)


def normalize_grid(raw):
    if not isinstance(raw, tuple):
        return [[raw]]
    grid = []
    for row in raw:
        grid.append(list(row) if isinstance(row, tuple) else [row])
    return grid


def looks_like_error(value) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value < -2_000_000_000


def check_grid_size(grid) -> None:
    rows = len(grid)
    cols = max((len(r) for r in grid), default=0)
    if rows * cols > MAX_CELLS:
        raise ValueError(f"Area {rows}x{cols} melebihi MAX_CELLS={MAX_CELLS}")


def check_formula_values(grid) -> None:
    if ALLOW_UNSAFE_FORMULAS:
        return
    for row in grid:
        for v in row:
            if is_dangerous_formula(v):
                raise ValueError(f"Formula berbahaya diblokir: {str(v)[:80]}")


def json_dumps(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, default=str)
