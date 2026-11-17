"""Smoke test for excel-mcp-power."""
from __future__ import annotations

import os
import sys
from pathlib import Path

# ensure package in path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config import ROOT
from session import session_create, session_open, session_close, session_info
from tools import sheet, range, vba, table, named
from com_executor import get_executor


def test_workflow():
    print("--- Running excel-mcp-power Smoke Test ---")
    test_file = os.path.join(ROOT, "test_power.xlsb")
    if os.path.exists(test_file):
        os.remove(test_file)

    # 1. Create workbook (.xlsb)
    print("1. Creating workbook (.xlsb)...")
    res = session_create(test_file, sheets="Summary,Data", force=True)
    sid = res["sessionId"]
    print("   Created session:", sid, "sheets:", res["sheets"])

    # 2. Sheet operations
    print("2. Testing sheet operations...")
    shs = sheet.sheet_list(sid)
    print("   Sheets:", [s["name"] for s in shs["sheets"]])
    sheet.sheet_add(sid, name="Extra")
    sheet.sheet_rename(sid, "Extra", "Details")

    # 3. Range write and read
    print("3. Testing range write/read...")
    range.range_set_values(sid, "Data", "A1:C3", [
        ["ID", "Name", "Score"],
        [1, "Alice", 95],
        [2, "Bob", 88]
    ])
    vals = range.range_get_values(sid, "Data", "A1:C3")
    print("   Read values:", vals["values"])
    assert vals["values"][1][1] == "Alice"

    # 4. Range formula
    print("4. Testing formula...")
    range.range_set_formula(sid, "Data", "D1", "Status")
    range.range_set_formula(sid, "Data", "D2", "=IF(C2>90,\"High\",\"Normal\")")
    fvals = range.range_get_values(sid, "Data", "D1:D2")
    print("   Formula values:", fvals["values"])

    # 5. Table & Named range
    print("5. Testing Table & Named range...")
    table.table_add(sid, "Data", "A1:C3", name="ScoreTable")
    tbls = table.table_list(sid, "Data")
    print("   Tables:", [t["name"] for t in tbls["tables"]])

    named.named_range_add(sid, "TopScore", "Data!$C$2")
    nrs = named.named_range_list(sid)
    print("   Named ranges:", [n["name"] for n in nrs["namedRanges"]])

    # 6. Save and close
    print("6. Saving and closing session...")
    session_close(sid, save=True)
    print("   Session closed successfully.")

    # 7. Re-open and verify
    print("7. Re-opening saved workbook (.xlsb)...")
    res2 = session_open(test_file)
    sid2 = res2["sessionId"]
    info = session_info(sid2)
    print("   Re-opened info - sheets:", [s["name"] for s in info["sheets"]])
    session_close(sid2)

    print("=== ALL SMOKE TESTS PASSED SUCCESSFULLY! ===")


if __name__ == "__main__":
    try:
        test_workflow()
    finally:
        get_executor().shutdown()
