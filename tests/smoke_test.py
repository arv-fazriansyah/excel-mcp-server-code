"""Smoke test for consolidated excel-mcp-power."""
from __future__ import annotations

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from excel_engine import (
    ROOT, session_create, session_open, session_close, session_info,
    sheet_list, sheet_add, sheet_rename,
    range_get_values, range_set_values, range_set_formula,
    table_add, table_list, named_range_add, named_range_list
)


def test_workflow():
    print("--- Running Consolidated Smoke Test ---")
    test_file = os.path.join(ROOT, "test_power.xlsb")
    if os.path.exists(test_file): os.remove(test_file)

    # 1. Create workbook
    res = session_create(test_file, sheets="Summary,Data", force=True)
    sid = res["sessionId"]

    # 2. Sheet operations
    sheet_add(sid, name="Details")

    # 3. Range write/read
    range_set_values(sid, "Data", "A1:C3", [
        ["ID", "Name", "Score"],
        [1, "Alice", 95],
        [2, "Bob", 88]
    ])
    vals = range_get_values(sid, "Data", "A1:C3")
    print("Read values:", vals["values"])
    assert vals["values"][1][1] == "Alice"

    # 4. Formula
    range_set_formula(sid, "Data", "D1", "Status")
    range_set_formula(sid, "Data", "D2", '=IF(C2>90,"High","Normal")')

    # 5. Table & Named range
    table_add(sid, "Data", "A1:C3", name="ScoreTable")
    named_range_add(sid, "TopScore", "Data!$C$2")

    # 6. Save & close
    session_close(sid, save=True)

    # 7. Re-open
    res2 = session_open(test_file)
    sid2 = res2["sessionId"]
    info = session_info(sid2)
    print("Re-opened sheets:", [s["name"] for s in info["sheets"]])
    session_close(sid2)

    print("=== ALL CONSOLIDATED TESTS PASSED SUCCESSFULLY! ===")


if __name__ == "__main__":
    test_workflow()
