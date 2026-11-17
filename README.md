# Excel MCP Power

Powerful Model Context Protocol (MCP) server for Microsoft Excel built in Python (COM-based / `pywin32`), featuring robust support for **`.xlsb`** read/write, **VBA** inspection/read/write/execution, tables, named ranges, and session management.

## Features

- **Multi-Format Support**: Read and write `.xlsx`, `.xlsm`, `.xlsb`, and `.xls`.
- **VBA Prototyping & Macros**: Inspect modules, read/write source code, and run macros directly via Excel's native VBE/COM interface.
- **Advanced Excel Objects**: Excel Tables (`ListObjects`), Named Ranges, Worksheet management, and 2D range read/write with formulas.
- **Security & Sandbox**: Root directory sandboxing (`EXCEL_MCP_ROOT`), read-only modes, formula injection filters, and configurable macro/VBA execution permissions.
- **Clean Lifecycle**: Thread-local STA COM execution ensuring zero lingering Excel processes.

---

## Installation

1. Ensure Python 3.10+ and Microsoft Excel are installed on Windows.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

---

## Configuration (Environment Variables)

- `EXCEL_MCP_ROOT`: Root directory for Excel files (default: `~/Documents/excel_files`).
- `EXCEL_MCP_READONLY`: Run in read-only mode (`0` or `1`).
- `EXCEL_MCP_ALLOW_MACROS`: Allow macro execution (`0` or `1`).
- `EXCEL_MCP_ALLOW_VBA_WRITE`: Allow VBA source code modification (`0` or `1`).
- `EXCEL_MCP_VISIBLE`: Show Excel GUI (`0` or `1`, default `0`).

---

## Running Smoke Tests

```bash
python tests/smoke_test.py
```
