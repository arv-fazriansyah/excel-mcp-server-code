# Excel MCP Server

Powerful MCP (Model Context Protocol) server for Microsoft Excel. Enables LLMs to interact with Excel files (.xlsx, .xlsm, .xlsb, .xls) and perform advanced automation using COM (Windows) or File-based backends.

## Features

- **Automated Backend:** Automatically detects and uses Excel COM API (Windows) for full functionality or File-based backends for cross-platform support.
- **Excel Automation:** Read/Write ranges, manage sheets, format data, create tables and charts.
- **VBA Management:** List, read, write, import, export, and execute VBA macros (with configurable safety).
- **Safe & Configurable:** Built-in formula safety filters, read-only mode, and granular VBA control.

## Prerequisites

- **Windows:** Microsoft Excel installed.
- **Python:** 3.10+
- **Packages:** `pywin32` (Windows/COM), `openpyxl`, `pyxlsb`, `oletools`, `mcp`.

## Quick Start (Plug & Play)

1. **Clone the repository:**
   ```bash
   git clone https://github.com/arv-fazriansyah/excel-mcp-server-code.git
   cd excel-mcp-server-code
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure as an MCP Server:**
   Add the following to your MCP client configuration (e.g., Cursor, Claude Desktop):

   ```json
   {
     "mcpServers": {
       "excel": {
         "command": "python",
         "args": ["C:/path/to/excel_mcp_server.py"]
       }
     }
   }
   ```

4. **Usage:**
   The server automatically uses `~/Documents/excel_files` as the working directory. Simply place your Excel files in this folder for the server to access them securely.

## Environment Variables

| Variable | Description | Default |
| :--- | :--- | :--- |
| `EXCEL_MCP_ROOT` | Allowed root directory for Excel files | `Documents/excel_files` |
| `EXCEL_MCP_BACKEND` | `auto`, `com`, or `file` | `auto` |
| `EXCEL_MCP_ALLOW_MACROS` | Set to `1` to enable VBA macro execution | `0` |
| `EXCEL_MCP_READONLY` | Set to `1` to block all writing operations | `0` |

## License

MIT
