# Excel MCP Server

Powerful MCP (Model Context Protocol) server for Microsoft Excel. Enables LLMs to interact with Excel files (.xlsx, .xlsm, .xlsb, .xls) and perform advanced automation using COM (Windows) or File-based backends.

## Quick Start (Plug & Play)

1. **Clone the repository:**
   ```bash
   git clone https://github.com/arv-fazriansyah/excel-mcp-server-code.git
   cd excel-mcp-server-code
   ```

2. **Run the setup script:**
   - **Windows:** Right-click `setup.ps1` and select "Run with PowerShell" or run `.\setup.ps1` in terminal.
   - **Linux/macOS:** Run `bash setup.sh`.

   *This will automatically install dependencies and create the working directory at `~/Documents/excel_files`.*

3. **Configure as an MCP Server:**
   Add this to your MCP client configuration (e.g., Claude Desktop, Cursor):

   ```json
   {
     "mcpServers": {
       "excel": {
         "command": "python",
         "args": ["<PATH_TO_REPO>/excel_mcp_server.py"]
       }
     }
   }
   ```

## Features

- **Automated Backend:** Uses Excel COM API (Windows) or File-based backends (Cross-platform).
- **Excel Automation:** Read/Write ranges, manage sheets, format data, create tables and charts.
- **VBA Management:** List, read, write, import, export, and execute VBA macros.

## Environment Variables

| Variable | Description | Default |
| :--- | :--- | :--- |
| `EXCEL_MCP_ROOT` | Allowed root directory | `~/Documents/excel_files` |
| `EXCEL_MCP_ALLOW_MACROS` | Set to `1` to enable VBA macro execution | `0` |

## License

MIT
