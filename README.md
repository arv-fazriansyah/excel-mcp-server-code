# Excel MCP Server

Powerful MCP (Model Context Protocol) server for Microsoft Excel. Enables LLMs to interact with Excel files (.xlsx, .xlsm, .xlsb, .xls) and perform advanced automation using COM (Windows) or File-based backends.

## ⚡ One-Line Install (Recommended)

### Windows (PowerShell)
```powershell
iex (iwr -useb "https://raw.githubusercontent.com/arv-fazriansyah/excel-mcp-server-code/main/install.ps1").Content
```

### Linux / macOS (Bash)
```bash
bash -c "$(curl -fsSL https://raw.githubusercontent.com/arv-fazriansyah/excel-mcp-server-code/main/install.sh)"
```

**What this does:**
- Clones/updates the server to `%LOCALAPPDATA%\ExcelMCPServer` (Windows) or `~/.local/share/ExcelMCPServer` (Linux/macOS)
- Installs Python dependencies automatically
- Creates your working directory at `~/Documents/excel_files`
- Prints the exact MCP configuration to copy-paste

---

## Manual Setup (Alternative)

1. **Clone the repository:**
   ```bash
   git clone https://github.com/arv-fazriansyah/excel-mcp-server-code.git
   cd excel-mcp-server-code
   ```

2. **Run the setup script:**
   - **Windows:** `.\setup.ps1`
   - **Linux/macOS:** `bash setup.sh`

3. **Configure MCP Client:**
   Add to your MCP config (Claude Desktop, Cursor, etc.):
   ```json
   {
     "mcpServers": {
       "excel": {
         "command": "python",
         "args": ["<FULL_PATH_TO>/excel_mcp_server.py"]
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