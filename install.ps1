<# 
.SYNOPSIS
    One-line installer for Excel MCP Server.
    Usage: iex (iwr -useb "https://raw.githubusercontent.com/arv-fazriansyah/excel-mcp-server-code/main/install.ps1").Content
#>

$ErrorActionPreference = "Stop"
$repoUrl = "https://github.com/arv-fazriansyah/excel-mcp-server-code.git"
$installDir = Join-Path $env:LOCALAPPDATA "ExcelMCPServer"

Write-Host "🚀 Installing Excel MCP Server..." -ForegroundColor Cyan

# 1. Clone or update repo
if (Test-Path $installDir) {
    Write-Host "📁 Updating existing installation..." -ForegroundColor Yellow
    git -C $installDir pull
} else {
    Write-Host "📥 Cloning repository..." -ForegroundColor Yellow
    git clone $repoUrl $installDir
}

# 2. Install dependencies
Write-Host "📦 Installing Python dependencies..." -ForegroundColor Yellow
pip install -r (Join-Path $installDir "requirements.txt")

# 3. Create working directory
$docDir = Join-Path $env:USERPROFILE "Documents" "excel_files"
if (!(Test-Path $docDir)) {
    New-Item -ItemType Directory -Path $docDir | Out-Null
    Write-Host "📂 Created working directory: $docDir" -ForegroundColor Green
}

# 4. Show MCP configuration
$serverPath = Join-Path $installDir "excel_mcp_server.py"
Write-Host "`n✅ Installation complete!" -ForegroundColor Green
Write-Host "`n📋 Add this to your MCP client config (Claude Desktop, Cursor, etc.):" -ForegroundColor Cyan
@"
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["$serverPath"]
    }
  }
}
"@

Write-Host "`n📁 Your Excel files go in: $docDir" -ForegroundColor Gray