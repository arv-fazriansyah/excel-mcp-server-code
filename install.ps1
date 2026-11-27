<#
.SYNOPSIS
  Plug-and-play installer for the Excel MCP Power server (OpenCode).

.DESCRIPTION
  Clones/updates the repo, creates an isolated Python venv, installs
  dependencies, runs the smoke test, and registers the MCP server in the
  global OpenCode config. Windows + Microsoft Excel required.

.EXAMPLE
  irm https://raw.githubusercontent.com/arv-fazriansyah/excel-mcp-server-code/main/install.ps1 | iex

.EXAMPLE
  .\install.ps1 -InstallDir "D:\tools\excel-mcp"
#>
[CmdletBinding()]
param(
    [string]$InstallDir = "$env:USERPROFILE\.config\opencode\mcp-servers\excel-mcp-server-code",
    [string]$RepoUrl    = "https://github.com/arv-fazriansyah/excel-mcp-server-code.git",
    [switch]$SkipSmokeTest,
    [switch]$UseSystemPython
)

$ErrorActionPreference = "Stop"

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "WARNING: $msg" -ForegroundColor Yellow }
function Fail($msg) { Write-Host "ERROR: $msg" -ForegroundColor Red; exit 1 }

# --- 1. Platform check -------------------------------------------------------
if ($env:OS -ne "Windows_NT") { Fail "This installer supports Windows only." }

# --- 2. Prerequisite checks --------------------------------------------------
$gitCmd = Get-Command git -ErrorAction SilentlyContinue
if (-not $gitCmd) { Fail "git not found. Install Git for Windows: https://git-scm.com/download/win" }

$pyCmd = $null
foreach ($c in @("python", "py")) {
    if (Get-Command $c -ErrorAction SilentlyContinue) { $pyCmd = $c; break }
}
if (-not $pyCmd) { Fail "Python not found. Install Python 3.10+: https://www.python.org/downloads/" }

if ($pyCmd -eq "py") {
    $basePy = (& py -3 -c "import sys; print(sys.executable)").Trim()
} else {
    $basePy = (& python -c "import sys; print(sys.executable)").Trim()
}
if (-not (Test-Path $basePy)) { Fail "Could not resolve Python executable." }

$pyVer = & $basePy -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
if ([version]$pyVer -lt [version]"3.10") {
    Fail "Python 3.10+ required (found $pyVer)."
}

# Microsoft Excel check (best effort)
$excelExe = $null
try {
    $reg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction Stop
    $excelExe = $reg.'(Default)'
} catch {}
if (-not $excelExe -or -not (Test-Path $excelExe)) {
    $w = Get-Command excel -ErrorAction SilentlyContinue
    if ($w) { $excelExe = $w.Source }
}
if (-not $excelExe) {
    Write-Warn "Microsoft Excel not detected. VBA features require Excel. Continuing..."
}

$opencodeCmd = Get-Command opencode -ErrorAction SilentlyContinue
if (-not $opencodeCmd) {
    Write-Warn "opencode CLI not found in PATH. MCP registration will be skipped; add it manually."
}

# --- 3. Clone or update ------------------------------------------------------
$parent = Split-Path $InstallDir -Parent
if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }

if (Test-Path (Join-Path $InstallDir ".git")) {
    Write-Step "Updating existing installation..."
    Push-Location $InstallDir
    git pull --ff-only
    Pop-Location
} else {
    if (Test-Path $InstallDir) {
        $backup = "$InstallDir.bak.$(Get-Date -Format 'yyyyMMddHHmmss')"
        Write-Warn "Existing folder found (not a git repo). Moving to $backup"
        Move-Item $InstallDir $backup
    }
    Write-Step "Cloning repository..."
    git clone $RepoUrl $InstallDir
}

$serverPy = Join-Path $InstallDir "server.py"
if (-not (Test-Path $serverPy)) { Fail "server.py not found after clone/update." }

# --- 4. Virtual environment + dependencies ----------------------------------
if ($UseSystemPython) {
    $runPy = $basePy
    Write-Step "Using system Python: $runPy"
} else {
    $venvDir = Join-Path $InstallDir ".venv"
    $venvPy  = Join-Path $venvDir "Scripts\python.exe"
    if (-not (Test-Path $venvPy)) {
        Write-Step "Creating virtual environment..."
        & $basePy -m venv $venvDir
        if ($LASTEXITCODE -ne 0) { Fail "Failed to create venv." }
    }
    $runPy = $venvPy
    Write-Step "Installing dependencies..."
    & $runPy -m pip install --upgrade pip --quiet
    & $runPy -m pip install -r (Join-Path $InstallDir "requirements.txt")
    if ($LASTEXITCODE -ne 0) { Fail "pip install failed." }

    # pywin32 COM registration (required for Excel automation)
    $postInstall = Join-Path $venvDir "Scripts\pywin32_postinstall.py"
    if (Test-Path $postInstall) {
        Write-Step "Registering pywin32 COM..."
        & $runPy $postInstall -install | Out-Null
    }
}

# --- 5. Smoke test -----------------------------------------------------------
if (-not $SkipSmokeTest) {
    Write-Step "Running smoke test..."
    Push-Location $InstallDir
    & $runPy "tests\smoke_test.py"
    $smokeRc = $LASTEXITCODE
    Pop-Location
    if ($smokeRc -ne 0) { Fail "Smoke test failed. Check Python/Excel setup." }
}

# --- 6. Register MCP server in OpenCode --------------------------------------
$configPath = "$env:USERPROFILE\.config\opencode\opencode.json"
if ($opencodeCmd -and (Test-Path $configPath)) {
    Write-Step "Registering MCP server 'excel' in OpenCode (global)..."
    & opencode mcp add excel --global -- $runPy $serverPy
    if ($LASTEXITCODE -ne 0) {
        Write-Warn "opencode mcp add failed. Add manually to opencode.json:"
        Write-Host "  `"excel`": { `"type`": `"local`", `"command`": [`"$runPy`", `"$serverPy`"] }"
    }
} else {
    if (-not (Test-Path $configPath)) {
        Write-Warn "OpenCode config not found at $configPath. Skipping registration."
    }
    Write-Host "Manually add to your opencode.json:"
    Write-Host "  `"excel`": { `"type`": `"local`", `"command`": [`"$runPy`", `"$serverPy`"] }"
}

# --- 7. Ensure Excel files root exists ---------------------------------------
$excelRoot = "$env:USERPROFILE\Documents\excel_files"
if (-not (Test-Path $excelRoot)) { New-Item -ItemType Directory -Path $excelRoot -Force | Out-Null }

# --- Done --------------------------------------------------------------------
Write-Host ""
Write-Host "Installation complete!" -ForegroundColor Green
Write-Host "  Install dir : $InstallDir"
Write-Host "  Python      : $runPy"
Write-Host "  Excel files : $excelRoot"
Write-Host ""
Write-Host "Next: restart OpenCode (or run 'opencode service restart') to load the excel MCP." -ForegroundColor Yellow
