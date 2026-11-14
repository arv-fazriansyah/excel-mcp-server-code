# setup.ps1
Write-Host "Setting up Excel MCP Server..." -ForegroundColor Cyan
pip install -r requirements.txt

$docDir = [System.IO.Path]::Combine($env:USERPROFILE, "Documents", "excel_files")
if (!(Test-Path $docDir)) {
    New-Item -ItemType Directory -Path $docDir | Out-Null
    Write-Host "Created working directory at: $docDir" -ForegroundColor Green
}

Write-Host "Setup complete!" -ForegroundColor Green