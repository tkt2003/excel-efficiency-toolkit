$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot
$AppName = -join ([char[]](0x8001,0x5934,0x8868,0x683C,0x52A9,0x624B))

python -m pip show pyinstaller *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller is not installed. Installing..."
    python -m pip install pyinstaller
}

python -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --name $AppName `
  --hidden-import win32com.client `
  --hidden-import pythoncom `
  --hidden-import pywintypes `
  run_app.py

$ExePath = Join-Path $ProjectRoot ("dist\" + $AppName + ".exe")
if (-not (Test-Path $ExePath)) {
    throw "Build failed. Missing output file: $ExePath"
}

$ExeSize = (Get-Item $ExePath).Length
Write-Host "Build completed: $ExePath"
Write-Host "File size: $ExeSize bytes"

