$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot
$AppName = -join @(
    [char]0x0045, [char]0x0078, [char]0x0063, [char]0x0065, [char]0x006C,
    [char]0x6548, [char]0x7387, [char]0x5DE5, [char]0x5177, [char]0x53F0
)

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
