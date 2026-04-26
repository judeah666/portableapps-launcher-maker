param(
  [string]$Python32 = $env:PORTABLEAPPS_X86_PYTHON
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if (-not $Python32) {
  throw @"
No 32-bit Python was provided.

Set the PORTABLEAPPS_X86_PYTHON environment variable or pass -Python32 with the
full path to a 32-bit python.exe.

Example:
powershell -ExecutionPolicy Bypass -File .\build_portableapps_release_x86.ps1 -Python32 "C:\Path\To\Python32\python.exe"
"@
}

if (-not (Test-Path -LiteralPath $Python32)) {
  throw "32-bit Python not found: $Python32"
}

$pythonBits = & $Python32 -c "import struct; print(struct.calcsize('P') * 8)"
if ($LASTEXITCODE -ne 0) {
  throw "Failed to query Python architecture from: $Python32"
}

if (($pythonBits | Out-String).Trim() -ne "32") {
  throw "The supplied Python is not 32-bit: $Python32"
}

$buildName = "PortableAppsLauncherMaker-x86"

& $Python32 -m PyInstaller `
  --noconfirm `
  --clean `
  --onedir `
  --windowed `
  --noupx `
  --contents-directory . `
  --icon "app\assets\software_icon.ico" `
  --add-data "app\assets;app\assets" `
  --add-data "app\help_template;app\help_template" `
  --name $buildName `
  app\portableapps_main.py

if ($LASTEXITCODE -ne 0) {
  throw "PyInstaller x86 build failed."
}

$intermediateExe = Join-Path $root "build\$buildName\$buildName.exe"
if (Test-Path $intermediateExe) {
  Remove-Item -LiteralPath $intermediateExe -Force
}

$notePath = Join-Path $root "build\$buildName\DO_NOT_RUN_FROM_BUILD_FOLDER.txt"
@"
Do not run $buildName.exe from this build folder.

This folder is PyInstaller's intermediate work area and does not contain the
complete runtime DLL layout.

Run the release app from:

$root\dist\$buildName\$buildName.exe
"@ | Set-Content -LiteralPath $notePath -Encoding UTF8

Write-Host ""
Write-Host "32-bit release build ready:"
Write-Host "$root\dist\$buildName\$buildName.exe"
