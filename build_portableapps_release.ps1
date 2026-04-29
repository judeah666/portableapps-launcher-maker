param(
  [string]$Python64 = $env:PORTABLEAPPS_X64_PYTHON
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if (-not $Python64) {
  $uvPython = Get-Command uv -ErrorAction SilentlyContinue
  if ($uvPython) {
    try {
      $candidate = & $uvPython.Source python find 3.13 2>$null
      if ($LASTEXITCODE -eq 0 -and $candidate) {
        $Python64 = ($candidate | Out-String).Trim()
      }
    } catch {
    }
  }
}

if (-not $Python64) {
  $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
  if ($pyLauncher) {
    foreach ($selector in @("-3.13", "-3.14")) {
      try {
        $candidate = & $pyLauncher.Source $selector -c "import sys; print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0 -and $candidate) {
          $Python64 = ($candidate | Out-String).Trim()
          break
        }
      } catch {
      }
    }
  }
}

if (-not $Python64) {
  $resolvedPython = Get-Command python -ErrorAction SilentlyContinue
  if ($resolvedPython) {
    $Python64 = $resolvedPython.Source
  }
}

if (-not $Python64) {
  throw @"
No 64-bit Python was provided.

Set the PORTABLEAPPS_X64_PYTHON environment variable or pass -Python64 with the
full path to a 64-bit python.exe.
"@
}

if (-not (Test-Path -LiteralPath $Python64)) {
  throw "64-bit Python not found: $Python64"
}

$pythonBits = & $Python64 -c "import struct; print(struct.calcsize('P') * 8)"
if ($LASTEXITCODE -ne 0) {
  throw "Failed to query Python architecture from: $Python64"
}

if (($pythonBits | Out-String).Trim() -ne "64") {
  throw "The supplied Python is not 64-bit: $Python64"
}

& $Python64 -m PyInstaller `
  --noconfirm `
  --clean `
  --onedir `
  --windowed `
  --noupx `
  --contents-directory . `
  --icon "app\assets\icons\software_icon.ico" `
  --add-data "app\assets;app\assets" `
  --name PortableAppsLauncherMaker `
  app\portableapps_main.py

if ($LASTEXITCODE -ne 0) {
  throw "PyInstaller x64 build failed."
}

$intermediateExe = Join-Path $root "build\PortableAppsLauncherMaker\PortableAppsLauncherMaker.exe"
if (Test-Path $intermediateExe) {
  Remove-Item -LiteralPath $intermediateExe -Force
}

$notePath = Join-Path $root "build\PortableAppsLauncherMaker\DO_NOT_RUN_FROM_BUILD_FOLDER.txt"
@"
Do not run PortableAppsLauncherMaker.exe from this build folder.

This folder is PyInstaller's intermediate work area and does not contain the
complete runtime DLL layout.

Run the release app from:

$root\dist\PortableAppsLauncherMaker\PortableAppsLauncherMaker.exe
"@ | Set-Content -LiteralPath $notePath -Encoding UTF8

Write-Host ""
Write-Host "Release build ready:"
Write-Host "$root\dist\PortableAppsLauncherMaker\PortableAppsLauncherMaker.exe"
