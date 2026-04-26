param(
  [switch]$Skip64,
  [switch]$Skip32
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if ($Skip64 -and $Skip32) {
  throw "Nothing to package. Remove -Skip64 or -Skip32."
}

$releaseDir = Join-Path $root "dist\release"
New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null

function New-ReleaseZip {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDir,
    [Parameter(Mandatory = $true)]
    [string]$ZipName
  )

  $fullSource = Join-Path $root $SourceDir
  if (-not (Test-Path -LiteralPath $fullSource)) {
    throw "Build folder not found: $fullSource"
  }

  $zipPath = Join-Path $releaseDir $ZipName
  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
  }

  Compress-Archive -LiteralPath $fullSource -DestinationPath $zipPath -CompressionLevel Optimal
  Write-Host "Created $zipPath"
}

if (-not $Skip64) {
  New-ReleaseZip -SourceDir "dist\PortableAppsLauncherMaker" -ZipName "PortableAppsLauncherMaker-win64.zip"
}

if (-not $Skip32) {
  New-ReleaseZip -SourceDir "dist\PortableAppsLauncherMaker-x86" -ZipName "PortableAppsLauncherMaker-win32.zip"
}
