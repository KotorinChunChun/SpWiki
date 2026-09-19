<#
.SYNOPSIS
  Downloads the portable Node 22 runtime that the SPFx build requires.

.DESCRIPTION
  Fetches the official Windows x64 ZIP from nodejs.org, verifies it against the
  published SHA256 checksum, and extracts it to .tools\node22. Nothing is
  installed system wide and PATH is not modified; dev/scripts/spfx.ps1 points at this
  copy only while a build runs.

  Run it with pwsh (PowerShell 7+):
    pwsh -File dev/scripts/setup-node22.ps1
#>
[CmdletBinding()]
param(
  [string] $Version = 'v22.23.2'
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$toolsDir = Join-Path $root '.tools'
$target = Join-Path $toolsDir 'node22'
$name = "node-$Version-win-x64"
$zip = Join-Path $toolsDir "$name.zip"

if (Test-Path (Join-Path $target 'node.exe')) {
  Write-Host "Already present: $(& "$target\node.exe" -v) at $target"
  return
}

New-Item -ItemType Directory -Force $toolsDir | Out-Null

Write-Host "Downloading $name.zip ..."
Invoke-WebRequest "https://nodejs.org/dist/$Version/$name.zip" -OutFile $zip

Write-Host 'Verifying checksum ...'
$sums = (Invoke-WebRequest "https://nodejs.org/dist/$Version/SHASUMS256.txt").Content
$expected = (($sums -split "`n" | Where-Object { $_ -match [regex]::Escape("$name.zip") }) -split '\s+')[0]
$actual = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()

if ($expected -ne $actual) {
  Remove-Item $zip -Force
  throw "Checksum mismatch. expected=$expected actual=$actual"
}

Expand-Archive $zip -DestinationPath $toolsDir -Force
Rename-Item (Join-Path $toolsDir $name) 'node22'
Remove-Item $zip -Force

Write-Host "Installed $(& "$target\node.exe" -v) to $target"
