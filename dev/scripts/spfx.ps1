<#
.SYNOPSIS
  Runs an SPFx gulp task using the portable Node 22 in .tools\node22.

.DESCRIPTION
  The SPFx 1.23.2 build chain only accepts Node >=22.14.0 <23.0.0, while this
  machine's default Node is v24. This script prepends the portable runtime to
  PATH for the duration of the call, so nothing outside this project changes.

  Run it with pwsh (PowerShell 7+):
    pwsh -File dev/scripts/spfx.ps1 bundle
    pwsh -File dev/scripts/spfx.ps1 bundle --ship
    pwsh -File dev/scripts/spfx.ps1 package-solution --ship
    pwsh -File dev/scripts/spfx.ps1 serve

  Use dev/scripts/setup-node22.ps1 first if .tools\node22 is missing.
#>
[CmdletBinding()]
param(
  [Parameter(ValueFromRemainingArguments = $true)]
  [string[]] $GulpArgs
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$nodeDir = Join-Path $root '.tools\node22'

if (-not (Test-Path (Join-Path $nodeDir 'node.exe'))) {
  throw "Portable Node 22 not found at $nodeDir. Run: pwsh -File dev/scripts/setup-node22.ps1"
}

if (-not $GulpArgs -or $GulpArgs.Count -eq 0) {
  $GulpArgs = @('bundle')
}

$env:PATH = "$nodeDir;$nodeDir\node_modules\npm\bin;$env:PATH"

Push-Location $root
try {
  Write-Host "node $(& "$nodeDir\node.exe" -v) -> gulp $($GulpArgs -join ' ')" -ForegroundColor Cyan
  & "$nodeDir\node.exe" (Join-Path $root 'node_modules\gulp\bin\gulp.js') @GulpArgs
  exit $LASTEXITCODE
}
finally {
  Pop-Location
}
