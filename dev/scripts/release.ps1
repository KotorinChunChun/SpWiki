<#
.SYNOPSIS
  SpWiki の配布用パッケージを作成します。
.DESCRIPTION
  clean、bundle --ship、package-solution --ship を実行し、
  packages/<バージョン>/SpWiki.sppkg に保存します。
  ファイル名は固定し、バージョンはフォルダー名で管理します。
.PARAMETER Version
  SPFx の4桁バージョンを指定します。初版は 1.0.0.0 です。
.PARAMETER Bump
  Major、Minor、Build のいずれかを増やします。
.PARAMETER SkipClean
  lib の再利用を許可します。削除・改名を伴う変更では使用しないでください。
.EXAMPLE
  pwsh -File dev/scripts/release.ps1 -Version 1.0.0.0
#>
[CmdletBinding()]
param(
  [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
  [string] $Version,

  [ValidateSet('Major', 'Minor', 'Build')]
  [string] $Bump,

  [switch] $SkipClean
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$spfx = Join-Path $PSScriptRoot 'spfx.ps1'
$solutionConfigPath = Join-Path $root 'config\package-solution.json'
$packageJsonPath = Join-Path $root 'package.json'
$packagesDir = Join-Path $root 'packages'

function Invoke-Spfx {
  param([string[]] $Arguments)

  & pwsh -NoProfile -File $spfx @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "gulp $($Arguments -join ' ') failed with exit code $LASTEXITCODE"
  }
}

if ($Version -and $Bump) {
  throw 'Specify either -Version or -Bump, not both.'
}

if (-not (Test-Path (Join-Path $root '.tools\node22\node.exe'))) {
  throw 'Portable Node 22 is missing. Run: npm run setup:node22'
}

if (-not (Test-Path (Join-Path $root 'node_modules'))) {
  throw 'Dependencies are missing. Run: npm install'
}

# ---------------------------------------------------------------- version

$solutionConfig = Get-Content $solutionConfigPath -Raw | ConvertFrom-Json
$currentVersion = $solutionConfig.solution.version

if ($Bump) {
  $parts = $currentVersion.Split('.')
  switch ($Bump) {
    'Major' { $parts[0] = [int]$parts[0] + 1; $parts[1] = 0; $parts[2] = 0 }
    'Minor' { $parts[1] = [int]$parts[1] + 1; $parts[2] = 0 }
    'Build' { $parts[2] = [int]$parts[2] + 1 }
  }
  $Version = $parts -join '.'
}

if ($Version -and $Version -ne $currentVersion) {
  Write-Host "Version: $currentVersion -> $Version" -ForegroundColor Cyan

  # Rewrite the raw text so the rest of the file keeps its formatting.
  # The pattern has to be built into a variable first: an inline `+` would be
  # applied before -replace and silently produce a no-op.
  # Both "version" fields carry the old value (solution and feature), and both
  # are meant to move together.
  $pattern = '("version"\s*:\s*")' + [regex]::Escape($currentVersion) + '(")'
  $raw = Get-Content $solutionConfigPath -Raw
  $raw = $raw -replace $pattern, "`${1}$Version`${2}"
  Set-Content $solutionConfigPath $raw -NoNewline -Encoding utf8

  $npmVersion = ($Version.Split('.')[0..2]) -join '.'
  $packageRaw = Get-Content $packageJsonPath -Raw
  $packageRaw = $packageRaw -replace '("version"\s*:\s*")[^"]+(")', "`${1}$npmVersion`${2}"
  Set-Content $packageJsonPath $packageRaw -NoNewline -Encoding utf8

  # Fail loudly rather than shipping a package whose version did not move.
  $written = (Get-Content $solutionConfigPath -Raw | ConvertFrom-Json).solution.version
  if ($written -ne $Version) {
    throw "Failed to write the solution version: config\package-solution.json still says $written"
  }
}
else {
  $Version = $currentVersion
  Write-Host "Version: $Version (unchanged)" -ForegroundColor Cyan
}

# ---------------------------------------------------------------- build

if (-not $SkipClean) {
  Write-Host 'Cleaning previous output ...' -ForegroundColor Cyan
  Invoke-Spfx @('clean')
}
else {
  # The bundle output goes even when the slow tsc output in lib\ is kept.
  # package-solution packages everything sitting in release\assets, so a debug
  # bundle left there by an earlier `gulp bundle` would ship next to the
  # minified one, unnecessarily increasing the package size.
  foreach ($stale in @('dist', 'release')) {
    $path = Join-Path $root $stale
    if (Test-Path $path) {
      Write-Host "Removing $stale\ so no debug bundle is packaged ..." -ForegroundColor Cyan
      Remove-Item $path -Recurse -Force
    }
  }
}

Write-Host 'Baking the documentation into the bundle ...' -ForegroundColor Cyan
& node (Join-Path $root 'dev\scripts\build-docs.mjs')
if ($LASTEXITCODE -ne 0) { throw 'build-docs failed' }

Write-Host 'Building production bundle ...' -ForegroundColor Cyan
Invoke-Spfx @('bundle', '--ship')

Write-Host 'Packaging solution ...' -ForegroundColor Cyan
Invoke-Spfx @('package-solution', '--ship')

# ---------------------------------------------------------------- artifact

$built = Join-Path $root 'sharepoint\solution\SpWiki.sppkg'
if (-not (Test-Path $built)) {
  throw "Expected package not found: $built"
}

# Version in the folder name, fixed file name inside it — see the notes above.
$versionDir = Join-Path $packagesDir $Version
New-Item -ItemType Directory -Force $versionDir | Out-Null
$artifact = Join-Path $versionDir 'SpWiki.sppkg'
Copy-Item $built $artifact -Force

$sizeKb = [math]::Round((Get-Item $artifact).Length / 1KB, 1)
$hash = (Get-FileHash $artifact -Algorithm SHA256).Hash.ToLower()

Write-Host ''
Write-Host '================ Release ready ================' -ForegroundColor Green
Write-Host "  Package : $artifact"
Write-Host "  Version : $Version"
Write-Host "  Size    : $sizeKb KB"
Write-Host "  SHA256  : $hash"
Write-Host ''
Write-Host '  Next steps' -ForegroundColor Green
Write-Host '   1. Open the tenant app catalog:'
Write-Host '        https://<tenant>.sharepoint.com/sites/appcatalog'
Write-Host '        -> Apps for SharePoint'
Write-Host '   2. Upload SpWiki.sppkg and choose Deploy.'
Write-Host '        The file name must stay SpWiki.sppkg, otherwise the catalog'
Write-Host '        treats it as a new package and rejects it as a duplicate product ID.'
Write-Host '        Tick "Make this solution available to all sites" for tenant-wide use.'
Write-Host '   3. Add the app to the target site (skip when deployed tenant-wide).'
Write-Host '   4. Put the .md files in a document library.'
Write-Host '   5. Add the "SpWiki" web part to a page and set the library'
Write-Host '      name and start file, then publish the page.'
Write-Host '=============================================='
