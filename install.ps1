<#
.SYNOPSIS
  Copies the files of one demo from this public repo onto the demo machine.

.DESCRIPTION
  One-liner (Windows PowerShell 5.1 and PowerShell 7, no admin, no modules):

    & ([scriptblock]::Create((irm https://raw.githubusercontent.com/jenssgb/demo-kit-data/main/install.ps1))) -Demo <demo-id>

  Reads <demo-id>/manifest.json, downloads every listed file into a new folder
  "Demo-<demo-id>" in OneDrive for work or school (falls back to personal OneDrive,
  then the Desktop) and opens the folder in Explorer.
#>
param(
    [Parameter(Mandatory = $true)][string]$Demo,
    [string]$Repo = 'jenssgb/demo-kit-data',
    [string]$Branch = 'main',
    [string]$Target
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$raw = "https://raw.githubusercontent.com/$Repo/$Branch/$Demo"
$de = (Get-Culture).TwoLetterISOLanguageName -eq 'de'

try {
    $manifest = Invoke-RestMethod -Uri "$raw/manifest.json" -UseBasicParsing
} catch {
    throw "Demo '$Demo' not found in $Repo ($Branch). Check the demo id. / Demo '$Demo' nicht gefunden."
}

if (-not $Target) {
    $base = @($env:OneDriveCommercial, $env:OneDrive) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $base) { $base = [Environment]::GetFolderPath('Desktop') }
    $Target = Join-Path $base "Demo-$Demo"
}
New-Item -ItemType Directory -Force -Path $Target | Out-Null

$title = if ($de -and $manifest.title.de) { $manifest.title.de } else { $manifest.title.en }
Write-Host ""
Write-Host "  $title" -ForegroundColor Cyan
Write-Host "  -> $Target"
Write-Host ""

$i = 0
foreach ($file in $manifest.files) {
    $i++
    $dest = Join-Path $Target ($file -replace '/', '\')
    $dir = Split-Path $dest -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $url = "$raw/" + ((($file -split '/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/')
    Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
    Write-Host ("  [{0}/{1}] {2}" -f $i, $manifest.files.Count, $file) -ForegroundColor Green
}

$next = if ($de -and $manifest.next.de) { $manifest.next.de } else { $manifest.next.en }
if ($next) {
    Write-Host ""
    Write-Host ($(if ($de) { '  Naechste Schritte:' } else { '  Next steps:' })) -ForegroundColor Cyan
    $n = 0
    foreach ($step in $next) { $n++; Write-Host "  $n. $step" }
}
Write-Host ""

Start-Process explorer.exe -ArgumentList "`"$Target`""
