<#
.SYNOPSIS
  Copies demo files from this public repo onto the demo machine – one demo or a whole bundle.

.DESCRIPTION
  One-liners (Windows PowerShell 5.1 and PowerShell 7, no admin, no modules):

    # all demos of a customer deck (recommended, the command never changes)
    & ([scriptblock]::Create((irm https://raw.githubusercontent.com/jenssgb/demo-kit-data/main/install.ps1))) -Bundle bpw

    # one demo into the same bundle folder
    & ([scriptblock]::Create((irm https://raw.githubusercontent.com/jenssgb/demo-kit-data/main/install.ps1))) -Demo cowork-copilot -Bundle bpw

    # one demo on its own (folder Demo-<demo-id>)
    & ([scriptblock]::Create((irm https://raw.githubusercontent.com/jenssgb/demo-kit-data/main/install.ps1))) -Demo <demo-id>

  Bundle = bundles/<bundle>.json (folder + list of demos). Each demo lands in <folder>\<demo-id>
  in OneDrive for work or school (falls back to personal OneDrive, then the Desktop).
  Every demo folder is self-contained: files shared with other demos ("shared" in the manifest) are copied too.
  Running it again updates: new and changed files are downloaded, unchanged files (same SHA-256) are skipped.

  -Tenant (demo VM, signed in as the demo admin, e.g. MOD Administrator in a CDX tenant):
    & ([scriptblock]::Create((irm https://raw.githubusercontent.com/jenssgb/demo-kit-data/main/install.ps1))) -Bundle bpw -Tenant
  Files are cached in %LOCALAPPDATA%\DemoKit and uploaded to the signed-in user's OneDrive with Microsoft Graph;
  then tenant.ps1 creates what the demos need in the tenant (mails, meetings, OneNote, see <demo>/tenant.json)
  and prints the steps that have to be done by hand. -Language de|en picks the mail language (default en).
  -WhatIf shows what would happen without changing the tenant. Log: %LOCALAPPDATA%\DemoKit\logs.
#>
param(
    [string]$Demo,
    [string]$Bundle,
    [string]$Repo = 'jenssgb/demo-kit-data',
    [string]$Branch = 'main',
    [string]$Target,
    [switch]$NoExplorer,
    [switch]$Tenant,
    [ValidateSet('en', 'de')] [string]$Language = 'en',
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$raw = "https://raw.githubusercontent.com/$Repo/$Branch"
$de = (Get-Culture).TwoLetterISOLanguageName -eq 'de'
function T($en, $deText) { if ($de) { $deText } else { $en } }
function Loc($o) { if ($null -eq $o) { return $null }; if ($de -and $o.de) { $o.de } else { $o.en } }
function Url($path) { "$raw/" + ((($path -split '/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/') }

if (-not $Demo -and -not $Bundle) {
    throw (T 'Use -Bundle <bundle-id> or -Demo <demo-id>.' 'Bitte -Bundle <bundle-id> oder -Demo <demo-id> angeben.')
}

if ($Tenant) {
    $logDir = Join-Path $env:LOCALAPPDATA 'DemoKit\logs'
    New-Item -ItemType Directory -Force -Path $logDir | Out-Null
    $log = Join-Path $logDir ("tenant-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
    Start-Transcript -Path $log | Out-Null
    if (-not $Target) { $Target = Join-Path $env:LOCALAPPDATA 'DemoKit' }
}

if (-not $Target) {
    $base = @($env:OneDriveCommercial, $env:OneDrive) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $base) { $base = [Environment]::GetFolderPath('Desktop') }
} else { $base = $Target }

$demos = @()
if ($Bundle) {
    try { $bundleInfo = Invoke-RestMethod -Uri (Url "bundles/$Bundle.json") -UseBasicParsing }
    catch { throw "Bundle '$Bundle' not found in $Repo ($Branch). / Bundle '$Bundle' nicht gefunden." }
    $root = Join-Path $base $bundleInfo.folder
    $demos = if ($Demo) { @($Demo) } else { @($bundleInfo.demos) }
    Write-Host ""
    Write-Host ("  " + (Loc $bundleInfo.title)) -ForegroundColor Cyan
    Write-Host "  -> $root"
} else {
    $root = $base
    $demos = @($Demo)
}

function Get-Sha([string]$file) {
    $sha = [Security.Cryptography.SHA256]::Create()
    $stream = [IO.File]::OpenRead($file)
    try { ([BitConverter]::ToString($sha.ComputeHash($stream)) -replace '-', '').ToLowerInvariant() }
    finally { $stream.Dispose(); $sha.Dispose() }
}

# Manifest entries are either "path" (old format) or { path, sha256 }; shared entries also carry "from".
function Install-File($srcPath, $relPath, $sha, $dir, $stats) {
    $dest = Join-Path $dir ($relPath -replace '/', '\')
    $exists = Test-Path $dest
    if ($exists -and $sha -and ((Get-Sha $dest) -eq $sha)) { $stats.same++; return }
    $parent = Split-Path $dest -Parent
    if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
    Invoke-WebRequest -Uri (Url $srcPath) -OutFile $dest -UseBasicParsing
    if ($exists) { $stats.updated++; $tag = T 'updated' 'aktualisiert'; $color = 'Yellow' }
    else { $stats.new++; $tag = T 'new' 'neu'; $color = 'Green' }
    Write-Host ("    {0,-13} {1}" -f $tag, $relPath) -ForegroundColor $color
}

$summary = @()
$items = @()
foreach ($d in $demos) {
    try { $manifest = Invoke-RestMethod -Uri (Url "$d/manifest.json") -UseBasicParsing }
    catch { throw "Demo '$d' not found in $Repo ($Branch). / Demo '$d' nicht gefunden." }

    $dir = if ($Bundle) { Join-Path $root $d } else { Join-Path $root "Demo-$d" }
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    Write-Host ""
    Write-Host ("  " + (Loc $manifest.title)) -ForegroundColor Cyan
    Write-Host "  -> $dir" -ForegroundColor DarkGray

    $stats = @{ new = 0; updated = 0; same = 0 }
    foreach ($f in @($manifest.files)) {
        if ($f -is [string]) { Install-File "$d/$f" $f $null $dir $stats }
        else { Install-File "$d/$($f.path)" $f.path $f.sha256 $dir $stats }
    }
    foreach ($f in @($manifest.shared)) {
        if ($null -eq $f) { continue }
        Install-File "$($f.from)/$($f.path)" $f.path $f.sha256 $dir $stats
    }
    if ($stats.new -eq 0 -and $stats.updated -eq 0) {
        Write-Host ("    " + (T 'all files up to date' 'alle Dateien aktuell')) -ForegroundColor DarkGray
    }

    $next = Loc $manifest.next
    if ($next -and -not ($Bundle -and -not $Demo)) {
        Write-Host ""
        Write-Host (T '  Next steps:' '  Naechste Schritte:') -ForegroundColor Cyan
        $n = 0
        foreach ($step in $next) { $n++; Write-Host "  $n. $step" }
    }
    $summary += [pscustomobject]@{ Demo = $d; New = $stats.new; Updated = $stats.updated; Unchanged = $stats.same }
    $drivePath = if ($Bundle) { "$($bundleInfo.folder)/$d" } else { "Demo-$d" }
    $items += @{ id = $d; dir = $dir; drivePath = $drivePath }
}

Write-Host ""
$summary | Format-Table -AutoSize | Out-String | Write-Host
if ($Bundle -and -not $Demo) {
    Write-Host (T '  Next steps per demo: see the Demo Kit page of each demo.' '  Naechste Schritte je Demo: siehe die jeweilige Seite im Demo Kit.') -ForegroundColor Cyan
    Write-Host ""
}

if ($Tenant) {
    try {
        $code = (Invoke-WebRequest -Uri (Url 'tenant.ps1') -UseBasicParsing).Content
        if ($code -is [byte[]]) { $code = [Text.Encoding]::UTF8.GetString($code) }
        & ([scriptblock]::Create($code)) -Items $items -Raw $raw -Language $Language -WhatIf:$WhatIf
    } catch {
        Write-Host ("  " + (T 'Tenant step failed: ' 'Tenant-Schritt fehlgeschlagen: ') + $_) -ForegroundColor Red
    } finally {
        Write-Host ("  Log: $log") -ForegroundColor DarkGray
        Stop-Transcript | Out-Null
    }
    return
}

$open = if ($Bundle -and -not $Demo) { $root } elseif ($Bundle) { Join-Path $root $Demo } else { Join-Path $root "Demo-$Demo" }
if (-not $NoExplorer) { Start-Process explorer.exe -ArgumentList "`"$open`"" }
