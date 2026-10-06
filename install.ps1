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
  -WhatIf shows what would happen without changing the tenant.
  Every run writes a log to Desktop\DemoKit-Logs (last 20 kept) - send that file when something goes wrong.
  Demo people are the real users of the CDX demo tenant (Teresa Sac, Vance DeLeon, ...), written directly into the files.
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

# ---------- log on the Desktop (every run) ----------
$log = $null; $failed = $false
try {
    $logDir = Join-Path ([Environment]::GetFolderPath('Desktop')) 'DemoKit-Logs'
    New-Item -ItemType Directory -Force -Path $logDir | Out-Null
    $name = if ($Bundle) { $Bundle } elseif ($Demo) { $Demo } else { 'run' }
    $log = Join-Path $logDir ("{0}{1}-{2:yyyyMMdd-HHmmss}.log" -f $name, $(if ($Tenant) { '-tenant' } else { '' }), (Get-Date))
    Get-ChildItem $logDir -Filter *.log | Sort-Object LastWriteTime -Descending | Select-Object -Skip 19 | Remove-Item -Force -ErrorAction SilentlyContinue
    Start-Transcript -Path $log | Out-Null
    $commit = try { (Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/commits/$Branch" -UseBasicParsing -TimeoutSec 5).sha.Substring(0, 7) } catch { '?' }
    $os = try { (Get-CimInstance Win32_OperatingSystem).Caption } catch { [Environment]::OSVersion.VersionString }
    Write-Host ("  Demo Kit data | {0:yyyy-MM-dd HH:mm:ss} | {1} | {2}\{3} | {4} | PowerShell {5} | {6}@{7} ({8})" -f (Get-Date), $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERNAME, $os, $PSVersionTable.PSVersion, $Repo, $Branch, $commit) -ForegroundColor DarkGray
    Write-Host ("  Bundle={0} Demo={1} Tenant={2} Language={3} WhatIf={4} Target={5}" -f $Bundle, $Demo, [bool]$Tenant, $Language, [bool]$WhatIf, $Target) -ForegroundColor DarkGray
} catch { $log = $null }

try {
if (-not $Demo -and -not $Bundle) {
    throw (T 'Use -Bundle <bundle-id> or -Demo <demo-id>.' 'Bitte -Bundle <bundle-id> oder -Demo <demo-id> angeben.')
}

if ($Tenant -and -not $Target) { $Target = Join-Path $env:LOCALAPPDATA 'DemoKit' }

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
$links = @{ demos = [ordered]@{} }

# ---------- Desktop link page + Edge favorites ----------
function HtmlEnc($s) { [Net.WebUtility]::HtmlEncode("$s") }
function FileUri($p) { ([uri]$p).AbsoluteUri }
function Get-LinkGroups {
    # general links + one group per demo: @{ name; links = @(@{ name; url }) }
    $gen = @(@{ name = 'Microsoft Copilot'; url = 'https://copilot.cloud.microsoft/' })
    if ($links.root) {
        $gen += @{ name = 'Outlook'; url = 'https://outlook.office.com/mail/' }
        $gen += @{ name = (T 'OneDrive folder' 'OneDrive-Ordner') + " $($bundleName)"; url = $links.root }
    }
    $gen += @{ name = (T 'Local folder' 'Lokaler Ordner') + " $($bundleName)"; url = (FileUri $localRoot) }
    $groups = @(@{ name = (T 'Start' 'Start'); links = $gen })
    foreach ($id in $links.demos.Keys) {
        $d = $links.demos[$id]
        $l = @()
        if ($d.web) { $l += @{ name = (T 'OneDrive folder' 'OneDrive-Ordner'); url = $d.web } }
        if ($d.onenote) { $l += $d.onenote }
        foreach ($x in @($d.extra)) { if ($x) { $l += $x } }
        if ($d.files -and @($d.files).Count) { foreach ($f in $d.files) { if ($f.url) { $l += @{ name = $f.name; url = $f.url } } } }
        else { foreach ($f in @($d.local)) { $l += @{ name = $f; url = (FileUri (Join-Path $d.dir ($f -replace '/', '\'))) } } }
        $groups += @{ name = $d.title; links = $l }
    }
    $groups
}
function Write-LinkPage($groups) {
    $file = Join-Path ([Environment]::GetFolderPath('Desktop')) "$bundleName - Links.html"
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append("<!doctype html><html><head><meta charset=`"utf-8`"><title>$(HtmlEnc $bundleTitle)</title><style>")
    [void]$sb.Append('body{font-family:"Segoe UI Variable","Segoe UI",sans-serif;background:#fafafa;color:#242424;margin:0;padding:32px}main{max-width:960px;margin:auto}')
    [void]$sb.Append('h1{font-size:28px;font-weight:600;margin:0 0 4px}.sub{color:#616161;margin-bottom:24px}section{background:#fff;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,.14);padding:16px 20px;margin-bottom:16px}')
    [void]$sb.Append('h2{font-size:18px;font-weight:600;margin:0 0 8px}ul{margin:0;padding-left:20px}li{margin:4px 0}a{color:#0f6cbd;text-decoration:none}a:hover{text-decoration:underline}ol li{margin:6px 0}</style></head><body><main>')
    [void]$sb.Append("<h1>$(HtmlEnc $bundleTitle)</h1><div class=`"sub`">$(HtmlEnc (T 'Demo resources' 'Demo-Ressourcen')) &middot; $(HtmlEnc ((Get-Date).ToString('yyyy-MM-dd HH:mm')))</div>")
    foreach ($g in $groups) {
        [void]$sb.Append("<section><h2>$(HtmlEnc $g.name)</h2><ul>")
        foreach ($l in $g.links) { [void]$sb.Append("<li><a href=`"$(HtmlEnc $l.url)`" target=`"_blank`">$(HtmlEnc $l.name)</a></li>") }
        [void]$sb.Append('</ul></section>')
    }
    if ($links.manual -and @($links.manual).Count) {
        [void]$sb.Append("<section><h2>$(HtmlEnc (T 'Still to do by hand' 'Noch von Hand'))</h2><ol>")
        foreach ($m in $links.manual) { [void]$sb.Append("<li>$(HtmlEnc $m)</li>") }
        [void]$sb.Append('</ol></section>')
    }
    [void]$sb.Append('</main></body></html>')
    [IO.File]::WriteAllText($file, $sb.ToString(), (New-Object Text.UTF8Encoding $false))
    $file
}
function Set-EdgeFavorites($groups, $page) {
    $admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $admin) {
        Write-Host ("  " + (T 'Edge favorites: skipped (run the terminal as administrator to get them).' 'Edge-Favoriten: uebersprungen (Terminal als Administrator starten, dann kommen sie mit).')) -ForegroundColor DarkGray
        return
    }
    $children = @(@{ name = (T 'Link page (Desktop)' 'Link-Seite (Desktop)'); url = (FileUri $page) })
    foreach ($g in $groups) {
        $l = @($g.links | ForEach-Object { @{ name = $_.name; url = $_.url } })
        if ($l.Count) { $children += @{ name = $g.name; children = $l } }
    }
    # one subfolder per customer bundle; other bundles set up earlier on this PC are kept
    $store = Join-Path $env:LOCALAPPDATA 'DemoKit\edge-favorites.json'
    $all = [ordered]@{}
    if (Test-Path $store) { try { (Get-Content $store -Raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $all[$_.Name] = $_.Value } } catch { } }
    $all[$bundleName] = @{ name = $bundleTitle; children = $children }
    New-Item -ItemType Directory -Force (Split-Path $store) | Out-Null
    [IO.File]::WriteAllText($store, (ConvertTo-Json $all -Depth 10))
    $value = @(@{ toplevel_name = 'Demo Kit' }) + @($all.Keys | Sort-Object | ForEach-Object { $all[$_] })
    $key = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    Set-ItemProperty -Path $key -Name ManagedFavorites -Value (ConvertTo-Json $value -Depth 10 -Compress) -Type String
    Write-Host ("  " + (T 'Edge favorites: folder "Demo Kit" > ' 'Edge-Favoriten: Ordner "Demo Kit" > ') + $bundleTitle) -ForegroundColor Green
}
function Publish-Links {
    if ($WhatIf -or ($Bundle -and $Demo)) { return }
    try {
        $groups = Get-LinkGroups
        $page = Write-LinkPage $groups
        Write-Host ""
        Write-Host ("  " + (T 'Links: ' 'Links: ') + $page) -ForegroundColor Green
        Set-EdgeFavorites $groups $page
    } catch { Write-Host ("  " + (T 'Link page / Edge favorites failed: ' 'Link-Seite / Edge-Favoriten fehlgeschlagen: ') + $_) -ForegroundColor Yellow }
}
foreach ($d in $demos) {
    try { $manifest = Invoke-RestMethod -Uri (Url "$d/manifest.json") -UseBasicParsing }
    catch { throw "Demo '$d' not found in $Repo ($Branch). / Demo '$d' nicht gefunden." }

    $dir = if ($Bundle) { Join-Path $root $d } else { Join-Path $root "Demo-$d" }
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    Write-Host ""
    Write-Host ("  " + (Loc $manifest.title)) -ForegroundColor Cyan
    Write-Host "  -> $dir" -ForegroundColor DarkGray

    $stats = @{ new = 0; updated = 0; same = 0 }
    $rels = @()
    foreach ($f in @($manifest.files)) {
        if ($f -is [string]) { Install-File "$d/$f" $f $null $dir $stats; $rels += $f }
        else { Install-File "$d/$($f.path)" $f.path $f.sha256 $dir $stats; $rels += $f.path }
    }
    foreach ($f in @($manifest.shared)) {
        if ($null -eq $f) { continue }
        Install-File "$($f.from)/$($f.path)" $f.path $f.sha256 $dir $stats; $rels += $f.path
    }
    if ($Tenant) {
        # The -Tenant cache only mirrors the manifest: drop files that were renamed or removed in the repo
        $keep = @{}; foreach ($r in $rels) { $keep[($r -replace '/', '\').ToLowerInvariant()] = $true }
        Get-ChildItem $dir -Recurse -File | Where-Object { -not $keep[$_.FullName.Substring($dir.Length + 1).ToLowerInvariant()] } | ForEach-Object {
            Remove-Item -LiteralPath $_.FullName -Force
            Write-Host ("    {0,-13} {1}" -f (T 'removed' 'entfernt'), ($_.FullName.Substring($dir.Length + 1) -replace '\\', '/')) -ForegroundColor DarkGray
        }
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
    $items += @{ id = $d; dir = $dir; drivePath = $drivePath; files = $rels }
    $links.demos[$d] = @{ title = (Loc $manifest.title); dir = $dir; local = @($rels) }
    try {
        $tj = Invoke-RestMethod -Uri (Url "$d/tenant.json") -UseBasicParsing
        if ($tj -is [string]) { $tj = $tj.TrimStart([char]0xFEFF) | ConvertFrom-Json }
        $links.demos[$d].extra = @(foreach ($l in @($tj.links)) { if ($l -and $l.url) { @{ name = $(if ($Language -eq 'de' -and $l.name.de) { $l.name.de } else { $l.name.en }); url = $l.url } } })
    } catch { }
}

Write-Host ""
$summary | Format-Table -AutoSize | Out-String | Write-Host
$bundleName = if ($Bundle) { $bundleInfo.folder } else { "Demo-$Demo" }
$bundleTitle = if ($Bundle) { Loc $bundleInfo.title } else { $links.demos[$Demo].title }
$localRoot = if ($Bundle) { $root } else { Join-Path $root "Demo-$Demo" }
if ($Bundle -and -not $Demo) {
    Write-Host (T '  Next steps per demo: see the Demo Kit page of each demo.' '  Naechste Schritte je Demo: siehe die jeweilige Seite im Demo Kit.') -ForegroundColor Cyan
    Write-Host ""
}

if ($Tenant) {
    try {
        $code = (Invoke-WebRequest -Uri (Url 'tenant.ps1') -UseBasicParsing).Content
        if ($code -is [byte[]]) { $code = [Text.Encoding]::UTF8.GetString($code) }
        & ([scriptblock]::Create($code)) -Items $items -Raw $raw -Language $Language -WhatIf:$WhatIf -Links $links
    } catch {
        Write-Host ("  " + (T 'Tenant step failed: ' 'Tenant-Schritt fehlgeschlagen: ') + $_) -ForegroundColor Red
        Write-Host ($_.ScriptStackTrace) -ForegroundColor DarkGray
        $failed = $true
    }
    Publish-Links
    return
}

Publish-Links

$open = if ($Bundle -and -not $Demo) { $root } elseif ($Bundle) { Join-Path $root $Demo } else { Join-Path $root "Demo-$Demo" }
if (-not $NoExplorer) { Start-Process explorer.exe -ArgumentList "`"$open`"" }
} catch {
    $failed = $true
    Write-Host ""
    Write-Host ("  " + (T 'Error: ' 'Fehler: ') + $_) -ForegroundColor Red
    Write-Host ($_ | Format-List * -Force | Out-String) -ForegroundColor DarkGray
    Write-Host ($_.ScriptStackTrace) -ForegroundColor DarkGray
} finally {
    if ($log) {
        Write-Host ""
        if ($failed) { Write-Host ("  " + (T 'Something went wrong - please send this log file: ' 'Etwas ist schiefgelaufen - bitte diese Log-Datei schicken: ') + $log) -ForegroundColor Yellow }
        else { Write-Host ("  Log: " + $log) -ForegroundColor DarkGray }
        try { Stop-Transcript | Out-Null } catch { }
    }
}
