<#
.SYNOPSIS
  Sets up ALL demos of the Demo Kit on the demo machine / in the demo tenant - one fixed command.

.DESCRIPTION
  The command (Windows PowerShell 5.1 and PowerShell 7, no modules needed for the files):

    & ([scriptblock]::Create((irm https://raw.githubusercontent.com/jenssgb/demo-kit-data/main/install.ps1))) -Tenant

  catalog.json (generated from the kit decks by tools\Update-Catalog.ps1) lists every app and its demos.
  Each demo lands in Demo-Kit\<App>\<demo-id>. New demos in the catalog are picked up automatically:
  running it again adds new and changed files and skips unchanged ones (SHA-256). Every demo folder is
  self-contained: files shared with other demos ("shared" in the manifest) are copied too.

  -Tenant (demo VM, signed in as the demo admin, e.g. MOD Administrator in a CDX tenant):
    Files are cached in %LOCALAPPDATA%\DemoKit and uploaded to the signed-in user's OneDrive (Demo-Kit/<App>/<demo-id>)
    with Microsoft Graph; then tenant.ps1 creates what the demos need in the tenant (mails, meetings, OneNote, see
    <demo>/tenant.json) and prints the steps that have to be done by hand.
  Without -Tenant the files are copied to OneDrive\Demo-Kit on this PC (fallback: Desktop).
  -Language de|en   mail/meeting/OneNote language (default en)
  -WhatIf           shows what would happen without changing the tenant
  -RemoveLegacy     deletes the folders of older kit versions (catalog.json "legacy", e.g. Demo-BPW) from OneDrive,
                    the local cache and the old Desktop link pages, so Copilot doesn't find duplicate files
  -Demo <id>[,<id>] troubleshooting only: just these demos (no link page)
  Afterwards: Desktop\Demo-Kit - Links.html (all links by app) and, in an elevated terminal, the Edge favorites
  folder "Demo Kit". Every run writes a log to Desktop\DemoKit-Logs (last 20 kept) - send that file when something
  goes wrong. Demo people are the real users of the CDX demo tenant (Teresa Sac, Vance DeLeon, ...).
#>
param(
    [string[]]$Demo,
    [string]$Repo = 'jenssgb/demo-kit-data',
    [string]$Branch = 'main',
    [string]$Target,
    [switch]$NoExplorer,
    [switch]$Tenant,
    [ValidateSet('en', 'de')] [string]$Language = 'en',
    [switch]$WhatIf,
    [switch]$RemoveLegacy
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$raw = "https://raw.githubusercontent.com/$Repo/$Branch"
$de = (Get-Culture).TwoLetterISOLanguageName -eq 'de'
function T($en, $deText) { if ($de) { $deText } else { $en } }
function Loc($o) { if ($null -eq $o) { return $null }; if ($o -is [string]) { return $o }; if ($de -and $o.de) { $o.de } else { $o.en } }
function Url($path) { "$raw/" + ((($path -split '/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/') }
function Get-Json($path) {
    $r = Invoke-RestMethod -Uri (Url $path) -UseBasicParsing
    if ($r -is [string]) { $r = $r.TrimStart([char]0xFEFF) | ConvertFrom-Json }
    $r
}
$Demo = @($Demo | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

# ---------- log on the Desktop (every run) ----------
$log = $null; $failed = $false
try {
    $logDir = Join-Path ([Environment]::GetFolderPath('Desktop')) 'DemoKit-Logs'
    New-Item -ItemType Directory -Force -Path $logDir | Out-Null
    $log = Join-Path $logDir ("demo-kit{0}-{1:yyyyMMdd-HHmmss}.log" -f $(if ($Tenant) { '-tenant' } else { '' }), (Get-Date))
    Get-ChildItem $logDir -Filter *.log | Sort-Object LastWriteTime -Descending | Select-Object -Skip 19 | Remove-Item -Force -ErrorAction SilentlyContinue
    Start-Transcript -Path $log | Out-Null
    $commit = try { (Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/commits/$Branch" -UseBasicParsing -TimeoutSec 5).sha.Substring(0, 7) } catch { '?' }
    $os = try { (Get-CimInstance Win32_OperatingSystem).Caption } catch { [Environment]::OSVersion.VersionString }
    Write-Host ("  Demo Kit data | {0:yyyy-MM-dd HH:mm:ss} | {1} | {2}\{3} | {4} | PowerShell {5} | {6}@{7} ({8})" -f (Get-Date), $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERNAME, $os, $PSVersionTable.PSVersion, $Repo, $Branch, $commit) -ForegroundColor DarkGray
    Write-Host ("  Demo={0} Tenant={1} Language={2} WhatIf={3} RemoveLegacy={4} Target={5}" -f ($Demo -join ','), [bool]$Tenant, $Language, [bool]$WhatIf, [bool]$RemoveLegacy, $Target) -ForegroundColor DarkGray
} catch { $log = $null }

try {
if ($Tenant -and -not $Target) { $Target = Join-Path $env:LOCALAPPDATA 'DemoKit' }

if (-not $Target) {
    $base = @($env:OneDriveCommercial, $env:OneDrive) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $base) { $base = [Environment]::GetFolderPath('Desktop') }
} else { $base = $Target }

# ---------- catalog: every app and its demos ----------
try { $catalog = Get-Json 'catalog.json' }
catch { throw "catalog.json not found in $Repo ($Branch). / catalog.json nicht gefunden." }
$kit = $catalog.folder
$kitTitle = Loc $catalog.title
$root = Join-Path $base $kit
$entries = @(foreach ($a in @($catalog.apps)) {
    foreach ($id in @($a.demos)) { [pscustomobject]@{ id = $id; app = $a.folder; appTitle = (Loc $a.title) } }
})
if ($Demo.Count) {
    $unknown = @($Demo | Where-Object { $_ -notin $entries.id })
    if ($unknown.Count) { throw ("Unknown demo: {0} (see catalog.json). / Unbekannte Demo: {0}" -f ($unknown -join ', ')) }
    $entries = @($entries | Where-Object { $_.id -in $Demo })
}
Write-Host ""
Write-Host ("  {0}: {1} {2}, {3} {4}" -f $kitTitle, @($catalog.apps).Count, (T 'apps' 'Apps'), $entries.Count, (T 'demos' 'Demos')) -ForegroundColor Cyan
Write-Host "  -> $root"

# ---------- folders of older kit versions (Demo-<customer>) ----------
$desktop = [Environment]::GetFolderPath('Desktop')
foreach ($lg in @($catalog.legacy)) {
    if (-not $lg) { continue }
    $old = @((Join-Path $base $lg), (Join-Path $desktop "$lg - Links.html"))
    if ($Tenant) { $old += Join-Path (Join-Path $env:LOCALAPPDATA 'DemoKit') $lg }
    foreach ($p in ($old | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $p)) { continue }
        if (-not $RemoveLegacy) {
            Write-Host ("  " + (T 'Old kit folder found: {0} - run again with -RemoveLegacy to delete it.' 'Alter Kit-Ordner gefunden: {0} - mit -RemoveLegacy erneut ausfuehren, um ihn zu loeschen.') -f $p) -ForegroundColor Yellow
        } elseif ($WhatIf) {
            Write-Host ("    {0,-13} {1}" -f (T 'would delete' 'wuerde loeschen'), $p) -ForegroundColor Yellow
        } else {
            Remove-Item -LiteralPath $p -Recurse -Force
            Write-Host ("    {0,-13} {1}" -f (T 'deleted' 'geloescht'), $p) -ForegroundColor DarkGray
        }
    }
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
    # "Start" with general links, then one group per app: @{ name; demos = @(@{ name; links = @(@{ name; url }) }) }
    $gen = @(@{ name = 'Microsoft Copilot'; url = 'https://copilot.cloud.microsoft/' })
    if ($links.root) {
        $gen += @{ name = 'Outlook'; url = 'https://outlook.office.com/mail/' }
        $gen += @{ name = (T 'OneDrive folder' 'OneDrive-Ordner') + " $kit"; url = $links.root }
    }
    $gen += @{ name = (T 'Local folder' 'Lokaler Ordner') + " $kit"; url = (FileUri $root) }
    $apps = [ordered]@{}
    foreach ($id in $links.demos.Keys) {
        $d = $links.demos[$id]
        $l = @()
        if ($d.web) { $l += @{ name = (T 'OneDrive folder' 'OneDrive-Ordner'); url = $d.web } }
        if ($d.onenote) { $l += $d.onenote }
        foreach ($x in @($d.extra)) { if ($x) { $l += $x } }
        if ($d.files -and @($d.files).Count) { foreach ($f in $d.files) { if ($f.url) { $l += @{ name = $f.name; url = $f.url } } } }
        else { foreach ($f in @($d.local)) { $l += @{ name = $f; url = (FileUri (Join-Path $d.dir ($f -replace '/', '\'))) } } }
        if (-not $apps.Contains($d.app)) { $apps[$d.app] = New-Object Collections.ArrayList }
        [void]$apps[$d.app].Add(@{ name = $d.title; links = $l })
    }
    $groups = @(@{ name = 'Start'; links = $gen })
    foreach ($a in $apps.Keys) { $groups += @{ name = $a; demos = @($apps[$a]) } }
    $groups
}
function Write-LinkPage($groups) {
    $file = Join-Path $desktop "$kit - Links.html"
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append("<!doctype html><html><head><meta charset=`"utf-8`"><title>$(HtmlEnc $kitTitle)</title><style>")
    [void]$sb.Append('body{font-family:"Segoe UI Variable","Segoe UI",sans-serif;background:#fafafa;color:#242424;margin:0;padding:32px}main{max-width:960px;margin:auto}')
    [void]$sb.Append('h1{font-size:28px;font-weight:600;margin:0 0 4px}.sub{color:#616161;margin-bottom:24px}section{background:#fff;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,.14);padding:16px 20px;margin-bottom:16px}')
    [void]$sb.Append('h2{font-size:20px;font-weight:600;margin:28px 0 12px}h3{font-size:16px;font-weight:600;margin:0 0 8px}ul{margin:0;padding-left:20px}li{margin:4px 0}a{color:#0f6cbd;text-decoration:none}a:hover{text-decoration:underline}ol li{margin:6px 0}</style></head><body><main>')
    [void]$sb.Append("<h1>$(HtmlEnc $kitTitle)</h1><div class=`"sub`">$(HtmlEnc (T 'Demo resources' 'Demo-Ressourcen')) &middot; $(HtmlEnc ((Get-Date).ToString('yyyy-MM-dd HH:mm')))</div>")
    foreach ($g in $groups) {
        $blocks = if ($g.demos) { [void]$sb.Append("<h2>$(HtmlEnc $g.name)</h2>"); $g.demos } else { @($g) }
        foreach ($b in $blocks) {
            [void]$sb.Append("<section><h3>$(HtmlEnc $b.name)</h3><ul>")
            foreach ($l in $b.links) { [void]$sb.Append("<li><a href=`"$(HtmlEnc $l.url)`" target=`"_blank`">$(HtmlEnc $l.name)</a></li>") }
            [void]$sb.Append('</ul></section>')
        }
    }
    if ($links.manual -and @($links.manual).Count) {
        [void]$sb.Append("<h2>$(HtmlEnc (T 'Still to do by hand' 'Noch von Hand'))</h2><section><ol>")
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
    # Demo Kit > link page, Start > links, <App> > <demo> > links (replaces the per-customer folders of older versions)
    $value = @(@{ toplevel_name = 'Demo Kit' }, @{ name = (T 'Link page (Desktop)' 'Link-Seite (Desktop)'); url = (FileUri $page) })
    foreach ($g in $groups) {
        if ($g.demos) {
            $sub = @(foreach ($d in $g.demos) {
                $l = @($d.links | ForEach-Object { @{ name = $_.name; url = $_.url } })
                if ($l.Count) { @{ name = $d.name; children = $l } }
            })
            if ($sub.Count) { $value += @{ name = $g.name; children = $sub } }
        } else {
            $l = @($g.links | ForEach-Object { @{ name = $_.name; url = $_.url } })
            if ($l.Count) { $value += @{ name = $g.name; children = $l } }
        }
    }
    $key = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    Set-ItemProperty -Path $key -Name ManagedFavorites -Value (ConvertTo-Json $value -Depth 10 -Compress) -Type String
    Remove-Item (Join-Path $env:LOCALAPPDATA 'DemoKit\edge-favorites.json') -Force -ErrorAction SilentlyContinue
    Write-Host ("  " + (T 'Edge favorites: folder "Demo Kit"' 'Edge-Favoriten: Ordner "Demo Kit"')) -ForegroundColor Green
}
function Publish-Links {
    if ($WhatIf -or $Demo.Count) { return }
    try {
        $groups = Get-LinkGroups
        $page = Write-LinkPage $groups
        Write-Host ""
        Write-Host ("  " + (T 'Links: ' 'Links: ') + $page) -ForegroundColor Green
        Set-EdgeFavorites $groups $page
    } catch { Write-Host ("  " + (T 'Link page / Edge favorites failed: ' 'Link-Seite / Edge-Favoriten fehlgeschlagen: ') + $_) -ForegroundColor Yellow }
}

$lastApp = $null
foreach ($e in $entries) {
    $d = $e.id
    try { $manifest = Get-Json "$d/manifest.json" }
    catch { throw "Demo '$d' not found in $Repo ($Branch). / Demo '$d' nicht gefunden." }

    if ($e.app -ne $lastApp) { Write-Host ""; Write-Host ("  == {0} ==" -f $e.appTitle) -ForegroundColor Cyan; $lastApp = $e.app }
    $dir = Join-Path (Join-Path $root $e.app) $d
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
    if ($next -and $Demo.Count) {
        Write-Host ""
        Write-Host (T '  Next steps:' '  Naechste Schritte:') -ForegroundColor Cyan
        $n = 0
        foreach ($step in $next) { $n++; Write-Host "  $n. $step" }
    }
    $summary += [pscustomobject]@{ App = $e.app; Demo = $d; New = $stats.new; Updated = $stats.updated; Unchanged = $stats.same }
    $items += @{ id = $d; dir = $dir; drivePath = "$kit/$($e.app)/$d"; files = $rels }
    $links.demos[$d] = @{ title = (Loc $manifest.title); app = $e.appTitle; dir = $dir; local = @($rels) }
    try {
        $tj = Get-Json "$d/tenant.json"
        $links.demos[$d].extra = @(foreach ($l in @($tj.links)) { if ($l -and $l.url) { @{ name = $(if ($Language -eq 'de' -and $l.name.de) { $l.name.de } else { $l.name.en }); url = $l.url } } })
    } catch { }
}

Write-Host ""
$summary | Format-Table -AutoSize | Out-String | Write-Host
if (-not $Demo.Count) {
    Write-Host (T '  Next steps per demo: see "Agenda & setup" of each app deck in the Demo Kit.' '  Naechste Schritte je Demo: siehe "Agenda & Setup" im jeweiligen App-Deck im Demo Kit.') -ForegroundColor Cyan
    Write-Host ""
}

if ($Tenant) {
    try {
        $code = (Invoke-WebRequest -Uri (Url 'tenant.ps1') -UseBasicParsing).Content
        if ($code -is [byte[]]) { $code = [Text.Encoding]::UTF8.GetString($code) }
        & ([scriptblock]::Create($code)) -Items $items -Raw $raw -Language $Language -WhatIf:$WhatIf -Links $links -Legacy @($catalog.legacy) -RemoveLegacy:$RemoveLegacy
    } catch {
        Write-Host ("  " + (T 'Tenant step failed: ' 'Tenant-Schritt fehlgeschlagen: ') + $_) -ForegroundColor Red
        Write-Host ($_.ScriptStackTrace) -ForegroundColor DarkGray
        $failed = $true
    }
    Publish-Links
    return
}

Publish-Links

if (-not $NoExplorer) { Start-Process explorer.exe -ArgumentList "`"$root`"" }
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
