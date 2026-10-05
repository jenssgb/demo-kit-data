<#
.SYNOPSIS
  Puts the demo data straight into the Microsoft 365 tenant of the signed-in user.
  Called by install.ps1 -Tenant (not meant to be run on its own).

.DESCRIPTION
  Official ways only (Microsoft Learn):
    files     Microsoft Graph PUT /me/drive/root:/<path>:/content, unchanged files skipped by quickXorHash
    mails     Graph POST /me/sendMail with "from" = sender (delegated Mail.Send.Shared + Exchange "Send As")
              -> real received mails (POST /me/messages would only create drafts)
    senders   existing users by display name, otherwise a shared mailbox (Exchange Online PowerShell,
              New-Mailbox -Shared, Add-RecipientPermission -AccessRights SendAs)
    meetings  Graph POST /me/events
    OneNote   Graph /me/onenote (delegated, app-only isn't supported)
  Each demo describes what it needs in <demo>/tenant.json:
    mails    ["mails-{lang}/x.eml", ...]   (local demo folder first, then the repo)
    events   [{ subject{de,en}, weekday, start "HH:mm", minutes, attendees["Display Name"], body{de,en} }]
    onenote  { notebook, section, pages[{ title{de,en}, file "tenant/pages/x-{lang}.html" }] }
    manual   { de[], en[] }   steps that can't be scripted, printed as a checklist at the end
  Running it again creates nothing twice.
  Keep this file ASCII-only (it is loaded with Invoke-RestMethod on Windows PowerShell 5.1).
#>
param(
    [Parameter(Mandatory)] [object[]]$Items,   # @{ id; dir; drivePath }
    [Parameter(Mandatory)] [string]$Raw,
    [ValidateSet('en', 'de')] [string]$Language = 'en',
    [switch]$WhatIf,
    [int]$SendAsWaitMinutes = 20,
    [switch]$SelfTest
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$de = (Get-Culture).TwoLetterISOLanguageName -eq 'de'
function T($en, $deText) { if ($de) { $deText } else { $en } }
function L($o) { if ($null -eq $o) { return $null }; if ($o -is [string]) { return $o }; if ($Language -eq 'de' -and $o.de) { $o.de } else { $o.en } }
function Url($path) { "$Raw/" + ((($path -split '/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/') }
function Say($text, $color = 'Gray') { Write-Host "  $text" -ForegroundColor $color }
function Row($tag, $text, $color) { Write-Host ("    {0,-13} {1}" -f $tag, $text) -ForegroundColor $color }

# ---------- helpers ----------
if (-not ('DemoKit.QuickXorHash' -as [type])) {
    # Microsoft's reference implementation: https://learn.microsoft.com/onedrive/developer/code-snippets/quickxorhash
    Add-Type -TypeDefinition @'
using System;
namespace DemoKit {
public class QuickXorHash : System.Security.Cryptography.HashAlgorithm {
    private const int BitsInLastCell = 32; private const byte Shift = 11; private const byte WidthInBits = 160;
    private UInt64[] _data; private Int64 _lengthSoFar; private int _shiftSoFar;
    public QuickXorHash() { this.Initialize(); }
    protected override void HashCore(byte[] array, int ibStart, int cbSize) {
        unchecked {
            int currentShift = this._shiftSoFar; int vectorArrayIndex = currentShift / 64; int vectorOffset = currentShift % 64;
            int iterations = Math.Min(cbSize, WidthInBits);
            for (int i = 0; i < iterations; i++) {
                bool isLastCell = vectorArrayIndex == this._data.Length - 1;
                int bitsInVectorCell = isLastCell ? BitsInLastCell : 64;
                if (vectorOffset <= bitsInVectorCell - 8) {
                    for (int j = ibStart + i; j < cbSize + ibStart; j += WidthInBits) { this._data[vectorArrayIndex] ^= (ulong)array[j] << vectorOffset; }
                } else {
                    int index1 = vectorArrayIndex; int index2 = isLastCell ? 0 : (vectorArrayIndex + 1);
                    byte low = (byte)(bitsInVectorCell - vectorOffset); byte xoredByte = 0;
                    for (int j = ibStart + i; j < cbSize + ibStart; j += WidthInBits) { xoredByte ^= array[j]; }
                    this._data[index1] ^= (ulong)xoredByte << vectorOffset; this._data[index2] ^= (ulong)xoredByte >> low;
                }
                vectorOffset += Shift;
                while (vectorOffset >= bitsInVectorCell) { vectorArrayIndex = isLastCell ? 0 : vectorArrayIndex + 1; vectorOffset -= bitsInVectorCell; }
            }
            this._shiftSoFar = (this._shiftSoFar + Shift * (cbSize % WidthInBits)) % WidthInBits;
        }
        this._lengthSoFar += cbSize;
    }
    protected override byte[] HashFinal() {
        byte[] rgb = new byte[(WidthInBits - 1) / 8 + 1];
        for (Int32 i = 0; i < this._data.Length - 1; i++) { Buffer.BlockCopy(BitConverter.GetBytes(this._data[i]), 0, rgb, i * 8, 8); }
        Buffer.BlockCopy(BitConverter.GetBytes(this._data[this._data.Length - 1]), 0, rgb, (this._data.Length - 1) * 8, rgb.Length - (this._data.Length - 1) * 8);
        var lengthBytes = BitConverter.GetBytes(this._lengthSoFar);
        for (int i = 0; i < lengthBytes.Length; i++) { rgb[(WidthInBits / 8) - lengthBytes.Length + i] ^= lengthBytes[i]; }
        return rgb;
    }
    public override sealed void Initialize() { this._data = new ulong[(WidthInBits - 1) / 64 + 1]; this._shiftSoFar = 0; this._lengthSoFar = 0; }
    public override int HashSize { get { return WidthInBits; } }
}
}
'@
}
function Get-QuickXor([string]$file) {
    $h = New-Object DemoKit.QuickXorHash
    $s = [IO.File]::OpenRead($file)
    try { [Convert]::ToBase64String($h.ComputeHash($s)) } finally { $s.Dispose(); $h.Dispose() }
}

# Non-ASCII as escapes, so the request encoding never matters.
function To-AsciiJson($obj) {
    $json = $obj | ConvertTo-Json -Depth 10 -Compress
    [regex]::Replace($json, '[^\x00-\x7F]', { param($m) '\u{0:x4}' -f [int][char]$m.Value })
}
function To-AsciiHtml([string]$html) {
    $sb = New-Object Text.StringBuilder
    for ($i = 0; $i -lt $html.Length; $i++) {
        $c = $html[$i]
        if ([char]::IsHighSurrogate($c) -and $i + 1 -lt $html.Length) {
            [void]$sb.Append('&#' + [char]::ConvertToUtf32($c, $html[$i + 1]) + ';'); $i++
        } elseif ([int]$c -gt 127) { [void]$sb.Append('&#' + [int]$c + ';') } else { [void]$sb.Append($c) }
    }
    $sb.ToString()
}

function Decode-Words([string]$s) {
    $s = [regex]::Replace($s, '\?=\s+=\?', '?==?')
    [regex]::Replace($s, '=\?([^?]+)\?([bBqQ])\?([^?]*)\?=', {
        param($m)
        $enc = [Text.Encoding]::GetEncoding($m.Groups[1].Value)
        if ($m.Groups[2].Value -in 'b', 'B') { return $enc.GetString([Convert]::FromBase64String($m.Groups[3].Value)) }
        $bytes = New-Object Collections.Generic.List[byte]
        $t = $m.Groups[3].Value -replace '_', ' '
        for ($i = 0; $i -lt $t.Length; $i++) {
            if ($t[$i] -eq '=' -and $i + 2 -lt $t.Length) { $bytes.Add([Convert]::ToByte($t.Substring($i + 1, 2), 16)); $i += 2 }
            else { $bytes.Add([byte][char]$t[$i]) }
        }
        $enc.GetString($bytes.ToArray())
    })
}

# Plain-text EML as written by the demo repo (headers, text/plain body, QP/base64/8bit).
function Read-Eml([byte[]]$bytes) {
    $text = [Text.Encoding]::GetEncoding(28591).GetString($bytes)   # byte-transparent
    $split = [regex]::Match($text, '\r?\n\r?\n')
    $head = $text.Substring(0, $split.Index) -replace '\r?\n[ \t]+', ' '
    $body = $text.Substring($split.Index + $split.Length)
    $h = @{}
    foreach ($line in $head -split '\r?\n') { if ($line -match '^([\w-]+):\s?(.*)$') { $h[$Matches[1].ToLowerInvariant()] = $Matches[2] } }
    $charset = 'utf-8'
    if ($h['content-type'] -match 'charset="?([\w-]+)') { $charset = $Matches[1] }
    $enc = [Text.Encoding]::GetEncoding($charset)
    $cte = "$($h['content-transfer-encoding'])".ToLowerInvariant()
    if ($cte -eq 'quoted-printable') {
        $body = $body -replace '=\r?\n', ''
        $out = New-Object Collections.Generic.List[byte]
        for ($i = 0; $i -lt $body.Length; $i++) {
            if ($body[$i] -eq '=' -and $i + 2 -lt $body.Length -and $body.Substring($i + 1, 2) -match '^[0-9A-Fa-f]{2}$') {
                $out.Add([Convert]::ToByte($body.Substring($i + 1, 2), 16)); $i += 2
            } else { $out.Add([byte][char]$body[$i]) }
        }
        $bodyText = $enc.GetString($out.ToArray())
    } elseif ($cte -eq 'base64') { $bodyText = $enc.GetString([Convert]::FromBase64String(($body -replace '\s', ''))) }
    else { $bodyText = $enc.GetString([Text.Encoding]::GetEncoding(28591).GetBytes($body)) }
    $from = Decode-Words "$($h['from'])"
    $name = if ($from -match '^\s*"?([^"<]+?)"?\s*<') { $Matches[1].Trim() } else { $from.Trim() }
    [pscustomobject]@{ Subject = (Decode-Words "$($h['subject'])").Trim(); FromName = $name; Body = ($bodyText -replace '\r?\n', "`r`n").TrimEnd() }
}

function Get-Resource($item, [string]$rel) {
    $rel = $rel.Replace('{lang}', $Language)
    $local = Join-Path $item.dir ($rel -replace '/', '\')
    if (Test-Path $local) { return [IO.File]::ReadAllBytes($local) }
    $tmp = [IO.Path]::GetTempFileName()
    try {
        $src = "$($item.id)/$rel"
        $mf = Join-Path $item.dir 'manifest.json'
        $manifest = if (Test-Path $mf) { [IO.File]::ReadAllText($mf) | ConvertFrom-Json } else { Invoke-RestMethod -Uri (Url "$($item.id)/manifest.json") -UseBasicParsing }
        $shared = @($manifest.shared) | Where-Object { $_ -and $_.path -eq $rel } | Select-Object -First 1
        if ($shared) {
            $own = Join-Path (Join-Path (Split-Path $item.dir -Parent) $shared.from) ($rel -replace '/', '\')
            if (Test-Path $own) { return [IO.File]::ReadAllBytes($own) }
            $src = "$($shared.from)/$rel"
        }
        Invoke-WebRequest -Uri (Url $src) -OutFile $tmp -UseBasicParsing; [IO.File]::ReadAllBytes($tmp)
    }
    finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
}

function Get-TenantConfig($item) {
    $local = Join-Path $item.dir 'tenant.json'
    if (Test-Path $local) { return ([IO.File]::ReadAllText($local, [Text.Encoding]::UTF8).TrimStart([char]0xFEFF) | ConvertFrom-Json) }
    try { $r = Invoke-WebRequest -Uri (Url "$($item.id)/tenant.json") -UseBasicParsing } catch { return $null }
    $txt = if ($r.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($r.Content) } else { [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray()) }
    $txt.TrimStart([char]0xFEFF) | ConvertFrom-Json
}

function Next-Weekday([string]$day) {
    $target = [DayOfWeek]$day
    $d = (Get-Date).Date.AddDays(1)
    while ($d.DayOfWeek -ne $target) { $d = $d.AddDays(1) }
    $d
}

# ---------- load what the demos need ----------
$configs = @()
foreach ($it in $Items) {
    $c = Get-TenantConfig $it
    if ($c) { $configs += [pscustomobject]@{ Item = $it; Config = $c } }
}
$mails = @()
foreach ($c in $configs) {
    foreach ($m in @($c.Config.mails)) {
        if (-not $m) { continue }
        try { $eml = Read-Eml (Get-Resource $c.Item $m) }
        catch { Say ((T 'Mail {0} ({1}) not found: {2}' 'Mail {0} ({1}) nicht gefunden: {2}') -f $m, $c.Item.id, $_) Red; continue }
        if (-not ($mails | Where-Object { $_.Subject -eq $eml.Subject })) { $mails += $eml | Add-Member -PassThru NoteProperty Demo $c.Item.id }
    }
}
$people = @($mails | ForEach-Object FromName) + @($configs | ForEach-Object { $_.Config.events } | Where-Object { $_ } | ForEach-Object { $_.attendees }) |
    Where-Object { $_ } | Sort-Object -Unique

if ($SelfTest) {
    $e = New-Object DemoKit.QuickXorHash
    Say ("quickXorHash(empty) = " + [Convert]::ToBase64String($e.ComputeHash([byte[]]@())))
    foreach ($c in $configs) { Say ("{0}: mails {1}, events {2}, onenote {3}, manual {4}" -f $c.Item.id, @($c.Config.mails | Where-Object { $_ }).Count, @($c.Config.events | Where-Object { $_ }).Count, [bool]$c.Config.onenote, @((L $c.Config.manual)).Count) }
    foreach ($m in $mails) { Say ("mail [{0}] {1} | {2} | {3} chars" -f $m.Demo, $m.FromName, $m.Subject, $m.Body.Length) }
    Say ("people: " + ($people -join ', '))
    foreach ($c in $configs) {
        if (-not $c.Config.onenote) { continue }
        foreach ($p in $c.Config.onenote.pages) { $html = [Text.Encoding]::UTF8.GetString((Get-Resource $c.Item $p.file)); Say ("page {0}: {1} chars, ascii {2}" -f (L $p.title), $html.Length, (To-AsciiHtml $html).Length) }
    }
    foreach ($it in $Items) {
        if (Test-Path $it.dir) { Get-ChildItem $it.dir -Recurse -File | Select-Object -First 2 | ForEach-Object { Say ("hash {0} {1}" -f $_.Name, (Get-QuickXor $_.FullName)) } }
    }
    return
}

# ---------- sign in ----------
Write-Host ""
Say (T 'Tenant: signing in (Microsoft Graph) ...' 'Tenant: Anmeldung (Microsoft Graph) ...') Cyan
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    Say (T 'Installing module Microsoft.Graph.Authentication (current user) ...' 'Installiere Modul Microsoft.Graph.Authentication (aktueller Benutzer) ...') DarkGray
    if ($PSVersionTable.PSEdition -ne 'Core') { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null }
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph.Authentication
$scopes = @('User.Read', 'Files.ReadWrite', 'Mail.ReadBasic', 'Mail.Send', 'Mail.Send.Shared', 'Calendars.ReadWrite', 'Notes.ReadWrite')
Connect-MgGraph -Scopes $scopes -NoWelcome
function G($method, $uri, $body, $contentType = 'application/json') {
    $p = @{ Method = $method; Uri = "https://graph.microsoft.com/v1.0$uri" }
    if ($null -ne $body) { $p.Body = $body; $p.ContentType = $contentType }
    Invoke-MgGraphRequest @p
}
$me = G GET '/me?$select=displayName,userPrincipalName,mail'
$myMail = if ($me.mail) { $me.mail } else { $me.userPrincipalName }
$domain = ($me.userPrincipalName -split '@')[1]
Say ("{0} <{1}>" -f $me.displayName, $myMail) Green
if ($WhatIf) { Say (T 'WhatIf: nothing is changed, only shown.' 'WhatIf: Es wird nichts geaendert, nur angezeigt.') Yellow }

$report = [ordered]@{}
function Count($area, $what) { if (-not $report[$area]) { $report[$area] = @{ new = 0; updated = 0; same = 0; failed = 0; pending = 0 } }; $report[$area][$what]++ }

# ---------- 1) senders (Exchange Online, separate process: Graph and EXO modules clash in one session) ----------
$senders = @{}
$others = @($people | Where-Object { $_ -ne $me.displayName })
if ($others.Count) {
    Write-Host ""
    Say (T 'Senders and attendees (Exchange Online) ...' 'Absender und Teilnehmende (Exchange Online) ...') Cyan
    $tmpIn = [IO.Path]::GetTempFileName(); $tmpOut = [IO.Path]::GetTempFileName(); $tmpPs = [IO.Path]::GetTempFileName() + '.ps1'
    [IO.File]::WriteAllText($tmpIn, (To-AsciiJson @{ upn = $me.userPrincipalName; domain = $domain; names = $others; whatIf = [bool]$WhatIf }))
    [IO.File]::WriteAllText($tmpPs, @'
param($In, $Out)
$ErrorActionPreference = 'Stop'
$cfg = Get-Content $In -Raw | ConvertFrom-Json
if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
    Write-Host '    Installing module ExchangeOnlineManagement (current user) ...'
    if ($PSVersionTable.PSEdition -ne 'Core') { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null }
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
}
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName $cfg.upn -ShowBanner:$false
$res = @()
foreach ($n in $cfg.names) {
    $o = [ordered]@{ name = $n; address = $null; created = $false; granted = $false; error = $null }
    try {
        $r = Get-Recipient -Filter "DisplayName -eq '$($n -replace "'", "''")'" -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize 1 -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $r) {
            $alias = ($n.ToLowerInvariant() -replace '[^a-z0-9]+', '.').Trim('.')
            if ($cfg.whatIf) { $o.address = "$alias@$($cfg.domain)"; $o.created = $true; $res += $o; continue }
            $r = New-Mailbox -Shared -Name $n -DisplayName $n -Alias $alias -PrimarySmtpAddress "$alias@$($cfg.domain)"
            $o.created = $true
        }
        $o.address = "$($r.PrimarySmtpAddress)"
        $has = Get-RecipientPermission -Identity $r.Identity -Trustee $cfg.upn -AccessRights SendAs -ErrorAction SilentlyContinue
        if (-not $has -and -not $cfg.whatIf) { Add-RecipientPermission -Identity $r.Identity -Trustee $cfg.upn -AccessRights SendAs -Confirm:$false | Out-Null; $o.granted = $true }
        elseif (-not $has) { $o.granted = $true }
    } catch { $o.error = "$_" }
    $res += $o
}
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
[IO.File]::WriteAllText($Out, (ConvertTo-Json @($res) -Depth 4))
'@)
    $exe = (Get-Process -Id $PID).Path
    & $exe -NoProfile -ExecutionPolicy Bypass -File $tmpPs -In $tmpIn -Out $tmpOut | Out-Host
    $result = @()
    if ((Get-Item $tmpOut).Length -gt 0) { $result = @(Get-Content $tmpOut -Raw | ConvertFrom-Json) }
    Remove-Item $tmpIn, $tmpOut, $tmpPs -ErrorAction SilentlyContinue
    if (-not $result.Count) { Say (T 'Exchange Online step failed - mails are skipped, invitations go without these attendees.' 'Exchange-Online-Schritt fehlgeschlagen - Mails werden uebersprungen, Termine ohne diese Teilnehmenden.') Red }
    foreach ($r in $result) {
        if ($r.error) { Row (T 'error' 'Fehler') "$($r.name): $($r.error)" Red; Count 'Senders' failed; continue }
        $senders[$r.name] = $r
        $note = @()
        if ($r.created) { $note += (T 'shared mailbox created' 'freigegebenes Postfach angelegt') }
        if ($r.granted) { $note += (T '"Send As" granted' '"Senden als" erteilt') }
        if ($note.Count) { Row (T 'new' 'neu') ("{0} <{1}> ({2})" -f $r.name, $r.address, ($note -join ', ')) Green; Count 'Senders' new }
        else { Row (T 'unchanged' 'unveraendert') ("{0} <{1}>" -f $r.name, $r.address) DarkGray; Count 'Senders' same }
    }
}

# ---------- 2) files -> OneDrive ----------
Write-Host ""
Say (T 'Files -> OneDrive' 'Dateien -> OneDrive') Cyan
function Seg($p) { (($p -split '/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/' }
$firstFolder = $null
foreach ($it in $Items) {
    if (-not (Test-Path $it.dir)) { continue }
    $files = Get-ChildItem $it.dir -Recurse -File | Where-Object {
        $_.Name -notin 'manifest.json', 'shared.json', 'tenant.json' -and $_.FullName.Substring($it.dir.Length + 1) -notmatch '^tenant\\' }
    foreach ($f in $files) {
        $rel = $f.FullName.Substring($it.dir.Length).TrimStart('\') -replace '\\', '/'
        $path = Seg "$($it.drivePath)/$rel"
        $remote = $null
        try { $remote = G GET "/me/drive/root:/$($path)?`$select=size,file" } catch { $remote = $null }
        $local = Get-QuickXor $f.FullName
        if ($remote -and $remote.file -and $remote.file.hashes -and $remote.file.hashes.quickXorHash -eq $local) { Count 'OneDrive' same; continue }
        $tag = if ($remote) { 'updated' } else { 'new' }
        if (-not $WhatIf) {
            try { Invoke-MgGraphRequest -Method PUT -Uri "https://graph.microsoft.com/v1.0/me/drive/root:/$($path):/content" -InputFilePath $f.FullName -ContentType 'application/octet-stream' | Out-Null }
            catch { Row (T 'error' 'Fehler') "$($it.drivePath)/$rel : $_" Red; Count 'OneDrive' failed; continue }
        }
        Count 'OneDrive' $tag
        $label = if ($tag -eq 'new') { T 'new' 'neu' } else { T 'updated' 'aktualisiert' }
        Row $label "$($it.drivePath)/$rel" $(if ($tag -eq 'new') { 'Green' } else { 'Yellow' })
    }
    if (-not $firstFolder) { $firstFolder = ($it.drivePath -split '/')[0] }
}
if ($report['OneDrive'] -and $report['OneDrive'].new -eq 0 -and $report['OneDrive'].updated -eq 0) { Say (T 'all files up to date' 'alle Dateien aktuell') DarkGray }

# ---------- 3) OneNote ----------
foreach ($c in $configs | Where-Object { $_.Config.onenote }) {
    $on = $c.Config.onenote
    Write-Host ""
    Say ("OneNote: {0} > {1}" -f $on.notebook, $on.section) Cyan
    try {
        $nb = @((G GET '/me/onenote/notebooks?$select=id,displayName').value) | Where-Object { $_.displayName -eq $on.notebook } | Select-Object -First 1
        if (-not $nb -and -not $WhatIf) { $nb = G POST '/me/onenote/notebooks' (To-AsciiJson @{ displayName = $on.notebook }); Row (T 'new' 'neu') (T "notebook $($on.notebook)" "Notizbuch $($on.notebook)") Green }
        $sec = $null
        if ($nb) {
            $sec = @((G GET "/me/onenote/notebooks/$($nb.id)/sections?`$select=id,displayName").value) | Where-Object { $_.displayName -eq $on.section } | Select-Object -First 1
            if (-not $sec -and -not $WhatIf) { $sec = G POST "/me/onenote/notebooks/$($nb.id)/sections" (To-AsciiJson @{ displayName = $on.section }); Row (T 'new' 'neu') (T "section $($on.section)" "Abschnitt $($on.section)") Green }
        }
        $titles = @()
        if ($sec) { $titles = @((G GET "/me/onenote/sections/$($sec.id)/pages?`$select=title&`$top=100").value | ForEach-Object { $_.title }) }
        foreach ($p in $on.pages) {
            $title = L $p.title
            if ($titles -contains $title) { Row (T 'unchanged' 'unveraendert') $title DarkGray; Count 'OneNote' same; continue }
            if (-not $WhatIf) {
                $html = [Text.Encoding]::UTF8.GetString((Get-Resource $c.Item $p.file))
                $html = [regex]::Replace($html, '<title>.*?</title>', "<title>$([Net.WebUtility]::HtmlEncode($title))</title>")
                G POST "/me/onenote/sections/$($sec.id)/pages" (To-AsciiHtml $html) 'text/html' | Out-Null
            }
            Row (T 'new' 'neu') $title Green; Count 'OneNote' new
        }
    } catch { Row (T 'error' 'Fehler') "OneNote: $_" Red; Count 'OneNote' failed }
}

# ---------- 4) meetings ----------
$tz = [TimeZoneInfo]::Local.Id
foreach ($c in $configs | Where-Object { $_.Config.events }) {
    Write-Host ""
    Say (T 'Calendar' 'Kalender') Cyan
    foreach ($e in $c.Config.events) {
        $subject = L $e.subject
        try {
            $from = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'); $to = (Get-Date).AddDays(8).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
            $existing = @((G GET "/me/calendarView?startDateTime=$from&endDateTime=$to&`$select=subject&`$top=200").value) | Where-Object { $_.subject -eq $subject }
            if ($existing) { Row (T 'unchanged' 'unveraendert') $subject DarkGray; Count 'Calendar' same; continue }
            $day = Next-Weekday $e.weekday
            $start = [datetime]::ParseExact("$($day.ToString('yyyy-MM-dd')) $($e.start)", 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
            $end = $start.AddMinutes([int]$e.minutes)
            $att = @(foreach ($n in $e.attendees) { if ($senders[$n]) { @{ type = 'required'; emailAddress = @{ address = $senders[$n].address; name = $n } } } })
            $ev = @{ subject = $subject; body = @{ contentType = 'text'; content = (L $e.body) }
                     start = @{ dateTime = $start.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = $tz }
                     end = @{ dateTime = $end.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = $tz } }
            if ($att.Count) { $ev.attendees = $att }
            if (-not $WhatIf) { G POST '/me/events' (To-AsciiJson $ev) | Out-Null }
            Row (T 'new' 'neu') ("{0} ({1:ddd dd.MM. HH:mm})" -f $subject, $start) Green; Count 'Calendar' new
        } catch { Row (T 'error' 'Fehler') "$subject : $_" Red; Count 'Calendar' failed }
    }
}

# ---------- 5) mails (last, so "Send As" has time to replicate) ----------
if ($mails.Count) {
    Write-Host ""
    Say ((T 'Mails -> Inbox ({0})' 'Mails -> Posteingang ({0})') -f $Language) Cyan
    $deadline = (Get-Date).AddMinutes($SendAsWaitMinutes)
    foreach ($m in $mails) {
        $s = $senders[$m.FromName]
        if (-not $s) { Row (T 'skipped' 'uebersprungen') ("{0} ({1})" -f $m.Subject, (T "no sender $($m.FromName)" "kein Absender $($m.FromName)")) Red; Count 'Mails' failed; continue }
        try {
            $q = [uri]::EscapeDataString("subject eq '$($m.Subject -replace "'", "''")'")
            $found = @((G GET "/me/mailFolders/inbox/messages?`$filter=$q&`$select=subject,from&`$top=25").value) |
                Where-Object { $_.from.emailAddress.address -eq $s.address }
            if ($found) { Row (T 'unchanged' 'unveraendert') ("{0}: {1}" -f $m.FromName, $m.Subject) DarkGray; Count 'Mails' same; continue }
        } catch { }
        if ($WhatIf) { Row (T 'new' 'neu') ("{0}: {1}" -f $m.FromName, $m.Subject) Green; Count 'Mails' new; continue }
        $msg = To-AsciiJson @{ message = @{ subject = $m.Subject; body = @{ contentType = 'Text'; content = $m.Body }
                                            from = @{ emailAddress = @{ address = $s.address; name = $m.FromName } }
                                            toRecipients = @(@{ emailAddress = @{ address = $myMail; name = $me.displayName } }) }
                               saveToSentItems = $false }
        $sent = $false; $waitedMsg = $false
        while (-not $sent) {
            try { G POST '/me/sendMail' $msg | Out-Null; $sent = $true }
            catch {
                $err = "$_"
                if ($err -match 'SendAs|ErrorSendAsDenied|on behalf|403|Forbidden' -and (Get-Date) -lt $deadline) {
                    if (-not $waitedMsg) { Say ((T '    Waiting for "Send As" on {0} (Exchange needs a few minutes) ...' '    Warte auf "Senden als" fuer {0} (Exchange braucht ein paar Minuten) ...') -f $m.FromName) DarkYellow; $waitedMsg = $true }
                    Start-Sleep -Seconds 60
                } else { break }
            }
        }
        if ($sent) { Row (T 'new' 'neu') ("{0}: {1}" -f $m.FromName, $m.Subject) Green; Count 'Mails' new }
        else { Row (T 'later' 'spaeter') ("{0}: {1} - {2}" -f $m.FromName, $m.Subject, (T 'run again later' 'spaeter erneut ausfuehren')) Yellow; Count 'Mails' pending }
    }
}

# ---------- summary + manual checklist ----------
Write-Host ""
$report.GetEnumerator() | ForEach-Object {
    [pscustomobject]@{ Area = $_.Key; New = $_.Value.new; Updated = $_.Value.updated; Unchanged = $_.Value.same; Later = $_.Value.pending; Failed = $_.Value.failed }
} | Format-Table -AutoSize | Out-String | Write-Host

$manual = @(foreach ($c in $configs) { foreach ($m in @(L $c.Config.manual)) { if ($m) { [pscustomobject]@{ Demo = $c.Item.id; Text = ($m -replace '\*\*', '') } } } })
if ($manual.Count) {
    Say (T 'Still to do by hand (no API for this):' 'Noch von Hand (dafuer gibt es keine API):') Cyan
    $n = 0
    foreach ($m in $manual) { $n++; Write-Host ("  {0,2}. [{1}] {2}" -f $n, $m.Demo, $m.Text) }
    Write-Host ""
}
if ($firstFolder -and -not $WhatIf) {
    try { $web = (G GET "/me/drive/root:/$(Seg $firstFolder)?`$select=webUrl").webUrl; Say ("OneDrive: $web") Green } catch { }
}
Disconnect-MgGraph | Out-Null
