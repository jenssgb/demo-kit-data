<#
.SYNOPSIS
  Rebuilds <demo>/manifest.json from the files in the demo folder (keeps title and next steps).
  Run after adding, changing or removing demo files (also after the build-*.py scripts):
    .\tools\Update-Manifest.ps1 -Demo onenote-copilot
    .\tools\Update-Manifest.ps1 -All

  files  : every file in the demo folder as { path, sha256 } (installer skips unchanged files).
  Not listed (only used by install.ps1 -Tenant): <demo>/tenant.json and the <demo>/tenant/ folder.
  shared : optional <demo>/shared.json – a list of "<other-demo>/<path>" entries. Those files are
           copied into this demo folder too, so every demo folder is self-contained.
#>
param([string]$Demo, [switch]$All)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$utf8 = New-Object Text.UTF8Encoding $false

function Get-Sha([string]$file) { (Get-FileHash -Algorithm SHA256 -LiteralPath $file).Hash.ToLowerInvariant() }

function Update-One([string]$id) {
    $dir = Join-Path $root $id
    if (-not (Test-Path $dir)) { throw "Folder $dir not found." }

    $path = Join-Path $dir 'manifest.json'
    $manifest = if (Test-Path $path) { [IO.File]::ReadAllText($path, $utf8) | ConvertFrom-Json } else {
        [pscustomobject]@{ title = [pscustomobject]@{ en = $id; de = $id }; next = [pscustomobject]@{ en = @(); de = @() } }
    }

    $files = Get-ChildItem $dir -Recurse -File |
        Where-Object { $_.Name -notin @('manifest.json', 'shared.json', 'tenant.json') -and
                       $_.FullName.Substring($dir.Length + 1) -notmatch '^tenant\\' } |
        Sort-Object { $_.FullName.Substring($dir.Length + 1) -replace '\\', '/' } |
        ForEach-Object { [ordered]@{ path = ($_.FullName.Substring($dir.Length + 1) -replace '\\', '/'); sha256 = (Get-Sha $_.FullName) } }

    $out = [ordered]@{ title = $manifest.title; next = $manifest.next; files = @($files) }

    $sharedPath = Join-Path $dir 'shared.json'
    if (Test-Path $sharedPath) {
        $shared = foreach ($entry in @([IO.File]::ReadAllText($sharedPath, $utf8) | ConvertFrom-Json)) {
            $from, $rel = $entry -split '/', 2
            $src = Join-Path $root (($entry) -replace '/', '\')
            if (-not (Test-Path $src)) { throw "$id/shared.json: $entry not found." }
            [ordered]@{ from = $from; path = $rel; sha256 = (Get-Sha $src) }
        }
        $out.shared = @($shared)
    }

    [IO.File]::WriteAllText($path, ($out | ConvertTo-Json -Depth 6) + "`n", $utf8)
    $n = if ($out.shared) { " + $(@($out.shared).Count) shared" } else { '' }
    Write-Host ("{0,-24} {1} files{2}" -f $id, @($files).Count, $n)
}

if ($All) {
    Get-ChildItem $root -Directory | Where-Object { Test-Path (Join-Path $_.FullName 'manifest.json') } |
        ForEach-Object { Update-One $_.Name }
} elseif ($Demo) { Update-One $Demo }
else { throw 'Use -Demo <id> or -All.' }

# people.json: every sender and attendee of every demo's tenant.json. tenant.ps1 prepares all of them on the one
# Exchange Online sign-in per tenant, so new demos need no second sign-in.
$people = foreach ($tj in Get-ChildItem $root -Directory | ForEach-Object { Join-Path $_.FullName 'tenant.json' } | Where-Object { Test-Path $_ }) {
    $cfg = [IO.File]::ReadAllText($tj, $utf8) | ConvertFrom-Json
    $dir = Split-Path $tj -Parent
    foreach ($m in @($cfg.mails)) {
        if (-not $m) { continue }
        foreach ($lang in 'en', 'de') {
            $p = Join-Path $dir (($m.Replace('{lang}', $lang)) -replace '/', '\')
            if (-not (Test-Path $p)) { continue }
            $from = (Get-Content -LiteralPath $p -TotalCount 40 | Where-Object { $_ -match '^From:' } | Select-Object -First 1)
            if ($from -match '^From:\s*"?([^"<]+?)"?\s*<') { $Matches[1].Trim() }
        }
    }
    foreach ($e in @($cfg.events)) { if ($e) { @($e.attendees) } }
}
$people = @($people | Where-Object { $_ -and $_ -notmatch '=\?' } | Sort-Object -Unique)
[IO.File]::WriteAllText((Join-Path $root 'people.json'), (ConvertTo-Json @($people)) + "`n", $utf8)
Write-Host ("people.json              {0} people" -f $people.Count)

# catalog.json (app > demos for install.ps1) from the kit decks, when the kit repo sits next to this one
if ($All -and (Test-Path (Join-Path (Split-Path $root -Parent) 'demo-kit-new2026\decks'))) { & (Join-Path $PSScriptRoot 'Update-Catalog.ps1') }
