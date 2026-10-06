<#
.SYNOPSIS
  Rebuilds <demo>/manifest.json from the files in the demo folder (keeps title and next steps).
  Run after adding, changing or removing demo files (also after the build-*.py scripts):
    .\tools\Update-Manifest.ps1 -Demo onenote-copilot
    .\tools\Update-Manifest.ps1 -All
    .\tools\Update-Manifest.ps1 -All -Root profiles\cdx      (after tools/build_profile.py cdx)

  files  : every file in the demo folder as { path, sha256 } (installer skips unchanged files).
  Not listed (only used by install.ps1 -Tenant): <demo>/tenant.json and the <demo>/tenant/ folder.
  shared : optional <demo>/shared.json – a list of "<other-demo>/<path>" entries. Those files are
           copied into this demo folder too, so every demo folder is self-contained.
#>
param([string]$Demo, [switch]$All, [string]$Root)

$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$root = if ($Root) { (Resolve-Path (Join-Path $repo $Root)).Path } else { $repo }
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
