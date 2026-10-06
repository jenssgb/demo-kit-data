<#
.SYNOPSIS
  Writes bundles/<bundle>.json from a Demo Kit deck: every demoData.demo in the deck, in deck order.
  Run whenever a demo is added to the deck, then commit and push – the bundle one-liner stays the same:
    .\tools\Update-Bundle.ps1 -Bundle bpw -Deck D:\code\demo-kit-new2026\decks\bpw-ai-multiplikatoren.json -Folder Demo-BPW
#>
param(
    [Parameter(Mandatory = $true)][string]$Bundle,
    [Parameter(Mandatory = $true)][string]$Deck,
    [string]$Folder
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$utf8 = New-Object Text.UTF8Encoding $false
$deckJson = [IO.File]::ReadAllText((Resolve-Path $Deck), $utf8)

# Every "demoData": { ... "demo": "<id>" ... } in document order
$ids = [regex]::Matches($deckJson, '"demoData"\s*:\s*\{[^{}]*?"demo"\s*:\s*"([^"]+)"') |
    ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique
if (-not $ids) { throw "No demoData.demo entries found in $Deck." }

foreach ($id in $ids) {
    if (-not (Test-Path (Join-Path $root "$id\manifest.json"))) { throw "Demo '$id' has no manifest.json in this repo." }
}

$deckObj = $deckJson | ConvertFrom-Json
$dir = Join-Path $root 'bundles'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$path = Join-Path $dir "$Bundle.json"
$old = if (Test-Path $path) { [IO.File]::ReadAllText($path, $utf8) | ConvertFrom-Json } else { $null }

if (-not $Folder) { $Folder = if ($old) { $old.folder } else { "Demo-$Bundle" } }
$title = [ordered]@{ en = $deckObj.meta.title.en; de = $deckObj.meta.title.de }
$deckId = [IO.Path]::GetFileNameWithoutExtension($Deck)

$out = [ordered]@{ title = $title; folder = $Folder; deck = $deckId; demos = @($ids) }

# Tenant casts (deck.personas + deck.profiles) -> profiles/<id>.json for tools/build_profile.py;
# deck.files.demoData.profile becomes the profile the bundle installs by default.
function Text($v) { if ($v -is [string]) { $v } elseif ($v.en) { $v.en } else { "$v" } }
if ($deckObj.personas -and $deckObj.profiles) {
    $pdir = Join-Path $root 'profiles'
    New-Item -ItemType Directory -Force -Path $pdir | Out-Null
    foreach ($p in $deckObj.profiles.PSObject.Properties) {
        $prof = [ordered]@{
            label = [ordered]@{ en = (Text $p.Value.label); de = $(if ($p.Value.label.de) { $p.Value.label.de } else { Text $p.Value.label }) }
            tenant = $p.Value.tenant
            personas = $deckObj.personas
            people = $p.Value.people
        }
        [IO.File]::WriteAllText((Join-Path $pdir "$($p.Name).json"), ($prof | ConvertTo-Json -Depth 6) + "`n", $utf8)
        Write-Host "profiles/$($p.Name).json -> run: python tools\build_profile.py $($p.Name); .\tools\Update-Manifest.ps1 -All -Root profiles\$($p.Name)"
    }
}
$bundleProfile = $deckObj.files.demoData.profile
if ($bundleProfile) {
    if (-not $deckObj.profiles.$bundleProfile) { throw "files.demoData.profile '$bundleProfile' is not in deck.profiles." }
    $out.profile = $bundleProfile
}

[IO.File]::WriteAllText($path, ($out | ConvertTo-Json -Depth 5) + "`n", $utf8)
Write-Host "bundles/$Bundle.json: $(@($ids).Count) demos -> $Folder"
$ids | ForEach-Object { Write-Host "  $_" }
