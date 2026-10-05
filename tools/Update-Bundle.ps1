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
[IO.File]::WriteAllText($path, ($out | ConvertTo-Json -Depth 5) + "`n", $utf8)
Write-Host "bundles/$Bundle.json: $(@($ids).Count) demos -> $Folder"
$ids | ForEach-Object { Write-Host "  $_" }
