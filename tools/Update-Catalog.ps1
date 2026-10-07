<#
.SYNOPSIS
  Rebuilds catalog.json (app > demos) from the Demo Kit decks. install.ps1 installs exactly what is listed there.

  Source: every deck in the kit (default ..\demo-kit-new2026\decks) that is not hidden and has
  files.demoData.app (= the app folder under OneDrive\Demo-Kit). Its demos are the pages' "data" folders.
  Runs automatically at the end of .\tools\Update-Manifest.ps1 -All. Run it on its own after deck changes:
    .\tools\Update-Catalog.ps1
    .\tools\Update-Catalog.ps1 -Decks D:\code\demo-kit-new2026\decks

  Errors when a deck references a data folder without manifest.json; warns about data folders no deck uses.
#>
param([string]$Decks)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$utf8 = New-Object Text.UTF8Encoding $false
if (-not $Decks) { $Decks = Join-Path (Split-Path $root -Parent) 'demo-kit-new2026\decks' }
if (-not (Test-Path $Decks)) { throw "Deck folder $Decks not found (use -Decks <path>)." }

# Folders of older kit versions (one per customer). install.ps1 -RemoveLegacy deletes them from OneDrive.
$legacy = @('Demo-BPW', 'Demo-DHL')

function Text($o, $lang) { if ($o -is [string]) { $o } elseif ($o.$lang) { $o.$lang } else { $o.en } }

$apps = foreach ($f in Get-ChildItem $Decks -Filter *.json) {
    $deck = [IO.File]::ReadAllText($f.FullName, $utf8) | ConvertFrom-Json
    $app = $deck.files.demoData.app
    if ($deck.hidden -or -not $app) { continue }
    $demos = @($deck.demos | Where-Object { $_.data } | ForEach-Object { $_.data })
    if (-not $demos.Count) { continue }
    [pscustomobject]@{
        order = if ($null -ne $deck.meta.order) { [int]$deck.meta.order } else { 100 }
        entry = [ordered]@{
            id     = $f.BaseName
            folder = $app
            title  = [ordered]@{ en = (Text $deck.meta.title 'en'); de = (Text $deck.meta.title 'de') }
            demos  = $demos
        }
    }
}
$apps = @($apps | Sort-Object order, { $_.entry.title.en } | ForEach-Object { $_.entry })

$used = @{}
$errors = @()
foreach ($a in $apps) {
    foreach ($d in $a.demos) {
        if ($used[$d]) { $errors += "$d is used by both $($used[$d]) and $($a.id)." }
        $used[$d] = $a.id
        if (-not (Test-Path (Join-Path $root "$d\manifest.json"))) { $errors += "$($a.id): data folder '$d' has no manifest.json." }
    }
}
if ($errors.Count) { throw ($errors -join "`n") }
Get-ChildItem $root -Directory | Where-Object { (Test-Path (Join-Path $_.FullName 'manifest.json')) -and -not $used[$_.Name] } |
    ForEach-Object { Write-Warning "$($_.Name) is not used by any deck (add `"data`": `"$($_.Name)`" to a page)." }

$catalog = [ordered]@{
    folder = 'Demo-Kit'
    title  = [ordered]@{ en = 'Demo Kit'; de = 'Demo Kit' }
    legacy = $legacy
    apps   = $apps
}
[IO.File]::WriteAllText((Join-Path $root 'catalog.json'), ($catalog | ConvertTo-Json -Depth 6) + "`n", $utf8)
foreach ($a in $apps) { Write-Host ("catalog  {0,-14} {1}" -f $a.folder, ($a.demos -join ', ')) }
