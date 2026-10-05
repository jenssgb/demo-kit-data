<#
.SYNOPSIS
  Rebuilds <demo>/manifest.json from the files in the demo folder (keeps title and next steps).
  Run after adding or removing demo files:  .\tools\Update-Manifest.ps1 -Demo onenote-copilot
#>
param([Parameter(Mandatory = $true)][string]$Demo)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$dir = Join-Path $root $Demo
if (-not (Test-Path $dir)) { throw "Folder $dir not found." }

$path = Join-Path $dir 'manifest.json'
$manifest = if (Test-Path $path) { Get-Content $path -Raw -Encoding utf8 | ConvertFrom-Json } else {
    [pscustomobject]@{ title = [pscustomobject]@{ en = $Demo; de = $Demo }; next = [pscustomobject]@{ en = @(); de = @() } }
}

$files = Get-ChildItem $dir -Recurse -File |
    Where-Object { $_.Name -ne 'manifest.json' } |
    ForEach-Object { $_.FullName.Substring($dir.Length + 1) -replace '\\', '/' } |
    Sort-Object

$out = [ordered]@{ title = $manifest.title; next = $manifest.next; files = @($files) }
$json = $out | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText($path, $json + "`n", (New-Object Text.UTF8Encoding $false))
Write-Host "manifest.json: $($files.Count) files"
