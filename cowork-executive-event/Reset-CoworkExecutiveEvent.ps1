<#
.SYNOPSIS
  Creates an isolated fresh LEF26 action workbook/output namespace. No deletion.
  Use -WhatIf first. Requires the existing authorized cached Graph session.
#>
[CmdletBinding(SupportsShouldProcess)]
param([ValidateSet('en', 'de')] [string]$Language = 'de')

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
. (Join-Path $PSScriptRoot 'ExecutiveEvent.ps1')
$config = [pscustomobject]@{
    account = 'admin@M365CPI98544940.OnMicrosoft.com'
    domain = 'M365CPI98544940.OnMicrosoft.com'
}
$settings = Get-ExecutiveEventSettings $config
$context = Get-MgContext
if (-not $context) { throw 'No cached demo Graph session. Sign in through the normal tenant setup; no device-code or automatic sign-in is used here.' }
$me = Invoke-ExecutiveEventGraph GET '/me?$select=userPrincipalName,mail'
Assert-ExecutiveEventIdentity $config $settings $context $me
$base = "Demo-Kit/Cowork/cowork-executive-event/runs/$Language"
$source = Invoke-ExecutiveEventGraph GET "/me/drive/root:/Demo-Kit/Cowork/cowork-executive-event/baseline/$Language/Action_Register.xlsx"
$parent = Invoke-ExecutiveEventGraph GET "/me/drive/root:/$base"
$run = 'fresh-{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8))
if (-not $source.file -or -not $parent.folder) { throw 'LEF26 baseline or run folder missing. Complete the authorized scenario setup first.' }
$download = $source.'@microsoft.graph.downloadUrl'
if ($download -notmatch '^https://') { throw 'LEF26 baseline download link missing.' }
$hashesPath = Join-Path $PSScriptRoot '..\cowork-executive-event\baseline\hashes.json'
if (-not (Test-Path -LiteralPath $hashesPath)) { $hashesPath = Join-Path $PSScriptRoot 'baseline\hashes.json' }
if (-not (Test-Path -LiteralPath $hashesPath)) { throw 'Run reset from the installed scenario folder or demo-kit-data tools, with baseline\hashes.json present.' }
$hashes = Get-Content -LiteralPath $hashesPath -Raw -Encoding UTF8 | ConvertFrom-Json
$expectedHash = $hashes.$Language
if ($expectedHash -notmatch '^[0-9a-f]{64}$') { throw 'Baseline SHA-256 is missing or invalid in baseline\hashes.json.' }
$temp = [IO.Path]::GetTempFileName()
try {
    Invoke-WebRequest -Uri $download -OutFile $temp -UseBasicParsing
    if ((Get-FileHash -LiteralPath $temp -Algorithm SHA256).Hash.ToLowerInvariant() -ne $expectedHash) {
        throw 'LEF26 baseline workbook differs from its registered hash. Do not clone edited rehearsal results as fresh input.'
    }
    if (-not $PSCmdlet.ShouldProcess("$base/$run", 'Create a fresh action workbook and empty outputs folder; keep all previous results')) { return }
    $folder = Invoke-ExecutiveEventGraph POST "/me/drive/items/$([uri]::EscapeDataString($parent.id))/children" @{
        name = $run; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail'
    }
    $id = [uri]::EscapeDataString($folder.id)
    $input = Invoke-ExecutiveEventGraph PUT "/me/drive/items/${id}:/Action_Register.xlsx:/content" -InputFile $temp
    $output = Invoke-ExecutiveEventGraph POST "/me/drive/items/$id/children" @{
        name = 'outputs'; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail'
    }
    $report = [ordered]@{
        scenario = 'LEF26'; account = $settings.account; tenantId = $settings.tenantId
        run = $run; language = $Language; input = $input.webUrl; outputs = $output.webUrl
        instructions = 'New conversation; BEFORE sources first. Put this run name in saved draft subjects. No send/post/invite/booking. Earlier results are untouched.'
    }
    $path = Join-Path $env:LOCALAPPDATA "DemoKit\executive-event\$($settings.tenantId)\active-$Language.json"
    New-Item -ItemType Directory -Force -Path (Split-Path $path -Parent) | Out-Null
    [IO.File]::WriteAllText($path, (ConvertTo-Json $report), (New-Object Text.UTF8Encoding $false))
    $report | Format-List | Out-Host
    Write-Host "Private active-run report: $path"
} finally {
    Remove-Item -LiteralPath $temp -Force -WhatIf:$false
}
