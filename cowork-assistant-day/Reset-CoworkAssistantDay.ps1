<#
.SYNOPSIS
  Creates a fresh EAD26 workbook/context/output folder; never deletes old results.
  Uses the existing cached authorized Graph session. Review -WhatIf first.
#>
[CmdletBinding(SupportsShouldProcess)]
param([ValidateSet('en', 'de')] [string]$Language = 'de')

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
. (Join-Path $PSScriptRoot 'DemoCalendar.ps1')
$account = 'admin@M365CPI98544940.OnMicrosoft.com'
$context = Get-MgContext
if (-not $context) { throw 'No cached Graph session. Complete the normal authorized tenant setup first; this reset never starts sign-in.' }
$me = Invoke-DemoGraph GET '/me?$select=userPrincipalName,mail'
Assert-DemoPresenter $account $context $me
$base = 'Demo-Kit/Cowork/cowork-assistant-day'
$source = Invoke-DemoGraph GET "/me/drive/root:/$base/baseline/$Language/Action_Register.xlsx"
$template = Invoke-DemoGraph GET "/me/drive/root:/$base/baseline/$Language/Run_Context.md"
$parent = Invoke-DemoGraph GET "/me/drive/root:/$base/runs/$Language"
if (-not $source.file -or -not $template.file -or -not $parent.folder) {
    throw 'EAD26 baseline or run context/folder missing; complete normal scenario installation first.'
}
$hashFile = Join-Path $PSScriptRoot '..\cowork-assistant-day\baseline\hashes.json'
if (-not (Test-Path -LiteralPath $hashFile)) { $hashFile = Join-Path $PSScriptRoot 'baseline\hashes.json' }
if (-not (Test-Path -LiteralPath $hashFile)) { throw 'Baseline hashes missing; run reset beside the installed baseline folder or from demo-kit-data tools.' }
$hashes = Get-Content -LiteralPath $hashFile -Raw -Encoding UTF8 | ConvertFrom-Json
$expected = $hashes.$Language
if ($expected -notmatch '^[0-9a-f]{64}$') { throw 'Invalid registered baseline hash.' }
$download = $source.'@microsoft.graph.downloadUrl'
$contextUrl = $template.'@microsoft.graph.downloadUrl'
if ($download -notmatch '^https://' -or $contextUrl -notmatch '^https://') { throw 'Baseline/context download URL missing.' }
$temp = [IO.Path]::GetTempFileName()
$contextTemp = [IO.Path]::GetTempFileName()
try {
    Invoke-WebRequest -Uri $download -OutFile $temp -UseBasicParsing
    if ((Get-FileHash -LiteralPath $temp -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expected) {
        throw 'Baseline changed; refusing to clone edited rehearsal output.'
    }
    Invoke-WebRequest -Uri $contextUrl -OutFile $contextTemp -UseBasicParsing
    if ($hashes.contexts.$Language -notmatch '^[0-9a-f]{64}$' -or
        (Get-FileHash -LiteralPath $contextTemp -Algorithm SHA256).Hash.ToLowerInvariant() -cne $hashes.contexts.$Language) {
        throw 'Baseline context changed; refusing unverified run instructions.'
    }
    $text = [IO.File]::ReadAllText($contextTemp, [Text.Encoding]::UTF8)
    if (-not $text.Contains('EAD26') -or -not $text.Contains("runs/$Language/01")) {
        throw 'Original run context is not the expected EAD26 template.'
    }
    $run = 'fresh-{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8))
    $text = $text.Replace("runs/$Language/01", "runs/$Language/$run").Replace('run 01', "run $run").Replace('Lauf 01', "Lauf $run")
    [IO.File]::WriteAllText($contextTemp, $text, (New-Object Text.UTF8Encoding $false))
    if (-not $PSCmdlet.ShouldProcess("$base/runs/$Language/$run", 'Create isolated fresh input and outputs; preserve all earlier results')) { return }
    $folder = Invoke-DemoGraph POST "/me/drive/items/$([uri]::EscapeDataString($parent.id))/children" @{
        name = $run; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail'
    }
    $id = [uri]::EscapeDataString($folder.id)
    $input = Invoke-DemoGraph PUT "/me/drive/items/${id}:/Action_Register.xlsx:/content" -InputFile $temp
    $runContext = Invoke-DemoGraph PUT "/me/drive/items/${id}:/Run_Context.md:/content" -InputFile $contextTemp
    $output = Invoke-DemoGraph POST "/me/drive/items/$id/children" @{
        name = 'outputs'; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail'
    }
    [pscustomobject]@{
        scenario = 'EAD26'; run = $run; context = $runContext.webUrl; input = $input.webUrl; outputs = $output.webUrl
        instructions = 'Select this Run_Context.md and one Case.md in a new conversation. No send/post/invite/calendar change/automation. Previous runs remain untouched.'
    } | Format-List | Out-Host
} finally {
    Remove-Item -LiteralPath $temp, $contextTemp -Force -WhatIf:$false
}
