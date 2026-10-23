<#
.SYNOPSIS
Read-only account/connectivity preflight; not a substitute for an in-app Copilot rehearsal.
.DESCRIPTION
Use -Connect only when ready for a browser/window sign-in. No device code, license assignment,
role assignment, mailbox creation, Send As grant, billing change or tenant content write.
The optional private JSON report can contain account information: never commit it.
Sources:
https://learn.microsoft.com/graph/api/user-get
https://learn.microsoft.com/graph/api/user-list-licensedetails
https://learn.microsoft.com/graph/api/calendar-get
#>
param(
    [Parameter(Mandatory)] [string]$ExpectedAccount,
    [switch]$Connect,
    [string]$Report
)
$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
if ($ExpectedAccount -notmatch '^[^@\s]+@[^@\s]+$') { throw 'ExpectedAccount must be the exact presenter UPN.' }
if ($Connect) {
    # Same union as the global installer: later tenant setup reuses this sign-in.
    Connect-MgGraph -ContextScope CurrentUser -NoWelcome -Scopes @(
        'User.Read', 'User.Read.All', 'Files.ReadWrite', 'Files.ReadWrite.All',
        'Mail.ReadBasic', 'Mail.Send', 'Mail.Send.Shared', 'Calendars.ReadWrite', 'Notes.ReadWrite',
        'Team.ReadBasic.All', 'Channel.ReadBasic.All', 'ChannelMessage.Read.All',
        'ChannelMessage.Send', 'Mail.ReadWrite'
    )
}
$checks = New-Object Collections.Generic.List[object]
$context = Get-MgContext
function Add-Check($Name, $Status, $Detail) {
    $checks.Add([pscustomobject]@{ name = $Name; status = $Status; detail = $Detail })
}
function Read-Graph($Path) {
    Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0$Path" -OutputType PSObject
}
if (-not $context) {
    Add-Check 'Graph sign-in' 'blocked' 'No cached Graph context. Run with -Connect when available for a browser/window sign-in.'
} elseif ($context.Account -ne $ExpectedAccount -or -not $context.TenantId) {
    Add-Check 'Presenter identity' 'blocked' 'The cached Graph account is not the expected presenter; no account resources were read.'
} else {
    try {
        $me = Read-Graph '/me?$select=displayName,userPrincipalName,mail,assignedPlans'
        if ($me.userPrincipalName -ne $ExpectedAccount -or $me.mail -ne $ExpectedAccount) {
            throw 'Actual Graph user and primary mailbox must match ExpectedAccount.'
        }
        Add-Check 'Presenter identity' 'ok' $me.userPrincipalName
        $exchange = @($me.assignedPlans | Where-Object { $_.service -eq 'exchange' -and $_.capabilityStatus -eq 'Enabled' })
        if (-not $exchange.Count) { throw 'No enabled Exchange Online plan; a Teams-only account is insufficient.' }
        Add-Check 'Exchange Online plan' 'ok' 'An enabled Exchange service plan was returned for the actual presenter.'
    } catch {
        Add-Check 'Presenter mailbox' 'blocked' "$_"
        $me = $null
    }
    if ($me) {
        foreach ($check in @(
            @{ name = 'OneDrive'; path = '/me/drive/root?$select=id,webUrl'; field = 'id' },
            @{ name = 'Primary inbox'; path = '/me/mailFolders/inbox?$select=id'; field = 'id' },
            @{ name = 'Calendar'; path = '/me/calendar?$select=id,allowedOnlineMeetingProviders'; field = 'id' }
        )) {
            try {
                $value = Read-Graph $check.path
                if (-not $value.($check.field)) { throw 'Expected resource identity was not returned.' }
                Add-Check $check.name 'ok' 'Read access confirmed; no data was written.'
                if ($check.name -eq 'Calendar' -and 'teamsForBusiness' -notin @($value.allowedOnlineMeetingProviders)) {
                    Add-Check 'Teams meeting provider' 'blocked' 'The calendar does not advertise teamsForBusiness.'
                }
            } catch { Add-Check $check.name 'blocked' "$_" }
        }
        try {
            $licenses = Read-Graph '/me/licenseDetails'
            if (-not @($licenses.value).Count) { throw 'No license details were returned.' }
            Add-Check 'License evidence' 'review' @($licenses.value | ForEach-Object {
                [pscustomobject]@{ sku = $_.skuPartNumber; plans = @($_.servicePlans | Select-Object servicePlanName,provisioningStatus) }
            })
        } catch {
            Add-Check 'License evidence' 'blocked' "License details could not be verified: $_. Do not grant an admin role to the presenter just to pass this check; have the tenant administrator verify the assignment."
        }
    }
}
$result = [ordered]@{
    checkedAt = [datetime]::UtcNow.ToString('o')
    expectedAccount = $ExpectedAccount
    checks = @($checks.ToArray())
    liveDemoReady = $false
    remainingRehearsal = @(
        'Confirm the actual in-app Microsoft 365 Copilot (Premium) label for the presenter.',
        'Run the global -Tenant -WhatIf setup as this presenter; Exchange admin authenticates separately if Send As needs preparing.',
        'Verify Cowork access and existing approved credits/spending policy; do not enable billing.',
        'Rehearse every chosen prompt and inspect its actual output.',
        'Verify saved Teams transcript, attached video transcript, OneNote desktop opening, and approved SharePoint/Planner destinations.'
    )
}
if ($Report) {
    $parent = Split-Path -Parent ([IO.Path]::GetFullPath($Report))
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw 'Report parent folder must already exist.' }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($Report), ($result | ConvertTo-Json -Depth 12), (New-Object Text.UTF8Encoding $false))
}
$checks | Select-Object name,status | Format-Table -AutoSize | Out-Host
Write-Host 'Account checks are not a completed Copilot rehearsal. See remainingRehearsal in the private report.'
if (@($checks | Where-Object status -eq 'blocked').Count) { exit 2 }
