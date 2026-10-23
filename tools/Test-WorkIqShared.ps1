$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WorkIqShared.ps1')
function Assert($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Must-Fail([scriptblock]$Action, [string]$Pattern = '.') {
    try { & $Action | Out-Null } catch { Assert ("$_" -match $Pattern) "Wrong failure: $_"; return }
    throw 'Expected an explicit failure.'
}
function Invoke-MgGraphRequest {
    param($Method, $Uri, $SkipHttpErrorCheck, $StatusCodeVariable, $OutputType, $Body, $ContentType, $InputFilePath)
    Set-Variable -Name $StatusCodeVariable -Value $script:httpStatus -Scope 1
    @{ value = @(@{ error = @{ code = 'recipientFailure' } }) }
}
$script:httpStatus = 207
Must-Fail { Invoke-WiqGraph POST '/test' @{} } 'HTTP 207'
$script:httpStatus = 403
Must-Fail { Invoke-WiqGraph GET '/test' -AllowNotFound } 'HTTP 403'
$script:httpStatus = 429
Must-Fail { Invoke-WiqGraph GET '/test' } 'HTTP 429'
$script:httpStatus = 404
Assert ($null -eq (Invoke-WiqGraph GET '/test' -AllowNotFound)) 'Only explicit GET 404 may become null.'
Must-Fail { Invoke-WiqGraph GET '//other.invalid' } 'invalid Graph path'

$script:tenant = '11111111-1111-1111-1111-111111111111'
function Get-MgContext { @{ TenantId = $script:tenant } }
function Get-QuickXor([string]$File) { (Get-FileHash -LiteralPath $File -Algorithm SHA256).Hash }
$me = @{ id = 'admin'; mail = 'admin@contoso.invalid'; displayName = 'Demo Admin' }
function New-User($Id, $Type = 'Member', $Enabled = $true, $Exchange = $true) {
    @{ id = $Id; userPrincipalName = "$Id@contoso.invalid"; mail = "$Id@contoso.invalid"; displayName = $Id
       userType = $Type; accountEnabled = $Enabled
       assignedPlans = @(@{ service = $(if ($Exchange) { 'exchange' } else { 'teams' }); capabilityStatus = 'Enabled' }) }
}
$script:users = @((New-User 'admin'), (New-User 'presenter'), (New-User 'colleague'),
    (New-User 'guest' 'Guest'), (New-User 'disabled' 'Member' $false), (New-User 'teams' 'Member' $true $false))
$script:store = @{}
$script:requests = New-Object Collections.ArrayList
$script:grants = @()
$script:sent = @()
$script:event = $null
$script:sequence = 0
$script:failSend = $false
$script:rejectSend = $false
$script:replicateOnce = $false
$script:waits = 0
function Start-Sleep { param([int]$Seconds) $script:waits++ }
$script:failCalendar = $false
$script:failGrant = $false
$script:denySite = $false
function Invoke-WiqGraph {
    param([string]$Method, [string]$Path, $Body, [switch]$AllowNotFound, [string]$InputFile)
    [void]$script:requests.Add(@{ method = $Method; path = $Path; body = $Body })
    if ($Path -like '/users?*') {
        if ($Path -like '*page=2*') { return @{ value = @($script:users | Select-Object -Skip 2) } }
        return @{ value = @($script:users | Select-Object -First 2); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users?page=2' }
    }
    if ($Path -like '/sites/root/drive?*') {
        if ($script:denySite) { throw 'HTTP 403 site denied' }
        return @{ id = 'shared'; webUrl = 'https://contoso.invalid/sites/root' }
    }
    if ($Path -eq '/me/sendMail') {
        if ($script:replicateOnce) {
            $script:replicateOnce = $false
            $failure = New-Object InvalidOperationException 'HTTP 403 Send As not replicated'
            $failure.Data['HttpStatus'] = 403
            $failure.Data['GraphCode'] = 'ErrorSendAsDenied'
            throw $failure
        }
        if ($script:rejectSend) {
            $failure = New-Object InvalidOperationException 'HTTP 403 Send As not replicated'
            $failure.Data['HttpStatus'] = 403
            throw $failure
        }
        if ($script:failSend) { throw 'connection lost; unknown send outcome' }
        $script:sent += $Body
        return
    }
    if ($Path -eq '/me/events') {
        if ($script:failCalendar) { throw 'calendar connection lost; unknown outcome' }
        $script:event = @{ id = 'calendar-1'; subject = $Body.subject; attendees = $Body.attendees; start = $Body.start; end = $Body.end }
        return $script:event
    }
    if ($Path -eq '/me/events/calendar-1') {
        if ($Method -eq 'PATCH') {
            if ($Body.attendees) { $script:event.attendees = $Body.attendees }
            if ($Body.start) { $script:event.start = $Body.start; $script:event.end = $Body.end }
        }
        return $script:event
    }
    if ($Path -match '/permissions$') { return @{ value = $script:grants } }
    if ($Path -match '/invite$') {
        if ($script:failGrant) { return @{ value = @(@{ error = @{ code = 'grantFailed' } }) } }
        $p = @{ id = "perm-$($Body.recipients[0].objectId)"; roles = @('read'); grantedToV2 = @{ user = @{ id = $Body.recipients[0].objectId } } }
        $script:grants += $p
        return @{ value = @($p) }
    }
    if ($Method -eq 'POST' -and $Path -match '^(.*)/children$') {
        $key = "$($Matches[1]):/$($Body.name)"
        if ($script:store.ContainsKey($key)) { throw 'HTTP 409 folder already exists' }
        $script:sequence++
        $value = @{ id = "folder-$script:sequence"; folder = @{}; webUrl = "https://contoso.invalid/folders/$script:sequence" }
        # Empty Graph folder facets are valid objects, not PowerShell's falsey empty dictionaries.
        $value.folder = [pscustomobject]@{ childCount = 0 }
        $script:store[$key] = $value
        return $value
    }
    if ($Method -eq 'DELETE') {
        $id = ($Path -split '/')[-1]
        foreach ($key in @($script:store.Keys)) { if ($script:store[$key].id -eq $id) { $script:store.Remove($key) } }
        return
    }
    $key = ($Path -split '\?')[0] -replace ':/content$', ''
    if ($Method -eq 'PUT') {
        $script:store[$key] = @{ id = "file-$key"; data = [IO.File]::ReadAllText($InputFile); file = @{ hashes = @{ quickXorHash = (Get-QuickXor $InputFile) } } }
        return $script:store[$key]
    }
    if ($Method -eq 'GET') {
        if (-not $script:store.ContainsKey($key)) {
            if ($AllowNotFound) { return $null }
            throw "HTTP 404 $key"
        }
        if ($Path -match ':/content$') { return $script:store[$key].data }
        return $script:store[$key]
    }
    throw "Unexpected mock request: $Method $Path"
}
$root = Join-Path $PSScriptRoot '..\workiq-port-operations'
$cfg = ([IO.File]::ReadAllText((Join-Path $root 'tenant.json')) | ConvertFrom-Json).workIqShared
$item = @{ id = 'workiq-port-operations'; dir = (Resolve-Path $root).Path }
$mails = @(1..4 | ForEach-Object {
    @{ Subject = "[WIQ26-M0$_] Contoso evidence"; FromName = 'Teresa Sac'; FromAddress = 'teresa@contoso.invalid'; Body = "evidence $_" }
})
$senders = @{ 'Teresa Sac' = @{ address = 'teresa@contoso.invalid' } }
$setupArgs = @{ Config = $cfg; Item = $item; Me = $me; Mails = $mails; Senders = $senders; Language = 'en' }

Invoke-WorkIqSharedSetup @setupArgs -WhatIf
Assert (@($script:requests | Where-Object { $_.method -ne 'GET' }).Count -eq 0) 'WhatIf must never write.'
$script:denySite = $true
Must-Fail { Invoke-WorkIqSharedSetup @setupArgs } '403'
$script:denySite = $false
Assert ($script:store.Count -eq 0) 'Denied root library must not create a fallback destination.'

$links = @(Invoke-WorkIqSharedSetup @setupArgs)
Assert ($links.Count -eq 1 -and $links[0].url -like 'https://contoso.invalid/folders/*') 'Actual shared link must be returned.'
Assert ($links[0].result -eq 'new') 'First publication must not be misreported as unchanged.'
Assert ($script:grants.Count -eq 3) 'Only admin, presenter and eligible colleague may receive read grants.'
Assert (@($script:requests | Where-Object { $_.path -match '/invite$' -and ($_.body.roles.Count -ne 1 -or $_.body.roles[0] -ne 'read' -or $_.body.sendInvitation -ne $false -or $_.body.requireSignIn -ne $true) }).Count -eq 0) 'Only signed-in folder read grants without sharing notifications are allowed.'
Assert ($script:sent.Count -eq 12) 'Four primary-mailbox messages per eligible user.'
Assert (@($script:sent | Where-Object { $_.message.toRecipients[0].emailAddress.address -eq 'presenter@contoso.invalid' }).Count -eq 4) 'Non-admin presenter must get own mail evidence.'
Assert ($script:event.attendees.Count -eq 2) 'Presenter and colleague need real invitations.'
Assert (@($script:requests | Where-Object { $_.path -match '/users/.+/(drive|messages|events)' }).Count -eq 0) 'Must not assume delegated admin can write other private drives/mailboxes.'
$uploads = @($script:requests | Where-Object { $_.method -eq 'PUT' -and $_.path -match '\.docx:/content$' }).Count
Assert ($uploads -eq 48) 'Both language corpora must be published.'
$repeat = Invoke-WorkIqSharedSetup @setupArgs
Assert ($repeat.result -eq 'same') 'A stable rerun must be reported as unchanged.'
Assert ($script:sent.Count -eq 12 -and $script:grants.Count -eq 3) 'Rerun must not duplicate mail or grants.'
Assert (@($script:requests | Where-Object { $_.method -eq 'PUT' -and $_.path -match '\.docx:/content$' }).Count -eq $uploads) 'Unchanged document uploads must be skipped.'
Assert (@($script:requests | Where-Object { $_.method -eq 'POST' -and $_.path -eq '/me/events' }).Count -eq 1) 'Rerun must preserve the calendar event.'
$script:event.end.dateTime = [datetime]::UtcNow.AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ss')
$rescheduled = Invoke-WorkIqSharedSetup @setupArgs
Assert ($rescheduled.result -eq 'updated' -and [datetime]::Parse($script:event.end.dateTime) -gt [datetime]::UtcNow) 'A past rehearsal must move forward using the same event, not require manual recreation.'
Assert (@($script:requests | Where-Object { $_.method -eq 'POST' -and $_.path -eq '/me/events' }).Count -eq 1) 'Rescheduling must not create a duplicate meeting.'
$script:users += New-User 'new-colleague'
Invoke-WorkIqSharedSetup @setupArgs | Out-Null
Assert ($script:sent.Count -eq 16 -and $script:grants.Count -eq 4 -and $script:event.attendees.Count -eq 3) 'New eligible users need grant, missing mails and invitation without duplicating others.'

$stateKey = @($script:store.Keys | Where-Object { $_ -match 'state.json$' })[0]
$state = $script:store[$stateKey].data | ConvertFrom-Json
$state.tenantId = 'wrong-tenant'
$script:store[$stateKey].data = $state | ConvertTo-Json -Depth 30
Must-Fail { Invoke-WorkIqSharedSetup @setupArgs } 'ownership mismatch'
$state.tenantId = $script:tenant
$state.publisherId = 'another-publisher'
$script:store[$stateKey].data = $state | ConvertTo-Json -Depth 30
Must-Fail { Invoke-WorkIqSharedSetup @setupArgs } 'ownership mismatch'
$state.publisherId = $me.id
$script:store[$stateKey].data = $state | ConvertTo-Json -Depth 30

$script:users += New-User 'uncertain-recipient'
$script:failGrant = $true
Must-Fail { Invoke-WorkIqSharedSetup @setupArgs } 'incomplete read grant'
Assert ($script:sent.Count -eq 16) 'An incomplete permission grant must stop all new message submission.'
$script:failGrant = $false
$script:failSend = $true
Must-Fail { Invoke-WorkIqSharedSetup @setupArgs } 'unknown send outcome'
$script:failSend = $false
Must-Fail { Invoke-WorkIqSharedSetup @setupArgs } 'outcome unresolved'
Assert ($script:sent.Count -eq 16) 'An uncertain send must not be automatically retried.'
$markerKey = @($script:store.Keys | Where-Object { $_ -match 'demo-kit-owner.json$' })[0]
$script:store[$markerKey].data = '{"owner":"not-demo-kit","scenario":"WIQ26","version":1}'
Must-Fail { Invoke-WorkIqSharedSetup @setupArgs } 'ownership'
Assert (@($script:store.Keys | Where-Object { $_ -match 'installer-lock$' }).Count -eq 0) 'The exact lock must be cleaned up after failures.'
$external = @{ roles = @('read'); grantedToV2 = @{ siteUser = @{ loginName = 'i:0#.f|membership|presenter@contoso.invalid' } } }
Assert (Test-WiqReadPermission $external (New-User 'presenter')) 'SharePoint login identities must be recognized.'

$script:store[$markerKey].data = '{"owner":"DemoKit","scenario":"WIQ26","version":1}'
$script:store[$stateKey].data = $state | ConvertTo-Json -Depth 30
$script:users = @($script:users | Where-Object { $_.id -ne 'uncertain-recipient' }) + @((New-User 'rejected-recipient'))
$script:rejectSend = $true
Must-Fail { Invoke-WorkIqSharedSetup @setupArgs } '403'
$rejectedState = $script:store[$stateKey].data | ConvertFrom-Json
Assert (@($rejectedState.deliveries | Where-Object { $_.status -eq 'pending' }).Count -eq 0) 'An explicit Send As rejection must be retryable after replication, not incorrectly marked uncertain.'
$script:rejectSend = $false
Invoke-WorkIqSharedSetup @setupArgs | Out-Null
Assert ($script:sent.Count -eq 20) 'Rerun after a confirmed rejection must submit only the new recipient evidence.'

$privateKey = @($script:store.Keys | Where-Object { $_ -match '^/me/drive/root:/DemoKit' })[0]
$lockKey = "/me/drive/items/$($script:store[$privateKey].id):/installer-lock"
$script:store[$lockKey] = @{ id = 'another-run-lock'; folder = [pscustomobject]@{} }
Must-Fail { Invoke-WorkIqSharedSetup @setupArgs } 'lock could not be acquired'
Assert ($script:store.ContainsKey($lockKey)) 'Another process lock must never be removed.'
$script:store.Remove($lockKey)

$script:store = @{}; $script:grants = @(); $script:sent = @(); $script:requests.Clear()
$script:failCalendar = $true
Must-Fail { Invoke-WorkIqSharedSetup @setupArgs } 'calendar connection lost'
$script:failCalendar = $false
Must-Fail { Invoke-WorkIqSharedSetup @setupArgs } 'calendar creation outcome unresolved'
Assert (@($script:requests | Where-Object { $_.method -eq 'POST' -and $_.path -eq '/me/events' }).Count -eq 1) 'Uncertain event creation must not send duplicate invitations.'
$script:replicateOnce = $true
$sentBefore = $script:sent.Count
Send-WiqMessage @{ subject = 'replication fixture' } 1
Assert ($script:waits -eq 1 -and $script:sent.Count -eq ($sentBefore + 1)) 'Only a confirmed Send As rejection may be retried while the grant replicates.'
$script:replicateOnce = $true
Must-Fail { Send-WiqMessage @{ subject = 'no-wait fixture' } 0 } '403'
Assert ($script:waits -eq 1) 'An expired replication wait must stop explicitly.'
Write-Host 'PASS: shared files, primary mail, calendar, eligibility, pagination, reruns, new users, WhatIf, explicit HTTP failures, ownership and uncertain-send protection.'
