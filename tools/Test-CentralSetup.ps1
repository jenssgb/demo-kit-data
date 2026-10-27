$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WorkIqShared.ps1')
. (Join-Path $PSScriptRoot 'PresenterSetup.ps1')
function Assert($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Must-Fail([scriptblock]$Action, [string]$Pattern) {
    try { & $Action | Out-Null } catch { Assert ("$_" -match $Pattern) "Wrong failure: $_"; return }
    throw "Expected failure: $Pattern"
}
function New-User($Id, $Type = 'Member', $Enabled = $true, $Exchange = $true) {
    [pscustomobject]@{ id = $Id; displayName = $Id; userPrincipalName = "$Id@contoso.invalid"; mail = "$Id@contoso.invalid"
        accountEnabled = $Enabled; userType = $Type
        assignedPlans = @(@{ service = $(if ($Exchange) { 'exchange' } else { 'teams' }); capabilityStatus = 'Enabled' }) }
}
$operator = New-User 'admin'
$presenter = New-User 'presenter'
$colleague = New-User 'colleague'
$script:users = @($operator, $presenter, $colleague, (New-User 'guest' 'Guest'), (New-User 'disabled' 'Member' $false), (New-User 'teams' 'Member' $true $false))
$script:requests = New-Object Collections.ArrayList
$script:folders = @{}
$script:stateFiles = @{}
$script:permissions = @()
$script:calendar = @()
$script:sent = @()
$script:sendFailure = ''
$script:grantFailure = $false
$script:notebooks = @()
$script:choices = New-Object Collections.Queue
function Read-Host { param($Prompt) $script:lastPrompt = $Prompt; $script:choices.Dequeue() }
function Get-MgContext { @{ TenantId = '11111111-1111-1111-1111-111111111111'; Account = $operator.userPrincipalName } }
function Row { param($Tag, $Text, $Color) }
function Say { param($Text, $Color) }
function Invoke-WiqGraph {
    param([string]$Method, [string]$Path, $Body, [switch]$AllowNotFound, [string]$InputFile)
    [void]$script:requests.Add(@{ method = $Method; path = $Path; body = $Body })
    if ($Path -match '^/users\?') {
        if ($Path -like '*page=2*') { return @{ value = @($script:users | Select-Object -Skip 2) } }
        return @{ value = @($script:users | Select-Object -First 2); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users?page=2' }
    }
    if ($Path -match '^/users/([^?]+)\?') {
        $user = $script:users | Where-Object { $_.id -eq $Matches[1] }
        if ($script:revalidateDisabled) { return @{ id = $user.id; accountEnabled = $false } }
        return $user
    }
    if ($Path -like '/me/drive/root/search*') { return @{ value = $script:notebooks } }
    if ($Path -like '*/permissions') { return @{ value = $script:permissions } }
    if ($Path -like '*/invite') {
        if ($script:grantFailure) { return @{ value = @(@{ error = @{ code = 'denied' } }) } }
        $grant = @{ id = 'grant'; roles = $Body.roles; grantedToV2 = @{ user = @{ id = $Body.recipients[0].objectId } } }
        $script:permissions += $grant
        return @{ value = @($grant) }
    }
    if ($Path -like '/me/calendarView?*') { return @{ value = $script:calendar } }
    if ($Path -eq '/me/events' -and $Method -eq 'POST') {
        $event = @{ id = 'event'; subject = $Body.subject; attendees = $Body.attendees; isOrganizer = $true }
        $script:calendar += $event
        return $event
    }
    if ($Path -eq '/me/events/event' -and $Method -eq 'PATCH') {
        $script:calendar[0].attendees = $Body.attendees
        return $script:calendar[0]
    }
    if ($Path -eq '/me/sendMail') {
        if ($script:sendFailure -eq 'timeout') { throw 'Unknown submission outcome: connection lost' }
        if ($script:sendFailure -eq 'rejected') {
            $failure = New-Object InvalidOperationException 'HTTP 403 rejected'
            $failure.Data['HttpStatus'] = 403
            throw $failure
        }
        $script:sent += $Body
        return
    }
    if ($Method -eq 'POST' -and $Path -match '^(.*)/children$') {
        $key = "$($Matches[1]):/$($Body.name)"
        if ($script:folders.ContainsKey($key)) { throw 'HTTP 409: installer-lock already exists' }
        $folder = @{ id = "folder-$($script:folders.Count)"; folder = [pscustomobject]@{} }
        $script:folders[$key] = $folder
        return $folder
    }
    if ($Method -eq 'DELETE') {
        foreach ($key in @($script:folders.Keys)) {
            if ($Path -eq "/me/drive/items/$($script:folders[$key].id)") { $script:folders.Remove($key); return }
        }
        throw "Unexpected delete $Path"
    }
    $key = $Path -replace ':/content$', ''
    if ($Method -eq 'PUT' -and $InputFile) { $script:stateFiles[$key] = [IO.File]::ReadAllText($InputFile); return }
    if ($Method -eq 'GET') {
        if ($script:folders.ContainsKey($Path)) { return $script:folders[$Path] }
        if ($script:stateFiles.ContainsKey($key)) {
            if ($Path -match ':/content$') { return $script:stateFiles[$key] }
            return @{ id = 'state-file' }
        }
        if ($AllowNotFound) { return $null }
    }
    throw "Unexpected Graph request $Method $Path"
}
$script:choices.Enqueue('not-a-number'); $script:choices.Enqueue('0'); $script:choices.Enqueue('3')
$selected = Select-DemoPresenter
Assert ($selected.id -eq 'presenter') 'The eligible, sorted third user should be selected, not a guest or Teams-only user.'
Assert (@($script:requests | Where-Object { $_.method -ne 'GET' }).Count -eq 0) 'Selection must not write tenant data.'
$script:choices.Enqueue('')
Must-Fail { Select-DemoPresenter } 'cancelled before tenant changes'
$script:choices.Enqueue('1')
$script:revalidateDisabled = $true
Must-Fail { Select-DemoPresenter } 'no longer'
$script:revalidateDisabled = $false
$script:choices.Enqueue('1')
Assert ((Select-DemoPresenter -Language de).id -eq 'admin' -and $script:lastPrompt -match 'Nummer des Demobenutzers') 'German selection must follow the requested language.'
Write-Host 'Selection: paged eligible users, invalid choices, cancellation and selected-user revalidation.'

$folder = @{ id = 'demo-folder'; folder = [pscustomobject]@{}; webUrl = 'https://contoso.invalid/demo' }
Assert ((Grant-DemoPresenterAccess $folder $presenter -WhatIf) -eq 'new') 'WhatIf must show an editing grant.'
Assert ($script:permissions.Count -eq 0) 'WhatIf must not grant access.'
Assert ((Grant-DemoPresenterAccess $folder $presenter) -eq 'new') 'First grant missing.'
$invite = @($script:requests | Where-Object { $_.path -like '*/invite' })[-1]
Assert ($invite.body.recipients[0].objectId -eq $presenter.id -and
    $invite.body.roles[0] -eq 'write' -and $invite.body.requireSignIn -and -not $invite.body.sendInvitation) 'Only the chosen authenticated user may receive editing access.'
Assert ((Grant-DemoPresenterAccess $folder $presenter) -eq 'same') 'Existing grant should be reused.'
$script:grantFailure = $true
Must-Fail { Grant-DemoPresenterAccess $folder $colleague } 'not fully granted'
$script:grantFailure = $false
Must-Fail { Grant-DemoPresenterAccess @{ id = 'root-file'; file = @{} } $presenter } 'Only an identified'
Must-Fail { Share-DemoNotebook 'Notebook' $presenter } 'sharing is not confirmed'
$script:notebooks = @(@{ id = 'notebook'; name = 'Notebook'; package = @{ type = 'oneNote' } })
Assert ((Share-DemoNotebook 'Notebook' $presenter) -eq 'same') 'Exact notebook package should be shareable.'
$script:notebooks += $script:notebooks[0]
Must-Fail { Share-DemoNotebook 'Notebook' $presenter } 'one exact'
Write-Host 'Sharing: authenticated folder-only editing, idempotence, partial failure and exact notebook identity.'

$payload = @{ subject = 'Contoso rehearsal'; attendees = @(@{ emailAddress = @{ address = 'cast@contoso.invalid' } }); transactionId = 'stable' }
Assert ((Sync-DemoPresenterInvitation $payload $presenter -WhatIf) -eq 'new') 'Expected invitation preview.'
Assert ($script:calendar.Count -eq 0) 'WhatIf created a meeting.'
Assert ((Sync-DemoPresenterInvitation $payload $presenter) -eq 'new') 'Expected new invitation.'
Assert ((Sync-DemoPresenterInvitation $payload $presenter) -eq 'same') 'Rerun must not duplicate invitations.'
Assert ((Sync-DemoPresenterInvitation $payload $colleague) -eq 'updated') 'A new presenter must be added to the existing meeting.'
Assert (@($script:calendar[0].attendees).Count -eq 3) 'Cast and both presenter attendees must be retained.'
$script:calendar += $script:calendar[0]
Must-Fail { Sync-DemoPresenterInvitation $payload $presenter } 'Multiple organizer events'
$script:calendar = @()
Write-Host 'Meetings: preview, creation, same-target rerun and adding another target without removing attendees.'

$payload = @{ message = @{ subject = 'Contoso update'; from = @{ emailAddress = @{ address = 'sender@contoso.invalid' } }
    toRecipients = @(@{ emailAddress = @{ address = $presenter.mail } }) }; saveToSentItems = $false }
$writes = @($script:requests | Where-Object { $_.method -ne 'GET' }).Count
Assert ((Send-DemoPresenterMessage $payload $operator $presenter 'en' -WhatIf) -eq 'new') 'Expected mail preview.'
Assert (@($script:requests | Where-Object { $_.method -ne 'GET' }).Count -eq $writes) 'WhatIf wrote delivery state.'
Assert ((Send-DemoPresenterMessage $payload $operator $presenter 'en') -eq 'new') 'First submission missing.'
Assert ((Send-DemoPresenterMessage $payload $operator $presenter 'en') -eq 'same') 'Submitted mail must not be sent twice.'
Assert ($script:sent.Count -eq 1 -and $script:sent[0].saveToSentItems) 'Submission must save operator Sent Items for diagnosis.'
Must-Fail { Send-DemoPresenterMessage $payload $operator $colleague 'en' } 'recipient does not match'
$payload.message.toRecipients[0].emailAddress.address = $colleague.mail
Assert ((Send-DemoPresenterMessage $payload $operator $colleague 'en') -eq 'new') 'Another target needs independent delivery state.'
$payload.message.subject = 'Contoso uncertain'
$script:sendFailure = 'timeout'
Must-Fail { Send-DemoPresenterMessage $payload $operator $colleague 'en' } 'Unknown submission outcome'
$script:sendFailure = ''
Must-Fail { Send-DemoPresenterMessage $payload $operator $colleague 'en' } 'no automatic resend'
$payload.message.subject = 'Contoso rejected'
$script:sendFailure = 'rejected'
Must-Fail { Send-DemoPresenterMessage $payload $operator $colleague 'en' -WaitMinutes 0 } '403 rejected'
$script:sendFailure = ''
Assert ((Send-DemoPresenterMessage $payload $operator $colleague 'en') -eq 'new') 'An explicit rejection may be safely retried.'
Assert (@($script:folders.Keys | Where-Object { $_ -like '*installer-lock' }).Count -eq 0) 'Delivery locks must be cleaned after success and failure.'
$savedPath = @($script:stateFiles.Keys | Where-Object { ($script:stateFiles[$_] | ConvertFrom-Json).recipientId -eq $colleague.id })[0]
$savedState = $script:stateFiles[$savedPath]
$foreign = $savedState | ConvertFrom-Json
$foreign.publisherId = 'another-publisher'
$script:stateFiles[$savedPath] = $foreign | ConvertTo-Json -Depth 20
Must-Fail { Send-DemoPresenterMessage $payload $operator $colleague 'en' } 'does not belong'
$script:stateFiles[$savedPath] = $savedState
Write-Host 'Mail: target-scoped ledger, preview, saved Sent Items, unknown-outcome protection and explicit-rejection retry.'

$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\tenant.ps1'), [ref]$tokens, [ref]$errors)
Assert (-not $errors.Count) ($errors | Out-String)
foreach ($name in 'New-DemoSeedMessage', 'T', 'L', 'Count') {
    $fn = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    . ([scriptblock]::Create($fn.Extent.Text))
}
$centralSetup = $true
$configs = @(
    @{ Item = @{ id = 'private' }; Config = @{ executiveEvent = @{} } },
    @{ Item = @{ id = 'locked' }; Config = @{ presenterAccount = 'owner@contoso.invalid' } },
    @{ Item = @{ id = 'generic' }; Config = @{} },
    @{ Item = @{ id = 'shared' }; Config = @{ workIqShared = @{ scenario = 'WIQ26' } } }
)
$configs[0].Config.executiveEvent = @{ scenario = 'private' }
$Items = @($configs | ForEach-Object { $_.Item })
$mails = @(@{ Demo = 'private' }, @{ Demo = 'locked' }, @{ Demo = 'generic' })
$Links = @{ manual = @(); demos = [ordered]@{ private = @{}; locked = @{}; generic = @{}; shared = @{} } }
$guard = $ast.Find({ param($n) $n -is [Management.Automation.Language.IfStatementAst] -and
    $n.Extent.Text.StartsWith('if ($centralSetup)') -and $n.Extent.Text -match '\$centralBlocked = @' }, $true)
. ([scriptblock]::Create($guard.Extent.Text))
Assert ($configs.Count -eq 2 -and $Items.Count -eq 2 -and $mails.Count -eq 1 -and
    @($configs | ForEach-Object { $_.Item.id }) -contains 'shared' -and
    $Links.manual.Count -eq 2 -and -not $Links.demos.Contains('private')) 'Central setup must isolate private scenarios without suppressing WIQ26.'
$selection = $ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Select-DemoPresenter' }, $true)
$privateWrite = $ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Initialize-ExecutiveEventState' }, $true)
Assert ($selection.Extent.StartOffset -lt $guard.Extent.StartOffset -and $guard.Extent.StartOffset -lt $privateWrite.Extent.StartOffset) 'Selection and privacy filter must precede private scenario writes.'
$me = $operator; $Language = 'en'; $WhatIf = $false; $de = $false; $SendAsWaitMinutes = 0
$report = [ordered]@{}
$senders = @{ Sender = @{ address = 'sender@contoso.invalid' } }
$configs = @(@{ Item = @{ id = 'generic' }; Config = @{} })
$mails = @(@{ Demo = 'generic'; FromName = 'Sender'; Subject = 'Contoso routed'; Body = 'Scenario evidence' })
function G { throw 'Central setup must not read the operator inbox to deduplicate target mail.' }
$loop = $ast.Find({ param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith('if ($mails.Count)') }, $true)
. ([scriptblock]::Create($loop.Extent.Text))
Assert ($report.Mails.new -eq 1 -and $script:sent[-1].message.toRecipients[0].emailAddress.address -eq $presenter.mail) 'Actual installer loop must route to the selected presenter.'
Assert (@($script:requests | Where-Object { $_.path -match '^/users/.*/(mail|events|drive)' }).Count -eq 0) 'No direct impersonation or target mailbox reads are permitted.'
$summary = $ast.Find({ param($n) $n -is [Management.Automation.Language.IfStatementAst] -and
    $n.Extent.Text.StartsWith('if ($centralSetup)') -and $n.Extent.Text -match '\$rows = @' }, $true)
$Items = @(@{ id = 'generic'; drivePath = 'Demo-Kit/App/generic' })
$Links = @{ manual = @(); demos = @{ generic = @{ presenterShared = $true; web = 'https://contoso.invalid/demo' } } }
$Language = 'de'; $de = $false; $myMail = $operator.mail
. ([scriptblock]::Create($summary.Extent.Text))
Assert ($script:sent[-1].message.subject -ceq 'Demo Kit - Links und Vorbereitung' -and
    $script:sent[-1].message.body.content -match 'Demo-Links und Vorbereitung') 'Links email must follow -Language rather than the Windows culture.'
Write-Host 'Integration: selected recipient routing, no target mailbox access, protected private scenarios and unchanged shared publishing.'
