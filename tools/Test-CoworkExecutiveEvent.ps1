$ErrorActionPreference = 'Stop'
$paths = @('tenant.ps1', 'tools\ExecutiveEvent.ps1', 'tools\Reset-CoworkExecutiveEvent.ps1')
$root = Split-Path $PSScriptRoot -Parent
foreach ($relative in $paths) {
    $tokens = $null; $errors = $null
    [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $root $relative), [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw ($errors | Out-String) }
}
. (Join-Path $PSScriptRoot 'ExecutiveEvent.ps1')
$config = (Get-Content (Join-Path $root 'cowork-executive-event\tenant.json') -Raw -Encoding UTF8 | ConvertFrom-Json).executiveEvent
$settings = [pscustomobject]@{
    tenantId = '11111111-1111-1111-1111-111111111111'; teamId = '22222222-2222-2222-2222-222222222222'
    account = $config.account; channelId = '19:fixture@thread.tacv2'
}
$context = [pscustomobject]@{ TenantId = $settings.tenantId; Account = $settings.account }
$me = [pscustomobject]@{ userPrincipalName = $settings.account; mail = $settings.account }
function Must-Fail([scriptblock]$Action, [string]$Message) {
    try { & $Action }
    catch {
        if ("$_" -notlike "*$Message*") { throw "Wrong error; expected '$Message', got $_" }
        return
    }
    throw "Expected a failure containing $Message"
}
Assert-ExecutiveEventIdentity $config $settings $context $me
Must-Fail { Assert-ExecutiveEventIdentity $config $settings $null $me } 'blocked'
Must-Fail { Assert-ExecutiveEventIdentity $config $settings ([pscustomobject]@{ TenantId = 'wrong'; Account = $settings.account }) $me } 'blocked'
Must-Fail { Assert-ExecutiveEventIdentity $config $settings $context ([pscustomobject]@{ userPrincipalName = 'personal@example.invalid'; mail = $settings.account }) } 'blocked'
$recipients = @{}
foreach ($person in $config.people.PSObject.Properties) { $recipients[$person.Name] = [pscustomobject]@{ address = $person.Value; created = $false } }
$goodRecipient = $recipients['Teresa Sac']
function Get-MgContext { $context }

$graphImplementation = ${function:Invoke-ExecutiveEventGraph}
function Invoke-MgGraphRequest {
    param($Method, $Uri, $OutputType, [switch]$SkipHttpErrorCheck, $StatusCodeVariable, $Headers)
    $script:lastHeaders = $Headers
    Set-Variable -Name $StatusCodeVariable -Value 200 -Scope 1
    @{}
}
[void](& $graphImplementation GET 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skip=50')
if ($script:lastHeaders.Prefer -cne 'IdType="ImmutableId"') { throw 'Mail pagination lost immutable IDs.' }
[void](& $graphImplementation GET 'https://graph.microsoft.com/v1.0/me/calendarView?$skip=100')
if ($script:lastHeaders.Prefer -cne 'outlook.timezone="W. Europe Standard Time"') { throw 'Calendar pagination lost the Frankfurt time zone.' }
Must-Fail { & $graphImplementation GET 'https://example.invalid/me/messages' } 'unexpected Graph URL'

$oldLocal = $env:LOCALAPPDATA
$temp = Join-Path ([IO.Path]::GetTempPath()) ("DemoKit-LEF26-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null
$env:LOCALAPPDATA = $temp
try {
    Must-Fail { Get-ExecutiveEventSettings $config } 'private settings'
    New-Item -ItemType Directory -Path (Join-Path $temp 'DemoKit') | Out-Null
    [IO.File]::WriteAllText((Join-Path $temp 'DemoKit\executive-event-settings.json'), (ConvertTo-Json $settings))
    $loaded = Get-ExecutiveEventSettings $config
    if ($loaded.channelId -cne $settings.channelId) { throw 'Settings did not round-trip.' }
    if ((Get-ExecutiveEventTransaction 'same') -cne (Get-ExecutiveEventTransaction 'same') -or
        (Get-ExecutiveEventTransaction 'same') -ceq (Get-ExecutiveEventTransaction 'different')) { throw 'Transaction IDs are not deterministic and distinct.' }

    $script:objects = @{}
    $script:writes = 0
    $script:sequence = 0
    $script:sendCount = 0
    $script:failSend = $false
    $script:forcedRecipients = @()
    $script:createdRuns = @()
    $script:parents = @{}
    foreach ($reply in $config.mailReplies) {
        $subject = Get-ExecutiveEventText $reply.parentSubject 'de'
        $script:parents[$subject] = @{
            id = "parent-$($reply.key)"; subject = $subject; conversationId = "conversation-$($reply.key)"
            from = @{ emailAddress = @{ address = $config.people.($reply.parentSender) } }
        }
    }
    function Invoke-ExecutiveEventGraph {
        param([string]$Method, [string]$Path, $Body, [string]$InputFile, [switch]$Optional)
        if ($Method -ne 'GET') { $script:writes++ }
        if ($Method -eq 'GET' -and $Path -eq '/me?$select=userPrincipalName,mail') { return $me }
        if ($Method -eq 'GET' -and $Path -like '/me/drive/root:/*/baseline/de/Action_Register.xlsx') {
            return @{ id = 'baseline'; file = @{ mimeType = 'application/octet-stream' }; '@microsoft.graph.downloadUrl' = 'https://fixture.invalid/baseline' }
        }
        if ($Method -eq 'GET' -and $Path -like '/me/drive/root:/*/runs/de') {
            return @{ id = 'runs-parent'; folder = @{ childCount = 4 } }
        }
        if ($Method -eq 'GET' -and $Path -like '/me/calendarView?*') { return @{ value = @() } }
        if ($Method -eq 'GET' -and $Path -like '/me/mailFolders/inbox/messages?*') {
            $decoded = [uri]::UnescapeDataString($Path)
            return @{ value = @($script:parents.Values | Where-Object { $decoded.Contains($_.subject) }) }
        }
        if ($Method -eq 'GET') {
            if ($script:objects.ContainsKey($Path)) { return $script:objects[$Path] }
            if ($Optional) { return $null }
            throw "Unexpected mock GET $Path"
        }
        $script:sequence++
        $id = "object-$($script:sequence)"
        if ($Method -eq 'PUT' -and $InputFile) {
            $object = @{ id = $id; cTag = "content-$id"; webUrl = "https://fixture.sharepoint.com/$id"; file = @{} }
            $script:objects[($Path -replace ':/content$', '')] = $object
            return $object
        }
        if ($Method -eq 'POST' -and $Path -like '*/children') {
            if ($Body.name -like 'fresh-*') { $script:createdRuns += $Body.name }
            $object = @{ id = $id; folder = @{ childCount = 0 }; webUrl = "https://fixture.sharepoint.com/$id" }
            $script:objects["/drives/drive/items/parent:/$($Body.name)"] = $object
            return $object
        }
        if ($Method -eq 'POST' -and $Path -eq '/me/events') {
            $object = @{}; foreach ($key in $Body.Keys) { $object[$key] = $Body[$key] }
            $object.id = $id; $object.webLink = "https://outlook.office.com/$id"
            $script:objects["/me/events/$id"] = $object
            return $object
        }
        if ($Method -eq 'POST' -and ($Path.EndsWith('/messages') -or $Path.EndsWith('/replies'))) {
            $object = @{ id = $id; body = $Body.body; webUrl = "https://teams.microsoft.com/$id" }
            $script:objects["$Path/$id"] = $object
            return $object
        }
        if ($Method -eq 'POST' -and $Path.EndsWith('/createReply')) {
            $parentId = $Path.Split('/')[-2]
            $parent = @($script:parents.Values | Where-Object { $_.id -ceq $parentId })[0]
            $to = @(@{ emailAddress = @{ address = $parent.from.emailAddress.address } })
            if ($script:forcedRecipients.Count) { $to = $script:forcedRecipients }
            $object = @{
                id = $id; conversationId = $parent.conversationId; isDraft = $true
                from = @{ emailAddress = @{ address = $config.account } }; body = @{ content = $Body.comment }
                toRecipients = $to; ccRecipients = @(); bccRecipients = @()
            }
            $script:objects["/me/messages/$id"] = $object
            return $object
        }
        if ($Method -eq 'POST' -and $Path.EndsWith('/send')) {
            $script:sendCount++
            if ($script:failSend) { throw 'Simulated uncertain send result' }
            $messagePath = $Path -replace '/send$', ''
            $script:objects[$messagePath].isDraft = $false
            return
        }
        throw "Unexpected mock $Method $Path"
    }
    $destination = [pscustomobject]@{
        settings = $settings; folder = @{ id = 'parent'; parentReference = @{ driveId = 'drive' } }
        channelPath = '/teams/team/channels/channel'; ledger = $null
    }
    $item = @{ dir = (Join-Path $root 'cowork-executive-event') }
    $recipients['Teresa Sac'] = [pscustomobject]@{ address = 'external@example.invalid'; created = $false }
    Must-Fail { Invoke-ExecutiveEventSetup $config $item $destination $recipients de } 'attendee'
    $recipients['Teresa Sac'] = [pscustomobject]@{ address = "wrong@$($config.domain)"; created = $false }
    Must-Fail { Invoke-ExecutiveEventSetup $config $item $destination $recipients de } 'attendee'
    $recipients['Teresa Sac'] = $goodRecipient
    [void](Invoke-ExecutiveEventSetup $config $item $destination $recipients de -WhatIf)
    if ($script:writes) { throw 'Denied recipient or WhatIf caused writes.' }
    $links = @(Invoke-ExecutiveEventSetup $config $item $destination $recipients de)
    if ($links.Count -ne 10) { throw "Expected 10 setup links, got $($links.Count)" }
    $destination.ledger = Get-ExecutiveEventLedger $settings de
    $snapshotWrites = $script:writes
    [void](Invoke-ExecutiveEventSetup $config $item $destination $recipients de)
    if ($script:writes -ne $snapshotWrites) { throw 'Setup rerun performed writes instead of reusing owned objects.' }
    Invoke-ExecutiveEventMailReplies $config $settings de -WhatIf
    if ($script:writes -ne $snapshotWrites) { throw 'Mail WhatIf caused writes.' }
    Invoke-ExecutiveEventMailReplies $config $settings de
    if ($script:sendCount -ne 2) { throw 'Expected exactly two internal setup replies.' }
    $snapshotWrites = $script:writes
    Invoke-ExecutiveEventMailReplies $config $settings de
    if ($script:writes -ne $snapshotWrites -or $script:sendCount -ne 2) { throw 'Mail rerun duplicated work or sent a second time.' }

    $entry = $destination.ledger.files[0]
    $filePath = "/drives/drive/items/$($destination.ledger.rootId):/$($entry.name)"
    $script:objects[$filePath].cTag = 'changed-outside-setup'
    Must-Fail { Invoke-ExecutiveEventSetup $config $item $destination $recipients de } 'baseline'
    if ($script:writes -ne $snapshotWrites) { throw 'Changed baseline caused a write.' }
    $script:objects[$filePath].cTag = $entry.cTag

    $resetAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Reset-CoworkExecutiveEvent.ps1'), [ref]$tokens, [ref]$errors)
    $body = ($resetAst.EndBlock.Statements | Where-Object {
        $_.Extent.Text -notlike 'Import-Module *' -and $_.Extent.Text -notlike '. (Join-Path *'
    } | ForEach-Object { $_.Extent.Text }) -join "`n"
    $body = $body.Replace('$PSScriptRoot', "'$($PSScriptRoot.Replace("'", "''"))'")
    Set-Item Function:Invoke-ResetUnderTest -Value ([scriptblock]::Create("[CmdletBinding(SupportsShouldProcess)]param([string]`$Language = 'de')`n$body"))
    $script:corruptBaseline = $false
    function Invoke-WebRequest {
        param($Uri, $OutFile, [switch]$UseBasicParsing)
        if ($Uri -ne 'https://fixture.invalid/baseline') { throw 'Unexpected reset download.' }
        if ($script:corruptBaseline) { [IO.File]::WriteAllText($OutFile, 'corrupt fixture') }
        else { [IO.File]::WriteAllBytes($OutFile, [IO.File]::ReadAllBytes((Join-Path $root 'cowork-executive-event\baseline\de\Action_Register.xlsx'))) }
    }
    $snapshotWrites = $script:writes
    Invoke-ResetUnderTest -WhatIf
    if ($script:writes -ne $snapshotWrites) { throw 'Reset WhatIf performed tenant writes.' }
    Invoke-ResetUnderTest -Confirm:$false
    if ($script:writes -ne $snapshotWrites + 3) { throw 'Reset must create exactly one folder, workbook and output folder.' }
    $active = Get-Content (Join-Path $temp "DemoKit\executive-event\$($settings.tenantId)\active-de.json") -Raw | ConvertFrom-Json
    if ($active.scenario -ne 'LEF26' -or $active.run -notlike 'fresh-*' -or -not $active.input -or -not $active.outputs) { throw 'Reset did not persist its new active-run links.' }
    Invoke-ResetUnderTest -Confirm:$false
    if ($script:createdRuns.Count -ne 2 -or $script:createdRuns[0] -ceq $script:createdRuns[1]) { throw 'Reset reused a previous run namespace.' }
    $script:corruptBaseline = $true
    $snapshotWrites = $script:writes
    Must-Fail { Invoke-ResetUnderTest -Confirm:$false } 'differs'
    if ($script:writes -ne $snapshotWrites) { throw 'Reset copied a changed baseline into the tenant.' }
    $context.Account = 'personal@example.invalid'
    Must-Fail { Invoke-ResetUnderTest -Confirm:$false } 'blocked'
    if ($script:writes -ne $snapshotWrites) { throw 'Reset wrote under an unauthorized identity.' }
    $context.Account = $settings.account

    # A separate language ledger exercises fail-closed recipient handling.
    $destination.ledger = $null
    [void](Invoke-ExecutiveEventSetup $config $item $destination $recipients en)
    foreach ($reply in $config.mailReplies) {
        $subject = Get-ExecutiveEventText $reply.parentSubject en
        $script:parents[$subject] = @{
            id = "en-parent-$($reply.key)"; subject = $subject; conversationId = "en-conversation-$($reply.key)"
            from = @{ emailAddress = @{ address = $config.people.($reply.parentSender) } }
        }
    }
    $script:forcedRecipients = @(@{ emailAddress = @{ address = 'external@example.invalid' } })
    Must-Fail { Invoke-ExecutiveEventMailReplies $config $settings en } 'unexpected recipients'
    if ($script:sendCount -ne 2) { throw 'External reply recipient was sent to.' }
    $script:forcedRecipients = @()
    $en = Get-ExecutiveEventLedger $settings en
    $draftPath = "/me/messages/$($en.mailReplies[0].id)"
    $script:objects[$draftPath].toRecipients = @(@{ emailAddress = @{ address = $config.people.'Teresa Sac' } })
    $script:failSend = $true
    Must-Fail { Invoke-ExecutiveEventMailReplies $config $settings en } 'uncertain send'
    $sentAttempts = $script:sendCount
    Must-Fail { Invoke-ExecutiveEventMailReplies $config $settings en } 'already submitted'
    if ($script:sendCount -ne $sentAttempts) { throw 'Uncertain send was retried automatically.' }
    Write-Host 'LEF26 offline checks passed: identity/recipient denial, dry-run, fixed calendar ownership, shared-file protection, Teams reuse, real reply API shape, idempotency, uncertain-send handling and hash-checked non-destructive reset. No real tenant calls.'
} finally {
    $env:LOCALAPPDATA = $oldLocal
    # Remove only files created beneath this exact uniquely named test directory.
    Remove-Item -LiteralPath $temp -Recurse -Force
}
