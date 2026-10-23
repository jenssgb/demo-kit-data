# Offline only: all Graph and web requests below are strict test doubles.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
function Must-Fail([scriptblock]$Action, [string]$Expected) {
    try { & $Action }
    catch { if ("$_" -notlike "*$Expected*") { throw "Expected '$Expected', got: $_" }; return }
    throw "Expected failure: $Expected"
}
foreach ($relative in @('tenant.ps1', 'tools\DemoCalendar.ps1', 'tools\Reset-CoworkAssistantDay.ps1')) {
    $tokens = $null; $errors = $null
    [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $root $relative), [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw ($errors | Out-String) }
}
. (Join-Path $PSScriptRoot 'DemoCalendar.ps1')
$config = Get-Content (Join-Path $root 'cowork-assistant-day\tenant.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$context = [pscustomobject]@{ Account = $config.presenterAccount; TenantId = '11111111-1111-1111-1111-111111111111' }
$me = [pscustomobject]@{ userPrincipalName = $context.Account; mail = $context.Account }
Assert-DemoPresenter $config.presenterAccount $context $me
Must-Fail { Assert-DemoPresenter $config.presenterAccount $null $me } 'blocked'
Must-Fail { Assert-DemoPresenter $config.presenterAccount $context ([pscustomobject]@{ userPrincipalName = 'other@example.invalid'; mail = $context.Account }) } 'blocked'
Must-Fail { Assert-DemoPresenter $config.presenterAccount ([pscustomobject]@{ Account = 'other@example.invalid'; TenantId = $context.TenantId }) $me } 'blocked'
Must-Fail { Invoke-DemoGraph GET 'https://example.invalid/me/events' } 'only permits'
Must-Fail { Invoke-DemoGraph DELETE '/users/other/events/1' } 'only permits'

$script:events = @()
$script:posts = 0
$script:failPost = $false
$script:pagination = $false
$request = {
    param($Method, $Path, $Body)
    if ($Method -eq 'GET' -and $Path -like '/me/events?*') {
        if ($script:pagination -and $Path -notlike '*skiptoken*') {
            return @{ value = @(); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/me/events?$skiptoken=page2' }
        }
        return @{ value = @($script:events) }
    }
    if ($Method -eq 'POST' -and $Path -eq '/me/events') {
        $script:posts++
        if ($script:failPost) { throw 'injected POST failure' }
        $saved = $Body | ConvertTo-Json -Depth 15 | ConvertFrom-Json
        $saved | Add-Member NoteProperty id 'owned-event'
        $script:events += $saved
        return $saved
    }
    throw "Unexpected mock calendar request $Method $Path"
}
$event = $config.events[0]
$plan = Get-DemoEventPlan $event 'de'
if ($plan.startUtc.ToString('yyyy-MM-ddTHH:mm') -cne '2026-10-19T08:00') { throw 'Fixed Frankfurt time not converted to UTC.' }
$bad = $event | ConvertTo-Json -Depth 10 | ConvertFrom-Json
$bad.attendees = @('external@example.invalid')
Must-Fail { Get-DemoEventPlan $bad 'de' } 'no attendees'
$bad.attendees = @(); $bad.fixedDate = '2026-10-25'; $bad.start = '02:30'
Must-Fail { Get-DemoEventPlan $bad 'de' } 'ambiguous'
$bad.fixedDate = '2026-02-30'
Must-Fail { Get-DemoEventPlan $bad 'de' } 'DateTime'
$first = Sync-DemoFixedEvent $event 'de' $request -WhatIf
if ($first -ne 'new' -or $script:posts) { throw 'WhatIf wrote an event.' }
[void](Sync-DemoFixedEvent $event 'de' $request)
if ($script:posts -ne 1) { throw 'First setup did not create exactly one event.' }
$script:pagination = $true
if ((Sync-DemoFixedEvent $event 'de' $request) -ne 'same' -or $script:posts -ne 1) { throw 'Repeat setup wrote a duplicate.' }
$script:events[0].start = [pscustomobject]@{ dateTime = '2026-10-19T08:00:00.0000000'; timeZone = 'UTC' }
$script:events[0].end = [pscustomobject]@{ dateTime = '2026-10-19T08:45:00.0000000'; timeZone = 'UTC' }
if ((Sync-DemoFixedEvent $event 'de' $request) -ne 'same') { throw 'Equivalent Graph UTC response rejected.' }
$script:events[0].start.dateTime = '2026-11-19T08:00:00'
Must-Fail { Sync-DemoFixedEvent $event 'de' $request } 'Changed fixed event'
$script:events = @($script:events[0], $script:events[0])
Must-Fail { Sync-DemoFixedEvent $event 'de' $request } 'Ambiguous'
$script:events = @(); $script:failPost = $true
Must-Fail { Sync-DemoFixedEvent $event 'de' $request } 'injected POST failure'
$script:failPost = $false; $script:posts = 0; $script:pagination = $false

# Import the actual generic legacy helpers without executing sign-in/setup.
$tenant = Get-Content (Join-Path $root 'tenant.ps1') -Raw -Encoding UTF8
$ast = [Management.Automation.Language.Parser]::ParseInput($tenant, [ref]$null, [ref]$null)
foreach ($name in @('Next-Weekday', 'L', 'To-AsciiJson', 'Assert-DemoMailbox', 'Assert-DemoSender', 'Read-Eml', 'Decode-Words', 'Set-DemoOnlineMeeting')) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }.GetNewClosure(), $true)
    . ([scriptblock]::Create($node.Extent.Text))
}
foreach ($upn in $config.requiredUsers) {
    $user = [pscustomobject]@{ userPrincipalName = $upn; mail = $upn; assignedPlans = @(@{ service = 'exchange'; capabilityStatus = 'Enabled' }) }
    Assert-DemoMailbox $user $upn ''
    $user.assignedPlans = @()
    Must-Fail { Assert-DemoMailbox $user $upn '' } 'Exchange Online'
}
Must-Fail { Assert-DemoSender ([pscustomobject]@{ address = 'external@example.invalid'; created = $false }) 'Teresa Sac' $config.requiredUsers[0] } 'No demo mail'
foreach ($language in 'en', 'de') {
    foreach ($rel in $config.mails) {
        $eml = Read-Eml ([IO.File]::ReadAllBytes((Join-Path $root ("cowork-assistant-day\" + $rel.Replace('{lang}', $language).Replace('/', '\')))))
        if ($eml.FromAddress -notin $config.requiredUsers -or $eml.Subject -notlike '[[]EAD26 *') { throw 'Sender or subject boundary failed.' }
    }
}
$Language = 'en'
$legacy = [pscustomobject]@{ subject = 'Legacy rolling appointment'; weekday = 'Wednesday'; start = '09:15'; minutes = 30; attendees = @(); body = 'Legacy body' }
$configs = @([pscustomobject]@{ Config = [pscustomobject]@{ events = @($legacy) } })
$senders = @{}; $tz = 'UTC'; $WhatIf = $false
$script:legacyWrites = 0; $script:legacyExists = $false
function T($en, $de) { $en }
function Say($text, $color) { }
function Row($tag, $text, $color) { }
function Count($area, $tag) { }
$section = $tenant.Substring($tenant.IndexOf('$tz = [TimeZoneInfo]::Local.Id'))
$section = $section.Substring(0, $section.IndexOf('# ---------- 5)'))
# The production line selects the system timezone; preserve it in this test too.
$tz = [TimeZoneInfo]::Local.Id
$section = $section -replace '\$tz = \[TimeZoneInfo\]::Local.Id', ''
function G($Method, $Path, $Body, $ContentType, $Headers) {
    if ($Method -eq 'GET' -and $Path -like '/me/calendarView?*') { return @{ value = $(if ($script:legacyExists) { @(@{ subject = $legacy.subject }) } else { @() }) } }
    if ($Method -eq 'POST' -and $Path -eq '/me/events') {
        $payload = $Body | ConvertFrom-Json
        if ($payload.start.dateTime -cne ((Next-Weekday 'Wednesday').ToString('yyyy-MM-dd') + 'T09:15:00') -or $payload.start.timeZone -cne $tz) { throw 'Legacy calendar behavior changed.' }
        $script:legacyWrites++; return @{}
    }
    throw "Unexpected production-loop mock $Method $Path"
}
. ([scriptblock]::Create($section))
$script:legacyExists = $true
. ([scriptblock]::Create($section))
if ($script:legacyWrites -ne 1) { throw 'Legacy repeat behavior changed.' }

# Execute the real fixed-event branch, including its JSON adapter and header.
$configs = @([pscustomobject]@{ Config = $config })
$script:productionWrites = 0
$script:productionReads = 0
function G($Method, $Path, $Body, $ContentType, $Headers) {
    if ($Method -eq 'GET' -and $Path -like '/me/events?*') {
        if ($Headers.Prefer -cne 'outlook.body-content-type="text"') { throw 'Text-body preference was lost.' }
        $script:productionReads++
        return @{ value = @() }
    }
    if ($Method -eq 'POST' -and $Path -eq '/me/events') {
        $payload = $Body | ConvertFrom-Json
        if ($payload.attendees -or -not $payload.singleValueExtendedProperties -or -not $payload.transactionId) { throw 'Fixed branch changed the safety contract.' }
        $script:productionWrites++; return @{}
    }
    throw "Unexpected fixed-loop mock $Method $Path"
}
. ([scriptblock]::Create($section))
if ($script:productionWrites -ne 9 -or $script:productionReads -ne 9) { throw 'Production fixed-event branch failed.' }
$WhatIf = $true
. ([scriptblock]::Create($section))
if ($script:productionWrites -ne 9 -or $script:productionReads -ne 18) { throw 'Production fixed-event WhatIf failed.' }
$WhatIf = $false

# Exercise the actual uploader's preservation branch with an edited cloud file.
$itConfig = $config
$configs = @([pscustomobject]@{ Item = @{ id = 'cowork-assistant-day' }; Config = $itConfig })
$Items = @(@{ id = 'cowork-assistant-day'; dir = (Join-Path $root 'cowork-assistant-day'); drivePath = 'Demo-Kit/Cowork/cowork-assistant-day'; files = @('runs/de/01/Action_Register.xlsx') })
$Links = @{ demos = @{ 'cowork-assistant-day' = @{} } }
$firstFolder = $null
function Seg($p) { $p }
function GOpt($path) { @{ file = @{ hashes = @{ quickXorHash = 'edited-cloud-version' } }; webUrl = 'https://fixture.sharepoint.com/preserved' } }
function Get-QuickXor($path) { throw 'Preserved input unexpectedly reached upload comparison.' }
function G { throw 'Preserved input unexpectedly caused a Graph write.' }
$upload = $tenant.Substring($tenant.IndexOf('# ---------- 2) files'))
$start = $upload.IndexOf('foreach ($it in $Items)')
$upload = $upload.Substring($start, $upload.IndexOf("if (`$report['OneDrive']") - $start)
. ([scriptblock]::Create($upload))
if (@($Links.demos['cowork-assistant-day'].files).Count -ne 1) { throw 'Preserved input lost its file link.' }

# The complete real reset executes against local byte fixtures and strict fakes.
function Import-Module { param($Name, $ErrorAction) }
function Get-MgContext { $context }
$resetState = @{ writes = (New-Object 'Collections.Generic.List[object]'); badBaseline = $false; badContext = $false }
function Invoke-MgGraphRequest {
    param($Method, $Uri, $OutputType, $Body, $ContentType, $InputFilePath)
    $path = $Uri -replace '^https://graph.microsoft.com/v1\.0', ''
    if ($Method -eq 'GET' -and $path -eq '/me?$select=userPrincipalName,mail') { return $me }
    if ($Method -eq 'GET' -and $path -like '*/baseline/de/Action_Register.xlsx') { return @{ file = @{}; '@microsoft.graph.downloadUrl' = 'https://fixture.invalid/workbook' } }
    if ($Method -eq 'GET' -and $path -like '*/baseline/de/Run_Context.md') { return @{ file = @{}; '@microsoft.graph.downloadUrl' = 'https://fixture.invalid/context' } }
    if ($Method -eq 'GET' -and $path -like '*/runs/de') { return @{ id = 'run-parent'; folder = @{ childCount = 3 } } }
    if ($Method -eq 'GET') { throw "Unexpected reset GET $path" }
    if ($Method -notin 'POST', 'PUT' -or $path -notlike '/me/drive/items/*') { throw "Forbidden reset write $Method $path" }
    $bytes = if ($InputFilePath) { [IO.File]::ReadAllBytes($InputFilePath) } else { $null }
    $resetState.writes.Add(@{ method = $Method; path = $path; body = $Body; bytes = $bytes })
    return @{ id = "new-$($resetState.writes.Count)"; webUrl = 'https://fixture.sharepoint.com/result' }
}
function Invoke-WebRequest {
    param($Uri, $OutFile, [switch]$UseBasicParsing)
    if ($Uri -eq 'https://fixture.invalid/workbook') {
        if ($resetState.badBaseline) { [IO.File]::WriteAllText($OutFile, 'changed'); return }
        Copy-Item -LiteralPath (Join-Path $root 'cowork-assistant-day\baseline\de\Action_Register.xlsx') -Destination $OutFile -WhatIf:$false
    } elseif ($Uri -eq 'https://fixture.invalid/context') {
        if ($resetState.badContext) { [IO.File]::WriteAllText($OutFile, 'changed'); return }
        Copy-Item -LiteralPath (Join-Path $root 'cowork-assistant-day\baseline\de\Run_Context.md') -Destination $OutFile -WhatIf:$false
    } else { throw "Unexpected web request $Uri" }
}
$reset = Join-Path $PSScriptRoot 'Reset-CoworkAssistantDay.ps1'
& $reset -Language de -WhatIf
if ($resetState.writes.Count) { throw 'Reset WhatIf wrote data.' }
& $reset -Language de
& $reset -Language de
if ($resetState.writes.Count -ne 8) { throw 'Reset did not create exactly two isolated runs (four writes each).' }
$firstFolder = $resetState.writes[0].body | ConvertFrom-Json
$secondFolder = $resetState.writes[4].body | ConvertFrom-Json
if ($firstFolder.name -ceq $secondFolder.name -or $firstFolder.'@microsoft.graph.conflictBehavior' -cne 'fail') { throw 'Reset reused a namespace.' }
$newContext = [Text.Encoding]::UTF8.GetString($resetState.writes[2].bytes)
if (-not $newContext.Contains("runs/de/$($firstFolder.name)") -or $newContext.Contains('runs/de/01')) { throw 'Fresh context still references the previous run.' }
$before = $resetState.writes.Count
$resetState.badBaseline = $true
Must-Fail { & $reset -Language de } 'Baseline changed'
$resetState.badBaseline = $false; $resetState.badContext = $true
Must-Fail { & $reset -Language de } 'Baseline context changed'
if ($resetState.writes.Count -ne $before) { throw 'Changed baseline/context caused a write.' }
Write-Host 'PASS: offline identity, sender, EML, fixed dates/DST/UTC, repeat/pagination, changed/ambiguous events, failure surfacing, legacy calendar and non-destructive reset tests. No tenant requests.'
