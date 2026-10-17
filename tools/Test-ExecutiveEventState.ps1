# Offline only: never authenticate or call a real Microsoft 365 endpoint.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'ExecutiveEvent.ps1')
$config = (Get-Content (Join-Path $root 'cowork-executive-event\tenant.json') -Raw -Encoding UTF8 | ConvertFrom-Json).executiveEvent
$settings = [pscustomobject]@{
    tenantId = '11111111-1111-1111-1111-111111111111'; account = $config.account
    teamId = '22222222-2222-2222-2222-222222222222'; channelId = '19:fixture@thread.tacv2'
}
$context = [pscustomobject]@{ TenantId = $settings.tenantId; Account = $settings.account }
$me = [pscustomobject]@{ userPrincipalName = $settings.account; mail = $settings.account }
$ledger = [pscustomobject]@{
    scenario = 'LEF26'; tenantId = $settings.tenantId; account = $settings.account; language = 'de'
    teamId = $settings.teamId; channelId = $settings.channelId
    driveId = 'drive'; parentId = 'parent'; rootId = 'root'
    files = @([pscustomobject]@{ name = 'Event_Brief.docx'; id = 'original-file'; cTag = 'original-tag' })
    events = @([pscustomobject]@{ key = 'forum'; id = 'original-event' })
    threads = @([pscustomobject]@{ phase = 'BEFORE'; id = 'original-thread' }); mailReplies = @()
}
$script:cloud = [pscustomobject]@{ schemaVersion = 1; language = 'de'; settings = $settings; ledger = $ledger }
$script:reads = 0
$script:writes = 0
$script:denyRead = $false
function Get-MgContext { $context }
function Invoke-RestMethod {
    param($Uri, [switch]$UseBasicParsing)
    if ($Uri -ne 'https://fixture.invalid/private-state') { throw "Unexpected download $Uri" }
    $script:cloud | ConvertTo-Json -Depth 25 -Compress | ConvertFrom-Json
}
function Invoke-ExecutiveEventGraph {
    param($Method, $Path, $Body, [switch]$Optional)
    if ($Method -eq 'GET') {
        $script:reads++
        if ($script:denyRead) { throw 'HTTP 403: private cache access denied' }
        if ($Path -eq '/me/drive/root:/DemoKit/executive-event-state-de-v1.json') {
            if ($script:cloud) { return @{ file = @{ mimeType = 'application/json' }; '@microsoft.graph.downloadUrl' = 'https://fixture.invalid/private-state' } }
            return $null
        }
        if ($Path -eq '/me/drive/root:/DemoKit/executive-event-state-en-v1.json') { return $null }
        if ($Path -eq '/me/drive/root:/DemoKit') { return @{ folder = @{} } }
    }
    if ($Method -eq 'PUT' -and $Path -eq '/me/drive/root:/DemoKit/executive-event-state-de-v1.json:/content') {
        $script:writes++
        $script:cloud = $Body | ConvertTo-Json -Depth 25 -Compress | ConvertFrom-Json
        return @{ id = 'private-cache' }
    }
    throw "Unexpected mock Graph call: $Method $Path"
}
function Must-Fail([scriptblock]$Action, [string]$Expected) {
    try { & $Action }
    catch { if ("$_" -notlike "*$Expected*") { throw "Expected '$Expected', got $_" }; return }
    throw "Expected failure: $Expected"
}
$originalLocal = $env:LOCALAPPDATA
$temp = Join-Path ([IO.Path]::GetTempPath()) ("DemoKit-private-state-test-" + [guid]::NewGuid().ToString('N'))
try {
    $env:LOCALAPPDATA = $temp
    $state = Initialize-ExecutiveEventState $config $me de -WhatIf
    if (Test-Path -LiteralPath $temp) { throw 'WhatIf restored local state.' }
    if ($script:writes -or $state.ledger.events[0].id -cne 'original-event') { throw 'WhatIf lost ownership or changed the tenant.' }
    $english = Initialize-ExecutiveEventState $config $me en -WhatIf
    if ($english.settings.teamId -ne $settings.teamId -or $null -ne $english.ledger -or
        $english.language -cne 'en' -or (Test-Path -LiteralPath $temp)) {
        throw 'Default English setup must reuse the approved destination, never the German ownership ledger.'
    }
    $state = Initialize-ExecutiveEventState $config $me de
    $path = Get-ExecutiveEventLedgerPath $settings de
    $settingsPath = Join-Path $temp 'DemoKit\executive-event-settings.json'
    $restored = Get-ExecutiveEventLedger $settings de
    if ($restored.files[0].cTag -cne 'original-tag' -or $restored.threads[0].id -cne 'original-thread') {
        throw 'Restoration did not retain exact object ownership.'
    }
    $before = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    [void](Initialize-ExecutiveEventState $config $me de)
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $before -or $script:writes) {
        throw 'Repeat initialization changed existing state.'
    }
    $originalTenant = $context.TenantId
    $context.TenantId = '33333333-3333-3333-3333-333333333333'
    Must-Fail { Initialize-ExecutiveEventState $config $me de } 'blocked'
    $context.TenantId = $originalTenant
    $context.Account = 'other@example.invalid'
    $reads = $script:reads
    Must-Fail { Initialize-ExecutiveEventState $config $me de } 'approved demo admin'
    if ($script:reads -ne $reads) { throw 'Wrong account accessed private state.' }
    $context.Account = $settings.account
    $script:cloud.language = 'en'
    Must-Fail { Initialize-ExecutiveEventState $config $me de } 'schema or language'
    $script:cloud.language = 'de'
    $script:cloud.ledger.teamId = '33333333-3333-3333-3333-333333333333'
    Must-Fail { Initialize-ExecutiveEventState $config $me de } 'mismatched ownership'
    $script:cloud.ledger.teamId = $settings.teamId
    $script:denyRead = $true
    Must-Fail { Initialize-ExecutiveEventState $config $me de } '403'
    $script:denyRead = $false
    $restored.rootId = 'conflicting-root'
    [IO.File]::WriteAllText($path, ($restored | ConvertTo-Json -Depth 25))
    Remove-Item -LiteralPath $settingsPath
    Must-Fail { Initialize-ExecutiveEventState $config $me de } 'ledger differs'
    if (Test-Path -LiteralPath $settingsPath) { throw 'Conflicting ledger allowed partial settings restore.' }
    [IO.File]::WriteAllText($path, ($ledger | ConvertTo-Json -Depth 25))
    [void](Initialize-ExecutiveEventState $config $me de)
    $different = $settings | ConvertTo-Json | ConvertFrom-Json
    $different.channelId = '19:other@thread.tacv2'
    [IO.File]::WriteAllText($settingsPath, ($different | ConvertTo-Json))
    Must-Fail { Initialize-ExecutiveEventState $config $me de } 'settings differ'
    [IO.File]::WriteAllText($settingsPath, ($settings | ConvertTo-Json))
    $savedCloud = $script:cloud
    $script:cloud = $null
    $localState = Initialize-ExecutiveEventState $config $me de
    if ($localState.ledger.rootId -cne 'root') { throw 'Original local state could not bootstrap private cache.' }
    $env:LOCALAPPDATA = Join-Path $temp 'empty-pc'
    Must-Fail { Initialize-ExecutiveEventState $config $me de } 'missing locally and in OneDrive'
    if (Test-Path -LiteralPath $env:LOCALAPPDATA) { throw 'Missing approval wrote local files.' }
    $env:LOCALAPPDATA = $temp
    $script:cloud = $savedCloud
    Enable-ExecutiveEventCloudSync $config $settings $me de
    $updated = Get-ExecutiveEventLedger $settings de
    $updated.files[0].cTag = 'new-tag'
    Save-ExecutiveEventLedger $updated $path
    if ($script:writes -ne 2 -or $script:cloud.ledger.files[0].cTag -cne 'new-tag') {
        throw 'Ownership ledger updates were not persisted in private OneDrive.'
    }
    $env:LOCALAPPDATA = Join-Path $temp 'second-pc'
    [void](Initialize-ExecutiveEventState $config $me de)
    if ((Get-ExecutiveEventLedger $settings de).files[0].cTag -cne 'new-tag') {
        throw 'A second PC recovered stale ownership.'
    }
    $tenant = Get-Content (Join-Path $root 'tenant.ps1') -Raw
    if ($tenant.IndexOf('Connect-MgGraph -Scopes') -gt $tenant.IndexOf('$state = Initialize-ExecutiveEventState') -or
        ([regex]::Matches($tenant, '(?m)^Connect-MgGraph ')).Count -ne 1) {
        throw 'Recovery must follow the single combined Graph sign-in.'
    }
    if ($tenant -notmatch '(?s)if \(@\(\$configs \| Where-Object \{ \$_\.Config\.executiveEvent \}\)\.Count\) \{\s+\$scopes') {
        throw 'New PCs must request the scenario scopes before restoring local settings.'
    }
    $ast = [Management.Automation.Language.Parser]::ParseInput($tenant, [ref]$null, [ref]$null)
    $loop = @($ast.EndBlock.Statements | Where-Object {
        $_ -is [Management.Automation.Language.ForEachStatementAst] -and
        $_.Extent.Text.Contains('$state = Initialize-ExecutiveEventState')
    })
    if ($loop.Count -ne 1) { throw 'Cannot locate the actual tenant private-state preflight loop.' }
    $filters = @($ast.EndBlock.Statements | Where-Object {
        $_.Extent.StartOffset -gt $loop[0].Extent.EndOffset -and
        $_ -is [Management.Automation.Language.AssignmentStatementAst]
    } | Select-Object -First 5)
    $production = [scriptblock]::Create($loop[0].Extent.Text + "`n" + (($filters | ForEach-Object { $_.Extent.Text }) -join "`n"))
    function Say($Text, $Color) { Write-Host $Text }
    function Test-ExecutiveEventDestination($Config, $Settings, $Me, $Language, $LedgerOverride) {
        if ($LedgerOverride.files[0].cTag -cne 'new-tag') { throw 'WhatIf did not pass recovered ownership to destination validation.' }
        @{ ledger = $LedgerOverride }
    }
    $WhatIf = $true
    $Language = 'de'
    $env:LOCALAPPDATA = Join-Path $temp 'production-dry-run'
    foreach ($blocked in @($false, $true)) {
        $configs = @(
            [pscustomobject]@{ Item = @{ id = 'event' }; Config = @{ executiveEvent = $config } },
            [pscustomobject]@{ Item = @{ id = 'ordinary' }; Config = @{} }
        )
        $Items = @($configs | ForEach-Object { $_.Item })
        $mails = @([pscustomobject]@{ Demo = 'event' }, [pscustomobject]@{ Demo = 'ordinary' })
        $executiveSettings = @{}
        $executiveDestinations = @{}
        $executiveBlocked = @()
        if ($blocked) { $context.Account = 'other@example.invalid' }
        . $production
        if ($blocked) {
            if ($executiveBlocked.Count -ne 1 -or $configs.Count -ne 1 -or
                $configs[0].Item.id -ne 'ordinary' -or $Items.Count -ne 1 -or
                $mails.Count -ne 1 -or $mails[0].Demo -ne 'ordinary') {
                throw 'Blocked private state did not isolate only its scenario; ordinary demos must continue.'
            }
        } elseif ($executiveDestinations.Count -ne 1 -or $configs.Count -ne 2) {
            throw 'Actual tenant preflight did not retain the recovered and ordinary demos.'
        }
        $context.Account = $settings.account
    }
    if ((Test-Path -LiteralPath $env:LOCALAPPDATA) -or $script:writes -ne 2) {
        throw 'Production WhatIf wrote local or tenant state.'
    }
    Write-Host 'Private setup state checks passed: cross-PC recovery, exact ownership, dry-run, idempotency, identity/schema/language/conflict/403 denial, local bootstrap and automatic ledger backup. No real tenant calls.'
} finally {
    $env:LOCALAPPDATA = $originalLocal
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Recurse -Force }
}
