$ErrorActionPreference = 'Stop'
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\tenant.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
foreach ($name in 'Read-DemoSenderCache', 'Set-DemoOnlineMeeting') {
    $fn = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if (-not $fn) { throw "$name not found." }
    . ([scriptblock]::Create($fn.Extent.Text))
}
function Must-Fail([scriptblock]$Action) {
    $blocked = $false
    try { & $Action | Out-Null } catch { $blocked = $true }
    if (-not $blocked) { throw 'Expected unsafe input to be rejected.' }
}
$tenant = '11111111-1111-1111-1111-111111111111'
$presenter = 'presenter@contoso.invalid'
$entry = [pscustomobject]@{
    version = 3; tenantId = $tenant; presenter = $presenter
    senders = [pscustomobject]@{ 'Demo Sender' = [pscustomobject]@{ address = 'sender@contoso.invalid' } }
}
$cache = Read-DemoSenderCache ($entry | ConvertTo-Json -Depth 5) $tenant $presenter.ToUpperInvariant()
if ($cache.Count -ne 1 -or $cache['Demo Sender'].address -ne 'sender@contoso.invalid') { throw 'Valid cache was not preserved.' }
Must-Fail { Read-DemoSenderCache $entry $tenant 'another@contoso.invalid' }
Must-Fail { Read-DemoSenderCache $entry 'another-tenant' $presenter }
Must-Fail { Read-DemoSenderCache $entry '' $presenter }
Must-Fail { Read-DemoSenderCache $entry $tenant '' }
Must-Fail { Read-DemoSenderCache '{"Demo Sender":{"address":"sender@contoso.invalid"}}' $tenant $presenter }
Must-Fail { Read-DemoSenderCache '{"version":2}' $tenant $presenter }
Write-Host 'Sender cache: exact tenant/presenter required; legacy and cross-account caches rejected.'

$payload = @{ subject = 'Demo' }
Set-DemoOnlineMeeting ([pscustomobject]@{ onlineMeeting = $true }) $payload
if (-not $payload.isOnlineMeeting -or $payload.onlineMeetingProvider -ne 'teamsForBusiness') { throw 'Teams meeting was not requested.' }
foreach ($event in @([pscustomobject]@{}, [pscustomobject]@{ onlineMeeting = $false })) {
    $payload = @{ subject = 'Demo' }
    Set-DemoOnlineMeeting $event $payload
    if ($payload.Count -ne 1) { throw 'Ordinary event behavior changed.' }
}
Must-Fail { Set-DemoOnlineMeeting ([pscustomobject]@{ onlineMeeting = 'true' }) @{} }
Write-Host 'Calendar: explicit true creates Teams payload; ordinary appointments unchanged; strings rejected.'

$child = $ast.Find({
    param($node)
    $node -is [Management.Automation.Language.StringConstantExpressionAst] -and
        $node.Value -match 'Connect-ExchangeOnline' -and $node.Value -match 'foreach \(\$n in \$cfg.names\)'
}, $true)
if (-not $child) { throw 'Exchange child not found.' }
$childAst = [Management.Automation.Language.Parser]::ParseInput($child.Value, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$connectionCheck = $childAst.Find({
    param($node)
    $node -is [Management.Automation.Language.IfStatementAst] -and $node.Extent.Text -match 'connections.Count -ne 1'
}, $true)
if (-not $connectionCheck) { throw 'Exchange tenant guard missing.' }
function Disconnect-ExchangeOnline { param([switch]$Confirm) }
$cfg = [pscustomobject]@{ tenantId = $tenant }
$connections = @([pscustomobject]@{ TenantID = $tenant })
. ([scriptblock]::Create($connectionCheck.Extent.Text))
foreach ($value in @('another-tenant', '')) {
    $connections = @([pscustomobject]@{ TenantID = $value })
    Must-Fail { . ([scriptblock]::Create($connectionCheck.Extent.Text)) }
}
$connections = @()
Must-Fail { . ([scriptblock]::Create($connectionCheck.Extent.Text)) }
if ($child.Value -match 'UserPrincipalName\s*=\s*\$cfg.upn') { throw 'Exchange admin must not be forced to use the non-admin presenter account.' }
if ($child.Value -notmatch 'Trustee \$cfg.upn') { throw 'Send As must target the presenter, not the Exchange admin.' }
Write-Host 'Exchange: separate admin choice, exact tenant guard, presenter remains the Send As trustee.'

$guardLoop = $ast.Find({
    param($node)
    $node -is [Management.Automation.Language.ForEachStatementAst] -and
        $node.Extent.Text -match 'Assert-DemoPresenter \$c.Config.presenterAccount' -and
        $node.Extent.Text -match 'presenterBlocked'
}, $true)
if (-not $guardLoop) { throw 'Account-locked scenario filtering missing.' }
. (Join-Path $PSScriptRoot 'DemoCalendar.ps1')
function Get-MgContext { [pscustomobject]@{ Account = $presenter; TenantId = $tenant } }
function Row { }
$me = [pscustomobject]@{ userPrincipalName = $presenter; mail = $presenter }
$Links = @{}
$presenterBlocked = @()
$configs = @(
    [pscustomobject]@{ Item = @{ id = 'locked' }; Config = @{ presenterAccount = 'owner@contoso.invalid' } },
    [pscustomobject]@{ Item = @{ id = 'own' }; Config = @{ presenterAccount = $presenter } }
)
. ([scriptblock]::Create($guardLoop.Extent.Text))
if ($presenterBlocked.Count -ne 1 -or $presenterBlocked[0] -ne 'locked' -or $Links.manual -notmatch 'BLOCKED') {
    throw 'An incompatible account must be explicitly blocked without stopping compatible scenarios.'
}
$Items = @(@{ id = 'locked' }, @{ id = 'own' })
$mails = @(@{ Demo = 'locked' }, @{ Demo = 'own' })
foreach ($variable in 'configs', 'Items', 'mails') {
    $filter = $ast.Find({
        param($node)
        $node -is [Management.Automation.Language.AssignmentStatementAst] -and
        $node.Left.Extent.Text -eq ('$' + $variable) -and $node.Extent.Text -match 'presenterBlocked'
    }, $true)
    if (-not $filter) { throw "Missing blocked-scenario filter: $variable" }
    . ([scriptblock]::Create($filter.Extent.Text))
}
if ($configs.Count -ne 1 -or $configs[0].Item.id -ne 'own' -or
    $Items.Count -ne 1 -or $Items[0].id -ne 'own' -or
    $mails.Count -ne 1 -or $mails[0].Demo -ne 'own') {
    throw 'Blocked scenario content must be removed before any writes.'
}
Write-Host 'Scenario isolation: incompatible presenter reported; matching presenter accepted.'
