$ErrorActionPreference = 'Stop'
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\tenant.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
foreach ($name in @('Assert-DemoMailbox', 'Assert-DemoSender')) {
    $fn = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if (-not $fn) { throw "Guard $name not found." }
    . ([scriptblock]::Create($fn.Extent.Text))
}

$upn = 'TeresaS@M365CPI98544940.OnMicrosoft.com'
function User($plans, $name = 'Teresa Sac', $mail = $upn) {
    [pscustomobject]@{ userPrincipalName = $upn; displayName = $name; mail = $mail; assignedPlans = @($plans) }
}
function Must-Fail($user, $name, $message) {
    try { Assert-DemoMailbox $user $upn $name }
    catch {
        if ("$_" -notlike "*$message*") { throw "Wrong failure: $_" }
        return
    }
    throw "Expected failure: $message"
}

Assert-DemoMailbox (User @([pscustomobject]@{ service = 'exchange'; capabilityStatus = 'Enabled' })) $upn 'Teresa Sac'
Must-Fail (User @([pscustomobject]@{ service = 'Teams'; capabilityStatus = 'Enabled' })) 'Teresa Sac' 'no enabled Exchange'
Must-Fail (User @([pscustomobject]@{ service = 'exchange'; capabilityStatus = 'Suspended' })) 'Teresa Sac' 'no enabled Exchange'
Must-Fail (User @()) 'Teresa Sac' 'no enabled Exchange'
Must-Fail (User @() 'Different person') 'Teresa Sac' 'not Teresa Sac'
Must-Fail $null 'Teresa Sac' 'not resolved exactly'
Must-Fail (User @([pscustomobject]@{ service = 'exchange'; capabilityStatus = 'Enabled' }) 'Teresa Sac' 'wrong@example.invalid') 'Teresa Sac' 'mailbox address'
Write-Host 'Mailbox guard: enabled Exchange accepted; Teams-only, suspended, absent, wrong-person and wrong-address cases blocked.'

Assert-DemoSender ([pscustomobject]@{ address = $upn; created = $false }) 'Teresa Sac' $upn
Assert-DemoSender ([pscustomobject]@{ address = $upn.ToLowerInvariant() }) 'Teresa Sac' $upn
foreach ($sender in @($null, [pscustomobject]@{ address = 'wrong@example.invalid' },
        [pscustomobject]@{ address = $upn; created = $true })) {
    $blocked = $false
    try { Assert-DemoSender $sender 'Teresa Sac' $upn }
    catch {
        if ("$_" -notlike '*must use the existing mailbox*') { throw }
        $blocked = $true
    }
    if (-not $blocked) { throw 'An unresolved, wrong or substitute sender was accepted.' }
}
Write-Host 'Sender guard: correct cached/fresh addresses accepted; unresolved, wrong and substitute senders blocked.'

$exchange = $ast.Find({
    param($node)
    $node -is [Management.Automation.Language.StringConstantExpressionAst] -and
        $node.Value -match 'Connect-ExchangeOnline' -and $node.Value -match 'foreach \(\$n in \$cfg.names\)'
}, $true)
if (-not $exchange) { throw 'Exchange child script not found.' }
$childAst = [Management.Automation.Language.Parser]::ParseInput($exchange.Value, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$loop = $childAst.Find({
    param($node)
    $node -is [Management.Automation.Language.ForEachStatementAst] -and $node.Variable.VariablePath.UserPath -eq 'n'
}, $true)
if (-not $loop) { throw 'Exchange sender loop not found.' }
& {
    $script:grants = 0
    $script:lookups = @()
    function Get-Recipient { param($Identity) $script:lookups += $Identity; $script:recipient }
    function Get-RecipientPermission { }
    function Add-RecipientPermission { $script:grants++ }
    function New-Mailbox { throw 'Verified user must never become a shared mailbox.' }
    $cfg = [pscustomobject]@{
        upn = "admin@$($upn.Split('@')[1])"; names = @('Teresa Sac'); whatIf = $false
        verifiedSenders = [pscustomobject]@{ 'Teresa Sac' = $upn }
    }
    foreach ($case in @(
        @{ address = $upn; name = 'Teresa Sac'; type = 'UserMailbox'; granted = 1 },
        @{ address = 'wrong@example.invalid'; name = 'Teresa Sac'; type = 'UserMailbox'; granted = 0 },
        @{ address = $upn; name = 'Different person'; type = 'UserMailbox'; granted = 0 },
        @{ address = $upn; name = 'Teresa Sac'; type = 'SharedMailbox'; granted = 0 }
    )) {
        $script:recipient = [pscustomobject]@{
            PrimarySmtpAddress = $case.address; DisplayName = $case.name
            RecipientTypeDetails = $case.type; Identity = $upn
        }
        $script:grants = 0; $script:lookups = @(); $res = @()
        . ([scriptblock]::Create($loop.Extent.Text))
        if ($script:grants -ne $case.granted -or @($script:lookups).Count -ne 1 -or $script:lookups[0] -ne $upn) {
            throw 'Exchange must resolve the verified UPN and validate it before granting Send As.'
        }
        if ($case.granted -eq 0 -and -not $res[0].error) { throw 'Mismatched Exchange recipient was not reported.' }
    }
}
Write-Host 'Exchange preflight: verified UPN lookup; wrong address/person/type rejected before Send As grants.'
