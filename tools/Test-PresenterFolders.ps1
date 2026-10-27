$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WorkIqShared.ps1')
. (Join-Path $PSScriptRoot 'PresenterSetup.ps1')
function Assert($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Must-Fail([scriptblock]$Action, [string]$Pattern) {
    try { & $Action | Out-Null } catch { Assert ("$_" -match $Pattern) "Wrong failure: $_"; return }
    throw "Expected failure: $Pattern"
}
$presenter = @{ id = 'presenter'; mail = 'presenter@contoso.invalid' }
$script:requests = New-Object Collections.ArrayList
$script:folders = @{}
$script:permissions = @{}
$script:uploads = @{}
$script:grantMode = ''
$script:denyPath = ''
function Invoke-WiqGraph {
    param([string]$Method, [string]$Path, $Body, [switch]$AllowNotFound, [string]$InputFile)
    [void]$script:requests.Add(@{ method = $Method; path = $Path; body = $Body })
    if ($script:denyPath -and $Path -match $script:denyPath) { throw 'HTTP 403 denied' }
    if ($Path -match '^/me/drive/items/([^/]+)/permissions$') {
        return @{ value = @($script:permissions[$Matches[1]] | Where-Object { $_ }) }
    }
    if ($Path -match '^/me/drive/items/([^/]+)/invite$') {
        $id = $Matches[1]
        if ($script:grantMode -eq 'partial') { return @{ value = @(@{ error = @{ code = 'denied' } }) } }
        $recipient = if ($script:grantMode -eq 'wrong-user') { 'someone-else' } else { $Body.recipients[0].objectId }
        $grant = @{ id = 'grant'; roles = $Body.roles; grantedToV2 = @{ user = @{ id = $recipient } } }
        if ($script:grantMode -ne 'unconfirmed') { $script:permissions[$id] = @($script:permissions[$id] | Where-Object { $_ }) + @($grant) }
        return @{ value = @($grant) }
    }
    if ($Method -eq 'POST' -and $Path -match '^(.*)/children$') {
        $key = "$($Matches[1]):/$([uri]::EscapeDataString($Body.name))"
        Assert (-not $script:folders.ContainsKey($key)) 'Existing folders must not be recreated.'
        Assert ($Body.'@microsoft.graph.conflictBehavior' -eq 'fail') 'No renamed duplicate folders.'
        $folder = @{ id = "folder-$($script:folders.Count)"; name = $Body.name; folder = [pscustomobject]@{}
            webUrl = "https://contoso.invalid/folder-$($script:folders.Count)" }
        $script:folders[$key] = $folder
        return $folder
    }
    if ($Method -eq 'PUT' -and $InputFile -and $Path -match '^/me/drive/root:/(.*):/content$') {
        $key = $Matches[1]
        $file = @{ webUrl = "https://contoso.invalid/$key"; file = @{ hashes = @{ quickXorHash = (Get-QuickXor $InputFile) } } }
        $script:uploads[$key] = $file
        return $file
    }
    if ($Method -eq 'GET') {
        if ($script:folders.ContainsKey($Path)) { return $script:folders[$Path] }
        if ($Path -match '^/me/drive/root:/(.*)\?') {
            if ($script:uploads.ContainsKey($Matches[1])) { return $script:uploads[$Matches[1]] }
        }
        if ($AllowNotFound) { return $null }
    }
    throw "Unexpected Graph request $Method $Path"
}
function GOpt($Path) { Invoke-WiqGraph GET $Path -AllowNotFound }
function Invoke-MgGraphRequest {
    param($Method, $Uri, $InputFilePath, $ContentType)
    Assert ($Uri.StartsWith('https://graph.microsoft.com/v1.0/')) 'Only the existing Graph endpoint is permitted.'
    Invoke-WiqGraph $Method $Uri.Substring('https://graph.microsoft.com/v1.0'.Length) -InputFile $InputFilePath
}
function Get-QuickXor($Path) { [string](Get-Item -LiteralPath $Path).Length }
function Row($Tag, $Text, $Color) {}
function Say($Text, $Color) {}

$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\tenant.ps1'), [ref]$tokens, [ref]$errors)
Assert (-not $errors.Count) ($errors | Out-String)
foreach ($name in 'Count', 'Seg', 'T') {
    $fn = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    . ([scriptblock]::Create($fn.Extent.Text))
}
$guard = $ast.Find({ param($n) $n -is [Management.Automation.Language.IfStatementAst] -and
    $n.Extent.Text.StartsWith('if ($centralSetup)') -and $n.Extent.Text -match '\$centralBlocked = @' }, $true)
$upload = $ast.Find({ param($n) $n -is [Management.Automation.Language.ForEachStatementAst] -and
    $n.Extent.Text -match '\$lk.files = New-Object' }, $true)
$publish = $ast.Find({ param($n) $n -is [Management.Automation.Language.IfStatementAst] -and
    $n.Extent.Text.StartsWith('if ($centralSetup)') -and $n.Extent.Text -match '\$tag = Publish-DemoPresenterRoot' }, $true)
$paths = $ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Set-DemoPresenterPaths' }, $true)
Assert ($guard -and $upload -and $publish -and $paths -and
    $guard.Extent.EndOffset -lt $paths.Extent.StartOffset -and
    $paths.Extent.EndOffset -lt $upload.Extent.StartOffset -and
    $upload.Extent.EndOffset -lt $publish.Extent.StartOffset) 'Privacy filtering must precede path selection, uploads and root sharing.'
function Invoke-UploadFixture([switch]$WhatIf, [switch]$MissingFile, [switch]$MissingDirectory) {
    $centralSetup = $true; $de = $false; $Language = 'en'; $report = [ordered]@{}
    $Items = @(
        @{ id = 'demo'; dir = $PSScriptRoot; drivePath = 'Demo-Kit/App/demo'; files = @('PresenterSetup.ps1') },
        @{ id = 'nested'; dir = (Split-Path $PSScriptRoot); drivePath = 'Demo-Kit/App/nested'; files = @('tools/PresenterSetup.ps1') },
        @{ id = 'private'; dir = $PSScriptRoot; drivePath = 'Demo-Kit/Private/private'; files = @('PresenterSetup.ps1') },
        @{ id = 'wiq'; dir = $PSScriptRoot; drivePath = 'Demo-Kit/Work-IQ/wiq'; files = @('PresenterSetup.ps1') }
    )
    $configs = @(
        @{ Item = $Items[0]; Config = @{} }, @{ Item = $Items[1]; Config = @{} },
        @{ Item = $Items[2]; Config = @{ presenterAccount = 'owner@contoso.invalid' } },
        @{ Item = $Items[3]; Config = @{ workIqShared = @{ scenario = 'WIQ26' } } }
    )
    $Links = @{ centralSetup = $true; manual = @(); demos = [ordered]@{ demo = @{}; nested = @{}; private = @{}; wiq = @{} } }
    $mails = @(); $firstFolder = $null; $folderCache = @{}
    . ([scriptblock]::Create($guard.Extent.Text))
    $presenterRootPath = Set-DemoPresenterPaths $Items $configs $presenter
    if ($MissingFile) { $Items[0].files = @('not-a-real-demo-file.txt') }
    if ($MissingDirectory) { $Items[0].dir = Join-Path $PSScriptRoot 'not-a-real-demo-folder' }
    . ([scriptblock]::Create($upload.Extent.Text))
    . ([scriptblock]::Create($publish.Extent.Text))
    @{ Links = $Links; Items = $Items; Report = $report; Cache = $folderCache; Path = $presenterRootPath }
}

$preview = Invoke-UploadFixture -WhatIf
Assert (@($script:requests | Where-Object { $_.method -ne 'GET' }).Count -eq 0) 'WhatIf must not create folders, upload or invite.'
Assert (-not $preview.Links.root -and -not $preview.Links.demos.demo.presenterShared -and
    $preview.Report.Sharing.pending -eq 1) 'Preview must not claim live access.'
$first = Invoke-UploadFixture
$rootId = $first.Cache[$first.Path].id
$invites = @($script:requests | Where-Object { $_.path -like '*/invite' })
Assert ($invites.Count -eq 1 -and $invites[0].path -eq "/me/drive/items/$rootId/invite") 'Share only the inner Demo-Kit root, not every child or the admin root.'
Assert ($invites[0].body.sendInvitation -and $invites[0].body.requireSignIn -and
    $invites[0].body.roles[0] -eq 'write' -and
    $invites[0].body.recipients[0].objectId -eq $presenter.id) 'Send one native, signed-in, selected-user editing invitation.'
Assert ($script:uploads.Count -eq 2 -and @($script:uploads.Keys | Where-Object {
    $_ -notlike 'DemoKit-Presenters/presenter/Demo-Kit/App/*' }).Count -eq 0) 'Upload both manifest files, including nested folders, only into the eligible recipient tree.'
Assert ($first.Links.root -eq $first.Cache[$first.Path].webUrl -and
    $first.Links.demos.demo.presenterShared -and $first.Links.demos.nested.presenterShared -and
    -not $first.Links.demos.Contains('private') -and -not $first.Links.demos.wiq.presenterShared -and
    ($first.Items | Where-Object id -eq 'wiq').drivePath -ceq 'Demo-Kit/Work-IQ/wiq') 'Publish inherited demo links, exclude private demos, and leave WIQ26 unchanged.'
Assert (($first.Links.manual -join ' ') -match 'Add shortcut to My files') 'The user must get the required one-time shortcut instruction.'
$writes = @($script:requests | Where-Object { $_.method -ne 'GET' }).Count
$again = Invoke-UploadFixture
Assert (@($script:requests | Where-Object { $_.method -ne 'GET' }).Count -eq $writes -and
    $again.Report.OneDrive.same -eq 2 -and $again.Report.Sharing.same -eq 1) 'Rerun must not re-upload, recreate folders or send another sharing invitation.'
[void](Invoke-UploadFixture -WhatIf)
Assert (@($script:requests | Where-Object { $_.method -ne 'GET' }).Count -eq $writes) 'WhatIf against an existing tree must remain read-only.'
Write-Host 'Real upload loop: safe single root, nested manifest files, private/WIQ26 isolation, inherited links, preview and idempotent rerun.'

$secondUser = @{ id = 'second'; mail = 'second@contoso.invalid' }
$secondItems = @(@{ id = 'demo'; drivePath = 'Demo-Kit/App/demo' })
$secondPath = Set-DemoPresenterPaths $secondItems @() $secondUser
$secondFolder = Get-DemoDriveFolder $secondPath @{}
$secondLinks = @{ demos = @{ demo = @{ web = 'https://contoso.invalid/second-demo' } }; manual = @() }
Assert ((Publish-DemoPresenterRoot $secondPath $secondItems $secondLinks $secondUser @{} -Language de) -eq 'new') 'The second recipient needs their own root grant.'
Assert ($secondFolder.id -ne $rootId -and
    $script:permissions[$secondFolder.id][0].grantedToV2.user.id -eq 'second' -and
    $script:permissions[$rootId][0].grantedToV2.user.id -eq 'presenter' -and
    ($secondLinks.manual -join ' ') -match 'Als second@contoso.invalid.*Einmal') 'Recipients and localized shortcut instructions must stay isolated.'
Must-Fail { Publish-DemoPresenterRoot 'Demo-Kit' $secondItems @{} $secondUser @{} } 'Only the selected'
Must-Fail { Set-DemoPresenterPaths @(@{ id = 'demo'; drivePath = 'Demo-Kit/../demo' }) @() $presenter } 'Unsafe'
Must-Fail { Set-DemoPresenterPaths @(@{ id = 'demo'; drivePath = 'Demo-Kit/App/other' }) @() $presenter } 'Unsafe'
Must-Fail { Set-DemoPresenterPaths @(@{ id = 'demo'; drivePath = $secondItems[0].drivePath }) @() $presenter } 'Unexpected'
Must-Fail { Set-DemoPresenterPaths @(@{ id = 'private'; drivePath = 'Demo-Kit/App/private' }) @(
    @{ Item = @{ id = 'private' }; Config = @{ executiveEvent = @{ scenario = 'private' } } }
) $presenter } 'Private/account-locked'
Write-Host 'Recipient separation: different roots, exact write principals, bilingual user step and unsafe/private path rejection.'

foreach ($mode in 'partial', 'wrong-user', 'unconfirmed') {
    $script:permissions.Remove($rootId)
    $script:grantMode = $mode
    $failed = Invoke-UploadFixture
    Assert (-not $failed.Links.root -and -not $failed.Links.demos.demo.presenterShared -and
        $failed.Report.Sharing.failed -eq 1) "Failed $mode grant must not publish readiness."
}
$script:grantMode = ''
$script:denyPath = ':/DemoKit-Presenters$'
$denied = Invoke-UploadFixture
Assert (-not $denied.Links.root -and $denied.Report.OneDrive.failed -eq 2) '403 must not be treated as folder-not-found.'
$script:denyPath = ''
$missing = Invoke-UploadFixture -MissingFile
Assert (-not $missing.Links.root -and $missing.Report.OneDrive.failed -eq 1) 'Missing manifest input must prevent a success-shaped start link.'
$missing = Invoke-UploadFixture -MissingDirectory
Assert (-not $missing.Links.root -and $missing.Report.OneDrive.failed -eq 1) 'Missing cache folder must prevent a success-shaped start link.'
$script:folders['/me/drive/root:/broken'] = @{ id = 'file'; file = @{} }
Must-Fail { Get-DemoDriveFolder 'broken' @{} } 'not a folder'
Must-Fail { Get-DemoDriveFolder 'broken' @{} -WhatIf } 'could not be confirmed'
Write-Host 'Failures: denied reads, partial/wrong/unconfirmed grants and missing local inputs never produce shared-ready links.'

$summary = $ast.Find({ param($n) $n -is [Management.Automation.Language.IfStatementAst] -and
    $n.Extent.Text.StartsWith('if ($centralSetup)') -and $n.Extent.Text -match '\$rows = @' }, $true)
$fn = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'L' }, $true)
. ([scriptblock]::Create($fn.Extent.Text))
function Send-DemoPresenterMessage($Payload, $Operator, $Recipient, $Language, $Revision, [switch]$WhatIf) {
    $script:lastMail = $Payload
    'new'
}
$centralSetup = $true; $WhatIf = $false; $de = $false; $report = [ordered]@{}
$me = @{ displayName = 'Publisher'; mail = 'admin@contoso.invalid' }; $myMail = $me.mail
$Items = $first.Items; $Links = $first.Links
foreach ($Language in 'en', 'de') {
    . ([scriptblock]::Create($summary.Extent.Text))
    Assert ($script:lastMail.message.toRecipients[0].emailAddress.address -eq $presenter.mail -and
        $script:lastMail.message.body.content.Contains("Demo-Kit: $($first.Links.root)") -and
        $script:lastMail.message.body.content.Contains('Add shortcut to My files')) 'Actual localized summary must send the verified start link and one-time user step to the chosen user.'
}
$Links = $failed.Links
. ([scriptblock]::Create($summary.Extent.Text))
Assert (-not $script:lastMail.message.body.content.Contains($first.Links.root) -and
    $script:lastMail.message.body.content -notmatch 'Add shortcut to My files') 'Failed root sharing must not email a working link or pretend the user can add its shortcut.'
Write-Host 'Actual summary email: selected recipient, verified root/user step in both languages and honest blocked output.'

$install = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\install.ps1'), [ref]$tokens, [ref]$errors)
Assert (-not $errors.Count) ($errors | Out-String)
foreach ($name in 'FileUri', 'Get-LinkGroups') {
    $fn = $install.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    . ([scriptblock]::Create($fn.Extent.Text))
}
$kit = 'Demo-Kit'; $root = Join-Path $PSScriptRoot 'unused-local-cache'
$links = @{ centralSetup = $true; demos = [ordered]@{
    blocked = @{ app = 'App'; title = 'Blocked'; dir = $PSScriptRoot; web = 'https://contoso.invalid/blocked-folder'
        files = @(@{ name = 'blocked-file'; url = 'https://contoso.invalid/blocked-file' }); local = @('blocked-local') }
    allowed = @{ app = 'App'; title = 'Allowed'; presenterShared = $true; web = 'https://contoso.invalid/allowed-folder'
        files = @(@{ name = 'allowed-file'; url = 'https://contoso.invalid/allowed-file' }) }
    wiq = @{ app = 'Work IQ'; title = 'WIQ26'; dir = $PSScriptRoot; local = @()
        extra = @(@{ name = 'WIQ26'; url = 'https://contoso.invalid/shared-wiq' }) }
} }
$groups = Get-LinkGroups | ConvertTo-Json -Depth 20
Assert ($groups -notmatch 'https://contoso.invalid/blocked-' -and $groups -notmatch 'blocked-local' -and
    $groups -match 'https://contoso.invalid/allowed-file' -and $groups -match 'https://contoso.invalid/shared-wiq') 'Desktop/Edge must exclude unconfirmed central cloud/local links without hiding WIQ26.'
$links.centralSetup = $false
Assert ((Get-LinkGroups | ConvertTo-Json -Depth 20) -match 'https://contoso.invalid/blocked-file') 'Self setup must retain its existing links.'
Assert (@($script:requests | Where-Object { $_.path -match '^/users/.*/(drive|mail|events)' -or
    $_.method -in 'DELETE', 'PATCH' }).Count -eq 0) 'No recipient impersonation, legacy deletion or existing grant revocation.'
Write-Host 'Link publishing: verified start/child links only, unchanged self setup, and no impersonation/destructive migration.'
