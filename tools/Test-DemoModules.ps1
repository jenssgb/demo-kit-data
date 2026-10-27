param([switch]$Online)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'DemoModules.ps1')
$tokens = $null; $errors = $null
$tenant = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\tenant.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
foreach ($file in 'DemoModules.ps1', 'Test-DemoModules.ps1') {
    [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $file), [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw ($errors | Out-String) }
}
$text = $tenant.Extent.Text
if ($text.IndexOf("Install-DemoModule 'Microsoft.Graph.Authentication'") -gt $text.IndexOf('Connect-MgGraph -Scopes')) {
    throw 'Module installation must precede Graph sign-in.'
}
if ($text -notmatch 'exchangeModule = \$exchangeModule' -or $text -notmatch 'Import-Module \$cfg.exchangeModule -ErrorAction Stop') {
    throw 'The Exchange child must reuse the module installed by its parent.'
}
$installer = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\install.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
if ($installer.Extent.Text -notmatch '\$raw = "https://raw.githubusercontent.com/\$Repo/\$revision"') {
    throw 'A run must fetch its tenant script and helper from the same immutable revision when available.'
}

$originalModulePath = $env:PSModulePath
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('demo-kit-module-test-' + [guid]::NewGuid().ToString('N'))
$script:moduleRoot = Join-Path $testRoot 'Modules'
function Get-DemoModuleRoot { $script:moduleRoot }
New-Item -ItemType Directory -Force -Path $testRoot | Out-Null
$env:PSModulePath = "$script:moduleRoot;$(Join-Path $PSHOME 'Modules')"
function Install-PackageProvider { param($Name, $MinimumVersion, $Scope, [switch]$Force, $ErrorAction) }
try {
    if ($Online) {
        function Install-Module { throw 'Simulated Gallery API 404: exercise the verified CDN fallback.' }
        $graph = Install-DemoModule 'Microsoft.Graph.Authentication'
        $exchange = Install-DemoModule 'ExchangeOnlineManagement'
        if ((Install-DemoModule 'Microsoft.Graph.Authentication') -ne $graph) { throw 'Rerun did not reuse Graph.' }
        $child = Join-Path $testRoot 'module-import.ps1'
        @'
param($Manifest, $Kind)
$ErrorActionPreference = 'Stop'
Import-Module $Manifest -ErrorAction Stop
if ($Kind -eq 'Graph') {
    Get-Command Connect-MgGraph -ErrorAction Stop | Out-Null
} elseif (-not (Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) {
    throw 'Exchange import is missing DisableWAM.'
}
Write-Host "$Kind imported in a clean child process; no sign-in performed."
'@ | Set-Content -LiteralPath $child -Encoding ASCII
        foreach ($module in @(@{ Kind = 'Graph'; Path = $graph }, @{ Kind = 'Exchange'; Path = $exchange })) {
            & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File $child -Manifest $module.Path -Kind $module.Kind
            if ($LASTEXITCODE -ne 0) { throw "$($module.Kind) child import failed." }
        }
        Write-Host "Online bootstrap passed on $($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion): actual downloads, checksum/manifest validation, imports and rerun."
        return
    }

    $script:name = 'DemoKitBootstrapFixture'
    $source = Join-Path $testRoot 'Source'
    New-Item -ItemType Directory -Path $source | Out-Null
    Set-Content -LiteralPath (Join-Path $source "$script:name.psm1") -Value 'function Get-DemoKitFixture { "fixture" }' -Encoding ASCII
    New-ModuleManifest -Path (Join-Path $source "$script:name.psd1") -RootModule "$script:name.psm1" -ModuleVersion '1.2.3'
    $script:archive = Join-Path $testRoot 'fixture.zip'
    Compress-Archive -Path (Join-Path $source '*') -DestinationPath $script:archive
    $script:package = @{ Version = '1.2.3'; Dependencies = @(); Sha256 = (Get-FileHash $script:archive).Hash }
    function Get-DemoModulePackage { param($Name) $script:package }
    $script:managerCalls = 0; $script:downloadCalls = 0; $script:managerFails = $false; $script:failDownloads = 0
    function Install-Module {
        [CmdletBinding()]
        param($Name, $RequiredVersion, $Repository, $Scope, [switch]$Force, [switch]$AllowClobber, [switch]$AcceptLicense)
        $script:managerCalls++
        if ($script:managerFails) { Write-Error 'Gallery returned HTTP 404'; return }
        Expand-Archive -LiteralPath $script:archive -DestinationPath (Join-Path (Join-Path $script:moduleRoot $Name) $RequiredVersion)
    }
    function Invoke-WebRequest {
        [CmdletBinding()]
        param($Uri, $OutFile, [switch]$UseBasicParsing, $TimeoutSec)
        $script:downloadCalls++
        if ($script:downloadCalls -le $script:failDownloads) { throw 'Simulated network failure.' }
        Copy-Item -LiteralPath $script:archive -Destination $OutFile
    }
    function Start-Sleep { param($Seconds) }
    function Must-Fail([scriptblock]$Action, [string]$Message) {
        $caught = $null
        try { & $Action | Out-Null } catch { $caught = "$_" }
        if (-not $caught -or $caught -notmatch $Message) { throw "Expected '$Message'; got '$caught'." }
    }
    function Reset-Fixture {
        if (Test-Path -LiteralPath $script:moduleRoot) { Remove-Item -LiteralPath $script:moduleRoot -Recurse -Force }
        $script:managerCalls = 0; $script:downloadCalls = 0; $script:failDownloads = 0
    }
    $manifest = Install-DemoModule $script:name
    if (-not (Test-Path -LiteralPath $manifest) -or $script:managerCalls -ne 1 -or $script:downloadCalls) { throw 'Normal installation failed.' }
    if ((Install-DemoModule $script:name) -ne $manifest -or $script:managerCalls -ne 1) { throw 'Existing module was reinstalled.' }
    Write-Host 'Normal package-manager install and idempotent rerun passed.'

    Reset-Fixture
    $script:managerFails = $true; $script:failDownloads = 2
    $manifest = Install-DemoModule $script:name
    if (-not (Test-Path -LiteralPath $manifest) -or $script:downloadCalls -ne 3) { throw 'Gallery failure/retry fallback failed.' }
    if (@(Get-ChildItem -LiteralPath $script:moduleRoot -Force | Where-Object Name -like '.demo-kit-*').Count) { throw 'Staging files leaked.' }
    Write-Host 'Non-terminating Gallery error falls back; two failed downloads retry; valid package publishes atomically.'

    Reset-Fixture
    $partial = Join-Path (Join-Path $script:moduleRoot $script:name) $script:package.Version
    New-Item -ItemType Directory -Force -Path $partial | Out-Null
    Set-Content -LiteralPath (Join-Path $partial 'incomplete.txt') -Value 'Interrupted installation'
    $manifest = Install-DemoModule $script:name
    if (-not (Test-Path -LiteralPath $manifest) -or (Test-Path -LiteralPath (Join-Path $partial 'incomplete.txt'))) {
        throw 'An interrupted installation was not replaced by a validated package.'
    }
    if (@(Get-ChildItem -LiteralPath $script:moduleRoot -Force | Where-Object Name -like '.demo-kit-*').Count) { throw 'Repair backup leaked.' }
    Write-Host 'Interrupted module folder repaired only after the replacement package is validated.'

    Reset-Fixture
    $validHash = $script:package.Sha256
    $script:package.Sha256 = '0' * 64
    Must-Fail { Install-DemoModule $script:name } 'Checksum mismatch'
    if (Test-Path -LiteralPath (Join-Path $script:moduleRoot $script:name)) { throw 'Unverified module was published.' }
    $script:package.Sha256 = $validHash
    Write-Host 'Checksum mismatch is rejected before installation.'

    Reset-Fixture
    $script:package.Version = '1.2.4'
    Must-Fail { Install-DemoModule $script:name } 'Unexpected module identity'
    $script:package.Version = '1.2.3'
    Write-Host 'Wrong manifest identity is rejected.'

    Reset-Fixture
    $script:failDownloads = 99
    Must-Fail { Install-DemoModule $script:name } 'after 3 attempts'
    if ($script:downloadCalls -ne 3 -or @(Get-ChildItem -LiteralPath $script:moduleRoot -Force).Count) { throw 'Failed download was not bounded or cleaned.' }
    Write-Host 'Permanent network failure stops after three attempts without a partial installation.'
} finally {
    $env:PSModulePath = $originalModulePath
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction Stop
}
