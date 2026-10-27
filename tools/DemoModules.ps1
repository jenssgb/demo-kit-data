# Known Microsoft packages for the installer, including Exchange's PowerShellGet dependencies.
# Fallback follows https://learn.microsoft.com/powershell/gallery/how-to/working-with-packages/manual-download
function Get-DemoModulePackage([string]$Name) {
    $packages = @{
        'Microsoft.Graph.Authentication' = @{
            Version = '2.41.2'; Dependencies = @()
            Sha256 = 'A4F41CAB6FC3CE0891A9295F191A218E0CFE692F21C6395763F339260C00D98C'
        }
        'ExchangeOnlineManagement' = @{
            # 3.10+ requires PowerShell 7.6; 3.9.2 also supports existing 7.4/7.5 demo VMs.
            Version = '3.9.2'; Dependencies = @('PowerShellGet')
            Sha256 = '3C9D862971E7E620FB3B3B76DCC5BF8949E8E40CF860FA1C3099FBC78669EF93'
        }
        'PowerShellGet' = @{
            Version = '2.2.5'; Dependencies = @('PackageManagement')
            Sha256 = '6B8CEBF2A464EAEB31B0A6D627355C30D9D1899DBA0CE3BDD0D4E7AFCA148673'
        }
        'PackageManagement' = @{
            Version = '1.4.8.1'; Dependencies = @()
            Sha256 = '7E1F8A75B6BC8A83D8ABFF79F6690FC1DFBD534FD3E5733D97E19BCB5954C13E'
        }
    }
    if (-not $packages.ContainsKey($Name)) { throw "Unsupported installer module: $Name" }
    $packages[$Name]
}

function Get-DemoModuleRoot {
    if (-not $env:LOCALAPPDATA) { throw 'LOCALAPPDATA is required for current-user module installation.' }
    Join-Path $env:LOCALAPPDATA 'DemoKit\Modules'
}

function Assert-DemoModule([string]$Manifest, [string]$Name, [string]$Version) {
    $info = Test-ModuleManifest -Path $Manifest -ErrorAction Stop
    if ($info.Name -ne $Name -or $info.Version -ne [version]$Version) {
        throw "Unexpected module identity in ${Manifest}: expected $Name $Version."
    }
}

function Install-DemoModule([string]$Name) {
    $package = Get-DemoModulePackage $Name
    $root = Get-DemoModuleRoot
    if ($root -notin @($env:PSModulePath -split ';')) { $env:PSModulePath = "$root;$env:PSModulePath" }
    foreach ($dependency in $package.Dependencies) { Install-DemoModule $dependency | Out-Null }
    $available = Get-Module -ListAvailable -Name $Name |
        Where-Object { $_.Version -eq [version]$package.Version } | Select-Object -First 1
    if ($available) {
        Assert-DemoModule $available.Path $Name $package.Version
        Write-Host "    Module ready: $Name $($package.Version)" -ForegroundColor DarkGray
        return $available.Path
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Write-Host "    Installing $Name $($package.Version) for the current user ..." -ForegroundColor Cyan
    try {
        if ($PSVersionTable.PSEdition -ne 'Core') {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.206 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
        }
        $options = @{
            Name = $Name; RequiredVersion = $package.Version; Repository = 'PSGallery'
            Scope = 'CurrentUser'; Force = $true; AllowClobber = $true; ErrorAction = 'Stop'
        }
        if ((Get-Command Install-Module -ErrorAction Stop).Parameters.ContainsKey('AcceptLicense')) { $options.AcceptLicense = $true }
        Install-Module @options | Out-Null
        $installed = Get-Module -ListAvailable -Name $Name |
            Where-Object { $_.Version -eq [version]$package.Version } | Select-Object -First 1
        if (-not $installed) { throw 'The package manager did not install a discoverable module.' }
        Assert-DemoModule $installed.Path $Name $package.Version
        return $installed.Path
    } catch {
        Write-Host "    Gallery installation failed for ${Name}: $_" -ForegroundColor DarkYellow
        Write-Host '    Trying the official PowerShell Gallery CDN with a pinned SHA-256 checksum ...' -ForegroundColor DarkYellow
    }

    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $stage = Join-Path $root ('.demo-kit-stage-' + [guid]::NewGuid().ToString('N'))
    $backup = $null
    $published = $false
    $destination = Join-Path (Join-Path $root $Name) $package.Version
    New-Item -ItemType Directory -Path $stage | Out-Null
    try {
        $archive = Join-Path $stage 'package.zip'
        $uri = "https://cdn.powershellgallery.com/packages/$($Name.ToLowerInvariant()).$($package.Version).nupkg"
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                Invoke-WebRequest -Uri $uri -OutFile $archive -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
                break
            } catch {
                if ($attempt -eq 3) { throw "Cannot download $Name from the official Gallery CDN after 3 attempts: $_" }
                Write-Host "    Download attempt $attempt failed for ${Name}: $_; retrying ..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds (2 * $attempt)
            }
        }
        if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash -ne $package.Sha256) {
            throw "Checksum mismatch for $Name $($package.Version); the package will not be installed."
        }
        $expanded = Join-Path $stage 'module'
        Expand-Archive -LiteralPath $archive -DestinationPath $expanded -ErrorAction Stop
        foreach ($metadata in @('_rels', 'package', '[Content_Types].xml', "$Name.nuspec", '.signature.p7s')) {
            $path = Join-Path $expanded $metadata
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop }
        }
        Assert-DemoModule (Join-Path $expanded "$Name.psd1") $Name $package.Version
        New-Item -ItemType Directory -Force -Path (Split-Path $destination) | Out-Null
        if (Test-Path -LiteralPath $destination) {
            $backup = Join-Path $root ('.demo-kit-backup-' + [guid]::NewGuid().ToString('N'))
            Move-Item -LiteralPath $destination -Destination $backup -ErrorAction Stop
        }
        try {
            Move-Item -LiteralPath $expanded -Destination $destination -ErrorAction Stop
            $published = $true
        }
        catch {
            if ($backup) { Move-Item -LiteralPath $backup -Destination $destination -ErrorAction Stop; $backup = $null }
            throw
        }
        Write-Host "    Installed and verified: $Name $($package.Version)" -ForegroundColor Green
        Join-Path $destination "$Name.psd1"
    } finally {
        Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction Stop
        if ($backup -and $published) { Remove-Item -LiteralPath $backup -Recurse -Force -ErrorAction Stop }
    }
}
