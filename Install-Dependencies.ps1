<#
    .SYNOPSIS
        Ensures that all third‑party dependencies required by the validation system are
        available without requiring administrative privileges.

    .DESCRIPTION
        This script checks for the presence of EPPlus (4.5.3.3), NPOI (2.6.1) and
        PnP.PowerShell.  If any module is missing, it attempts to download and
        install it into the current user's module path.  All operations are
        performed in the context of the current user to avoid the need for
        administrative rights, which is important in regulated IVDR environments.

        Note: network access may be required to fetch packages from NuGet or
        PowerShell Gallery.  If your environment restricts outbound connectivity,
        you will need to prepackage the DLLs in the `Modules` folder ahead of
        execution.

    .EXAMPLE
        .\Install-Dependencies.ps1
        Ensures that the required modules exist in `Modules` and imports them.
#>
[CmdletBinding()]
param()

function Install-PackageIfMissing {
    param(
        [string]$PackageName,
        [string]$RequiredVersion,
        [string]$Provider = 'NuGet'
    )
    $modulePath = Join-Path -Path $PSScriptRoot -ChildPath 'Modules'
    if (-not (Test-Path -Path $modulePath)) {
        New-Item -ItemType Directory -Path $modulePath | Out-Null
    }
    # Determine if the DLL already exists (used for EPPlus/NPOI) or if the module
    # is installed (used for PnP.PowerShell).  This example uses a simplified
    # approach; in practice, consider using Find-Module/Find-Package to query
    # installed versions.
    $dll = Join-Path -Path $modulePath -ChildPath "$PackageName.dll"
    $module = Get-Module -ListAvailable -Name $PackageName -ErrorAction SilentlyContinue | Where-Object { $_.Version -eq $RequiredVersion }
    if ($module) {
        Write-Verbose "$PackageName $RequiredVersion is already installed."
        return
    }
    try {
        Write-Host "Installing $PackageName $RequiredVersion..." -ForegroundColor Cyan
        # Use Save-Package to download the package without touching the system wide
        # module directories.  If Save-Package fails due to policy restrictions,
        # instruct the user to manually place the DLLs in the Modules folder.
        Save-Package -Name $PackageName -RequiredVersion $RequiredVersion -ProviderName $Provider -Path $modulePath -Force -ErrorAction Stop | Out-Null
        Write-Host "$PackageName saved to $modulePath" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to install $PackageName: $_.Exception.Message."
        Write-Warning "Please download $PackageName version $RequiredVersion manually and place the DLL/module in the 'Modules' folder."
    }
}

# Install third‑party libraries if they are missing.  Versions are locked to meet
# validation requirements.
Install-PackageIfMissing -PackageName 'EPPlus' -RequiredVersion '4.5.3.3'
Install-PackageIfMissing -PackageName 'NPOI'   -RequiredVersion '2.6.1'

# PnP.PowerShell is distributed as a PowerShell module on PSGallery.  Use
# Save-Module instead of Install-Module to write it into the local Modules folder.
if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
    try {
        Write-Host "Installing PnP.PowerShell..." -ForegroundColor Cyan
        Save-Module -Name 'PnP.PowerShell' -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modules') -Force -ErrorAction Stop
        Write-Host "PnP.PowerShell saved." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to install PnP.PowerShell: $_.Exception.Message."
        Write-Warning "Please download the module manually and place it in the 'Modules' folder."
    }
}

# Update $env:PSModulePath so that modules in Modules folder are discoverable
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath 'Modules'
if (-not ($env:PSModulePath -like "*$modulePath*")) {
    $env:PSModulePath = "$modulePath;" + $env:PSModulePath
}

# Import modules explicitly if they were downloaded as DLLs.  EPPlus and NPOI
# ship assemblies that can be loaded via Add-Type.  Import PnP.PowerShell as module.
try {
    $epPlusDll = Join-Path -Path $modulePath -ChildPath 'EPPlus.dll'
    if (Test-Path $epPlusDll) {
        Add-Type -Path $epPlusDll -ErrorAction SilentlyContinue
    }
    $npoiDll = Join-Path -Path $modulePath -ChildPath 'NPOI.dll'
    if (Test-Path $npoiDll) {
        Add-Type -Path $npoiDll -ErrorAction SilentlyContinue
    }
    Import-Module 'PnP.PowerShell' -ErrorAction SilentlyContinue
} catch {
    Write-Warning "Failed to load one or more modules: $_.Exception.Message"
}

Write-Verbose "Dependency installation complete."