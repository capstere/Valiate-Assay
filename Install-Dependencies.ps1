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
function Install-PackageIfMissing {
    param(
        [string]$ModulePath,
        [string]$PackageName,
        [string]$RequiredVersion,
        [string]$Provider = 'NuGet'
    )
    if (-not (Test-Path -LiteralPath $ModulePath)) {
        New-Item -ItemType Directory -Path $ModulePath | Out-Null
    }
    $dll    = Join-Path -Path $ModulePath -ChildPath "$PackageName.dll"
    $module = Get-Module -ListAvailable -Name $PackageName -ErrorAction SilentlyContinue | Where-Object { $_.Version -eq $RequiredVersion }
    if ($module -or (Test-Path -LiteralPath $dll)) { return }
    try {
        Write-Host "Installing $PackageName $RequiredVersion..." -ForegroundColor Cyan
        Save-Package -Name $PackageName -RequiredVersion $RequiredVersion -ProviderName $Provider -Path $ModulePath -Force -ErrorAction Stop | Out-Null
        Write-Host "$PackageName saved to $ModulePath" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to install $PackageName: $($_.Exception.Message)"
        Write-Warning "Please download $PackageName version $RequiredVersion manually and place it in '$ModulePath'."
    }
}

function Install-Dependencies {
    param([string]$ModulePath)
    Install-PackageIfMissing -ModulePath $ModulePath -PackageName 'EPPlus' -RequiredVersion '4.5.3.3'
    Install-PackageIfMissing -ModulePath $ModulePath -PackageName 'NPOI'   -RequiredVersion '2.6.1'
    if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
        try {
            Write-Host "Installing PnP.PowerShell..." -ForegroundColor Cyan
            Save-Module -Name 'PnP.PowerShell' -Path $ModulePath -Force -ErrorAction Stop
            Write-Host "PnP.PowerShell saved." -ForegroundColor Green
        } catch {
            Write-Warning "Failed to install PnP.PowerShell: $($_.Exception.Message)"
            Write-Warning "Please download the module manually and place it in '$ModulePath'."
        }
    }
    if (-not ($env:PSModulePath -like "*$ModulePath*")) {
        $env:PSModulePath = "$ModulePath;" + $env:PSModulePath
    }
}

function Load-EPPlus {
    param([string]$ModulePath)
    $dll = Join-Path -Path $ModulePath -ChildPath 'EPPlus.dll'
    if (Test-Path -LiteralPath $dll) {
        Add-Type -Path $dll -ErrorAction SilentlyContinue
    }
}

function Load-NPOI {
    param([string]$ModulePath)
    $dll = Join-Path -Path $ModulePath -ChildPath 'NPOI.dll'
    if (Test-Path -LiteralPath $dll) {
        Add-Type -Path $dll -ErrorAction SilentlyContinue
    }
}

Write-Verbose "Dependency installation module loaded."
