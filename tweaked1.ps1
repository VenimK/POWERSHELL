<#
.SYNOPSIS
    This script installs necessary PowerShell modules and applications 
    using winget on a new PC setup.

.DESCRIPTION
    It checks for and installs critical PowerShell modules, performs Windows updates,
    and installs/removes applications based on predefined lists.
#>

param (
    [string]$LogFilePath = "install_log.txt",
    [switch]$Verbose
)

# Logging function
function Log-Message {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $Message"
    
    # Write to log file
    Add-Content -Path $LogFilePath -Value $logEntry

    # Write to console output, always
    Write-Host $logEntry
}

function Install-ModuleIfNeeded {
    param (
        [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Log-Message "Installeren van module: $ModuleName..."
        Install-Module -Name $ModuleName -Force -Scope CurrentUser -AllowClobber
        Log-Message "$ModuleName installed successfully."
    } else {
        Log-Message "Module $ModuleName is al geinstalleerd."
    }
}

function Check-InstalledPackage {
    param (
        [string]$PackageDisplayName
    )
    return winget list --search $PackageDisplayName | Select-String $PackageDisplayName
}

function Install-WingetPackage {
    param (
        [string]$PackageName,
        [string]$PackageDisplayName
    )

    Log-Message "Controle installatie status voor $PackageDisplayName."
    
    if (Check-InstalledPackage -PackageDisplayName $PackageDisplayName) {
        Log-Message "$PackageDisplayName is al geinstalleerd. Sla over."
    } else {
        Log-Message "Start installatie van $PackageDisplayName."
        Write-Host "Installeren $PackageDisplayName..." -ForegroundColor Cyan
        
        $installResult = winget install $PackageName --silent --accept-source-agreements --accept-package-agreements 2>&1
        Log-Message "Install command executed for $PackageDisplayName. Output: $installResult."
        
        if ($LASTEXITCODE -eq 0) {
            Log-Message "$PackageDisplayName installatie successvol."
            Write-Host "$PackageDisplayName installatie succesvol." -ForegroundColor Green
        } else {
            if ($installResult -like "*Applicatie is al geinstalleerd*") {
                Log-Message "$PackageDisplayName is al geinstalleerd. No further action required."
            } elseif ($installResult -like "*Controle voor ugrade*") {
                Log-Message "Attempted to upgrade existing package $PackageDisplayName. Result: $installResult"
            } else {
                Log-Message "Failed to install package ${PackageDisplayName}. Error: $installResult"
                Write-Host "Failed to install $PackageDisplayName. Check logs for details." -ForegroundColor Red
            }
        }
    }
}

function Uninstall-WingetPackage {
    param (
        [string]$PackageName,
        [string]$PackageDisplayName
    )

    Log-Message "Checking installation status for $PackageDisplayName."

    # Check if the package is installed before trying to uninstall
    if (-not Check-InstalledPackage -PackageDisplayName $PackageDisplayName) {
        Log-Message "$PackageDisplayName is already removed or not found. No action required."
        Write-Host "$PackageDisplayName is already removed or not found." -ForegroundColor Yellow
        return  # Exit the function if the package is not installed
    }

    Log-Message "Starting uninstallation of $PackageDisplayName."
    
    try {
        # Capture the output and error of the uninstall command
        $uninstallResult = winget uninstall $PackageName --silent --source winget 2>&1
        if ($LASTEXITCODE -eq 0) {
            Log-Message "$PackageDisplayName successfully removed."
            Write-Host "$PackageDisplayName successfully removed." -ForegroundColor Green
        } else {
            Log-Message "Failed to uninstall package ${PackageDisplayName}. Command output: $uninstallResult"
            Write-Host "Failed to uninstall $PackageDisplayName. Check logs for details." -ForegroundColor Red
        }
    } catch {
        Log-Message "Failed to uninstall package ${PackageDisplayName}: $($_.Exception.Message)"
        Write-Host "Failed to uninstall $PackageDisplayName. Check logs for details." -ForegroundColor Red
    }
}

function Install-PackageManagementAndNuGetProvider {
    Log-Message "Checking for PackageManagement module."
    $pkgManagementModule = Get-Module -Name PackageManagement -ListAvailable
    if ($pkgManagementModule) {
        Log-Message "PackageManagement module version $($pkgManagementModule.Version) is currently loaded."
        return  # Exit the function here if it's already loaded
    }
    
    # Try to uninstall PackageManagement if it's loaded
    try {
        Uninstall-Module -Name PackageManagement -Force -ErrorAction Stop
        Log-Message "PackageManagement uninstalled successfully, if it was present."
    } catch {
        Log-Message "Failed to uninstall PackageManagement: $($_.Exception.Message)" 
    }

    try {
        Install-Module -Name PackageManagement -Force -AllowClobber -ErrorAction Stop
        Log-Message "PackageManagement installed successfully."
    } catch {
        Log-Message "Failed to install PackageManagement: $($_.Exception.Message)"
    }

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $null = Install-PackageProvider -Name NuGet -Force -ErrorAction Stop
        Log-Message "NuGet provider installed successfully."
    } catch {
        Log-Message "Failed to install NuGet provider: $($_.Exception.Message)"
    }
}

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

Install-PackageManagementAndNuGetProvider

# Install critical modules
$modules = @("PowerShellGet", "WriteAscii", "WinGet", "PSWindowsUpdate")
foreach ($module in $modules) {
    Install-ModuleIfNeeded -ModuleName $module
}

# Import needed modules
Import-Module -Name WriteAscii
Import-Module -Name WinGet
Import-Module -Name PSWindowsUpdate

# Add Windows Update service manager
Add-WUServiceManager -ServiceID 7971f918-a847-4430-9279-4a52d1efe18d -Confirm:$false

# Start script
Log-Message "Start MusicLover Setup"
Write-Ascii "Start MusicLover - Updates" -ForegroundColor Blue
Get-WuList -MicrosoftUpdate -Verbose
Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -Verbose -IgnoreReboot

Log-Message "Toevoegen Winget"
Write-Host "Toevoegen van Winget" -ForegroundColor Blue

Log-Message "Installeren Van MusicLover SoftPack"
Write-Ascii "Installeren Van MusicLover SoftPack" -ForegroundColor Green

# List of packages to install
$wingetPackages = @{
    "7-Zip" = "7-Zip"
    "Google.Chrome" = "Google Chrome"
    "VideoLAN.VLC" = "VLC Media Player"
    "Adobe.Acrobat.Reader.64-bit" = "Adobe Reader DC 64-Bit"
    "BelgianGovernment.eIDViewer" = "ID Viewer"
    "BelgianGovernment.Belgium-eIDmiddleware" = "ID Software"
    "SomePythonThings.WingetUIStore" = "Winget UI Store"
}

$totalPackages = $wingetPackages.Count
$currentPackage = 0

foreach ($package in $wingetPackages.GetEnumerator()) {
    $currentPackage++
    Write-Progress -Activity "Installing packages" -Status "Installing: $($package.Value)" -PercentComplete (($currentPackage / $totalPackages) * 100)
    Install-WingetPackage -PackageName $package.Key -PackageDisplayName $package.Value
}

# Summary of installed packages
Log-Message "MusicLover UnInstaller"
Write-Ascii "MusicLover UnInstaller" -ForegroundColor Red

# List of packages to uninstall
$wingetUninstallPackages = @{
    "9N1SQW2NKPDS" = "McAfee"
    "Xbox Game Bar Plugin" = "Xbox Game Bar Plugin"
    "Xbox Game Bar" = "Xbox Game Bar"
    "Xbox Identity Provider" = "Xbox Identity Provider"
    "Xbox Game Speech Window" = "Xbox Game Speech Window"
    "ExpressVPN" = "ExpressVPN"
    "Microsoft.MicrosoftOfficeHub_8wekyb3d8bbwe" = "Office - OneDrive - MSN - Nieuws - Evernote"
    "Microsoft.OneDrive" = "OneDrive"
    "Microsoft.BingWeather_8wekyb3d8bbwe" = "Bing Weer"
    "Microsoft.BingNews_8wekyb3d8bbwe" = "Bing Nieuws"
    "Evernote.Evernote_q4d96b2w5wcc2" = "Evernote"
}

$totalUninstallPackages = $wingetUninstallPackages.Count
$currentUninstallPackage = 0

foreach ($package in $wingetUninstallPackages.GetEnumerator()) {
    $currentUninstallPackage++
    Write-Progress -Activity "Uninstalling packages" -Status "Uninstalling: $($package.Value)" -PercentComplete (($currentUninstallPackage / $totalUninstallPackages) * 100)
    Uninstall-WingetPackage -PackageName $package.Key -PackageDisplayName $package.Value
}

Log-Message "Setup Compleet!"
Write-Host "Setup Compleet!" -ForegroundColor Cyan
Write-Ascii "Setup Compleet!" -ForegroundColor DarkBlue