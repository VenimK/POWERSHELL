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
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" Write-Ascii $timestamp
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
        Log-Message "Installing module: $ModuleName..."
        Install-Module -Name $ModuleName -Force -Scope CurrentUser -AllowClobber
        Log-Message "$ModuleName installed successfully."
    } else {
        Log-Message "Module $ModuleName is already installed."
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
        Log-Message "Installatie gestart voor $PackageDisplayName. Output: $installResult."
        
        if ($LASTEXITCODE -eq 0) {
            Log-Message "$PackageDisplayName installatie successvol."
            Write-Host "$PackageDisplayName installatie successvol." -ForegroundColor Green
        } else {
            if ($installResult -like "*Applicatie is al geinstalleerd*") {
                Log-Message "$PackageDisplayName is al geinstalleerd. Geen verder aktie nodig."
            } elseif ($installResult -like "*Poging voor update*") {
                Log-Message "Poging voor update app $PackageDisplayName. Resultaat: $installResult"
            } else {
                Log-Message "Poging gefaald voor app ${PackageDisplayName}. Fout: $installResult"
                Write-Host "Gefaald voor installatie $PackageDisplayName. Controleer log voor details." -ForegroundColor Red
            }
        }
    }
}

function Uninstall-WingetPackage {
    param (
        [string]$PackageName,
        [string]$PackageDisplayName
    )

    Log-Message "Start verwijderen van $PackageDisplayName."
    try {
        if (winget uninstall $PackageName --silent --source winget) {
            Log-Message "$PackageDisplayName successvol verwijderd."
            Write-Host "$PackageDisplayName successvol verwijdererd." -ForegroundColor Red
        } else {
            throw "Uninstallation command for ${PackageDisplayName} did not complete as expected."
        }
    } catch {
        Log-Message "Failed to uninstall package ${PackageDisplayName}: $($_.Exception.Message)"
        Write-Host "Failed to uninstall $PackageDisplayName. Check logs for details." -ForegroundColor Red
    }
}

function Install-PackageManagementAndNuGetProvider {
    Log-Message "Control voor PackageManagement module."
    $pkgManagementModule = Get-Module -Name PackageManagement -ListAvailable
    if ($pkgManagementModule) {
        Log-Message "PackageManagement module versie $($pkgManagementModule.Version) is momemteel geladen."
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
        Log-Message "PackageManagement succesvol geinstalleerd."
    } catch {
        Log-Message "Failed to install PackageManagement: $($_.Exception.Message)"
    }

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $null = Install-PackageProvider -Name NuGet -Force -ErrorAction Stop
        Log-Message "NuGet providersuccesvol geinstalleerd."
    } catch {
        Log-Message "Failed to install NuGet provider: $($_.Exception.Message)"
    }
}

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

Install-PackageManagementAndNuGetProvider

# Install critical modules
Write-Host "Installeren van Modules" -ForegroundColor Blue
$modules = @("PowerShellGet", "WriteAscii", "WinGet", "PSWindowsUpdate")
foreach ($module in $modules) {
    Install-ModuleIfNeeded -ModuleName $module
}

# Import needed modules
Write-Host "Importeren van Modules" -ForegroundColor Blue
Import-Module -Name WriteAscii
Import-Module -Name WinGet
Import-Module -Name PSWindowsUpdate

# Add Windows Update service manager
Write-Host "Toevoegen van WindowsUpdate" -ForegroundColor Blue
Add-WUServiceManager -ServiceID 7971f918-a847-4430-9279-4a52d1efe18d -Confirm:$false

# Start script
Log-Message "Start van MusicLover Setup"
Write-Ascii "Start MusicLover - Updates" -ForegroundColor Blue
Get-WuList -MicrosoftUpdate -Verbose
Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -Verbose -IgnoreReboot

Log-Message "Toevoegen van Winget"
Write-Host "Toevoegen van Winget" -ForegroundColor Blue

Log-Message "Installeren Van MusicLover Soft-Pack"
Write-Ascii "Installeren Van MusicLover Soft-Pack" -ForegroundColor Green

# List of packages to install
Write-Host "Installeren van Applicaties" -ForegroundColor Blue $wingetPackages
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
    Write-Progress -Activity "Installeren applicaties" -Status "Installeren: $($package.Value)" -PercentComplete (($currentPackage / $totalPackages) * 100)
    Install-WingetPackage -PackageName $package.Key -PackageDisplayName $package.Value
}

# Summary of installed packages
Log-Message "MusicLover UnInstaller"
Write-Ascii "MusicLover UnInstaller" -ForegroundColor Red

# List of packages to uninstall
Write-Host "Verwijderen van Applicaties" -ForegroundColor Blue $wingetUninstallPackages
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
    Write-Progress -Activity "Verwijderen applicaties" -Status "Verwijderen: $($package.Value)" -PercentComplete (($currentUninstallPackage / $totalUninstallPackages) * 100)
    Uninstall-WingetPackage -PackageName $package.Key -PackageDisplayName $package.Value
}

Log-Message "Setup Compleet!"
Write-Host "Setup Compleet!" -ForegroundColor Cyan
Write-Ascii "Setup Compleet!" -ForegroundColor DarkBlue