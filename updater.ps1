$ErrorActionPreference = 'Stop'

# Check for elevated privileges
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-Not (Test-Administrator)) {
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Check and set execution policy
$executionPolicy = Get-ExecutionPolicy
if ($executionPolicy -ne "RemoteSigned") {
    try {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
    } catch {
        Write-Output "Error: Unable to set execution policy to RemoteSigned."
        Pause
        exit
    }
}


Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201
Install-Module -Name WriteAscii -Force -AllowClobber 
Install-Module -Name WinGet -Force
Install-Module -Name PSWindowsUpdate -Force

Import-Module -Name PSWindowsUpdate
Add-WUServiceManager -ServiceID 7971f918-a847-4430-9279-4a52d1efe18d -Confirm:$false


$UpdateList = Get-WUList -Verbose
$UpdateList | Select Size, Status, ComputerName, KB, Title, IsDownloaded, IsHidden, IsInstalled, IsMandatory, IsUninstallable, RebootRequired, IsPresent | Export-Csv C:\Get-WUList-Example.csv -NoTypeInformation -Append


Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -Verbose -IgnoreReboot
