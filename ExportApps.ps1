# Script to export installed applications
$outputPath = Join-Path $PSScriptRoot "installed-apps.json"

# Get installed applications from multiple sources
$apps = @()

# Get apps from Programs and Features
$apps += Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
    Where-Object { $_.DisplayName -ne $null } |
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate

# Get apps from Programs and Features (32 bit)
$apps += Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Where-Object { $_.DisplayName -ne $null } |
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate

# Get Microsoft Store apps for current user
$apps += Get-AppxPackage -User $env:USERNAME |
    Select-Object @{N='DisplayName';E={$_.Name}}, 
                  @{N='DisplayVersion';E={$_.Version}},
                  @{N='Publisher';E={$_.Publisher}},
                  @{N='InstallDate';E={$_.InstallDate}}

# Convert to JSON and export
$apps | ConvertTo-Json | Out-File -FilePath $outputPath -Encoding UTF8

Write-Host "Successfully exported installed applications to: $outputPath"
Write-Host "Total apps found: $($apps.Count)"
