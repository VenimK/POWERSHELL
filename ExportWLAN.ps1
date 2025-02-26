# Set output encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

# Parameters
$defaultDirectory = "C:\WiFiProfiles"

# Add Windows Forms assembly for folder browser dialog
Add-Type -AssemblyName System.Windows.Forms

# Get currently connected WiFi network and its details
$interfaceInfo = netsh wlan show interfaces | Out-String
$currentNetwork = if ($interfaceInfo -match "SSID\s*:\s*(.*)") { $matches[1].Trim() } else { $null }
$signalStrength = if ($interfaceInfo -match "Signal\s*:\s*(.*)") { $matches[1].Trim() } else { "N/A" }

if ($currentNetwork) {
    Write-Host "Nu verbonden met: $currentNetwork" -ForegroundColor Green
    Write-Host "Signaalsterkte: $signalStrength" -ForegroundColor Green
}

# Create and configure folder browser dialog
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Selecteer waar u de WiFi-profielen wilt opslaan"
$folderBrowser.SelectedPath = $defaultDirectory

# Show folder browser dialog
Write-Host "Selecteer een map voor de WiFi-profielen..."
if ($folderBrowser.ShowDialog() -eq 'OK') {
    $outputDirectory = $folderBrowser.SelectedPath
    Write-Host "WiFi-profielen worden geexporteerd naar: $outputDirectory" -ForegroundColor Cyan
} else {
    Write-Host "Geen map geselecteerd. Standaard map wordt gebruikt: $defaultDirectory" -ForegroundColor Yellow
    $outputDirectory = $defaultDirectory
}

$logFile = Join-Path $outputDirectory "export_log.txt"

# Function to log messages
function Log-Message {
    param (
        [string]$message
    )
    Add-Content -Path $logFile -Value "$(Get-Date): $message"
}

# Create the output directory if it doesn't exist
if (-not (Test-Path -Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

# Get a list of all Wi-Fi profiles
$wifiProfilesOutput = netsh wlan show profiles | Select-String "All User Profile\s*:\s*(.*)" 

$profileNames = $wifiProfilesOutput | ForEach-Object {
    if ($_ -match "All User Profile\s*:\s*(.*)") {
        $matches[1].Trim()
    }
}

# Check if any profiles were found
if ($profileNames.Count -eq 0) {
    Write-Host "Geen WiFi-profielen gevonden."
    Log-Message "Geen WiFi-profielen gevonden."
    exit
}

$successfulExports = 0
$failedExports = 0

# Export each profile as an XML file
foreach ($profileName in $profileNames) {
    try {
        # Get the profile details
        $profileInfo = netsh wlan show profile name="$profileName" | Out-String
        $connectionMode = if ($profileInfo -match "Connection mode\s*:\s*(.*)") { 
            $matches[1].Trim() 
        } else { 
            "Onbekend" 
        }

        # Get the profile key (password)
        $profileInfo = netsh wlan show profile name="$profileName" key=clear
        $keyMatch = $profileInfo | Select-String "Key Content\s*:\s*(.*)"
        $key = if ($keyMatch -and $keyMatch.Matches.Groups.Count -gt 1) {
            $keyMatch.Matches.Groups[1].Value
        } else {
            "Geen wachtwoord"
        }

        # Export the Wi-Fi profile
        $exportResult = netsh wlan export profile name="$profileName" folder="$outputDirectory" key=clear
        
        # Display profile info with color if it's the current network
        $displayInfo = "Exporteren van: $profileName"
        $displayInfo += "`n  - Wachtwoord: $key"
        $displayInfo += "`n  - Verbindingsmodus: $connectionMode"
        if ($profileName -eq $currentNetwork) {
            $displayInfo += "`n  - Signaalsterkte: $signalStrength"
            Write-Host $displayInfo -ForegroundColor Green
        } else {
            Write-Host $displayInfo
        }
        
        Log-Message "Profiel succesvol geexporteerd: $profileName"
        Log-Message "Netwerk: $profileName - Wachtwoord: $key - Verbindingsmodus: $connectionMode"
        $successfulExports++
    } catch {
        Write-Host "Export mislukt: $profileName. Fout: $_" -ForegroundColor Red
        Log-Message "Export mislukt: $profileName. Fout: $_"
        $failedExports++
    }
}

# Permission settings (optional)
# Get-ChildItem -Path $outputDirectory | ForEach-Object { $_.SetAccessControl(...) } # Add security as needed

# Summary output
Write-Host "`nExporteren van WiFi-profielen voltooid." -ForegroundColor Cyan
Write-Host "Totaal succesvol: $successfulExports" -ForegroundColor Cyan
Write-Host "Totaal mislukt: $failedExports" -ForegroundColor $(if ($failedExports -eq 0) { 'Cyan' } else { 'Red' })
Log-Message "Exporteren voltooid: $successfulExports succesvol, $failedExports mislukt."