# Set output encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

# Add Windows Forms assembly for folder browser dialog
Add-Type -AssemblyName System.Windows.Forms

# Create and configure folder browser dialog
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Selecteer de map met WiFi-profielen om te importeren"
$folderBrowser.ShowNewFolderButton = $false

# Show folder browser dialog
Write-Host "Selecteer de map met WiFi-profielen..."
if ($folderBrowser.ShowDialog() -ne 'OK') {
    Write-Host "Geen map geselecteerd. Script wordt afgesloten." -ForegroundColor Red
    exit
}

$importDirectory = $folderBrowser.SelectedPath
Write-Host "WiFi-profielen worden geimporteerd uit: $importDirectory" -ForegroundColor Cyan
$logFile = Join-Path $importDirectory "import_log.txt"

# Function to log messages
function Log-Message {
    param (
        [string]$message
    )
    Add-Content -Path $logFile -Value "$(Get-Date): $message"
}

# Get currently connected WiFi network
$currentNetwork = (netsh wlan show interfaces | Select-String "SSID\s*:\s*(.*)").Matches.Groups[1].Value.Trim()
Write-Host "Nu verbonden met: $currentNetwork" -ForegroundColor Green

# Get all XML files in the directory
$wifiProfiles = Get-ChildItem -Path $importDirectory -Filter "*.xml"

if ($wifiProfiles.Count -eq 0) {
    Write-Host "Geen WiFi-profielen gevonden in de geselecteerde map." -ForegroundColor Red
    Log-Message "Geen WiFi-profielen gevonden in: $importDirectory"
    exit
}

Write-Host "Gevonden WiFi-profielen: $($wifiProfiles.Count)" -ForegroundColor Cyan

$successfulImports = 0
$failedImports = 0

foreach ($profile in $wifiProfiles) {
    try {
        # Get profile name from XML content
        $xmlContent = Get-Content $profile.FullName
        $profileName = ([xml]$xmlContent).WLANProfile.name

        # Import the profile
        $result = netsh wlan add profile filename="$($profile.FullName)" user=all
        
        if ($result -match "is toegevoegd" -or $result -match "is added") {
            # Try to get the password from the XML
            $key = ([xml]$xmlContent).WLANProfile.MSM.security.sharedKey.keyMaterial
            
            if ($profileName -eq $currentNetwork) {
                Write-Host "Importeren van: $profileName - Wachtwoord: $key" -ForegroundColor Green
            } else {
                Write-Host "Importeren van: $profileName - Wachtwoord: $key"
            }
            
            Log-Message "Profiel succesvol geimporteerd: $profileName"
            Log-Message "Netwerk: $profileName - Wachtwoord: $key"
            $successfulImports++
        } else {
            throw "Import mislukt: $result"
        }
    }
    catch {
        Write-Host "Import mislukt: $($profile.Name). Fout: $_" -ForegroundColor Red
        Log-Message "Import mislukt: $($profile.Name). Fout: $_"
        $failedImports++
    }
}

# Summary output
Write-Host "`nImporteren van WiFi-profielen voltooid." -ForegroundColor Cyan
Write-Host "Totaal succesvol: $successfulImports" -ForegroundColor Cyan
Write-Host "Totaal mislukt: $failedImports" -ForegroundColor $(if ($failedImports -eq 0) { 'Cyan' } else { 'Red' })
Log-Message "Importeren voltooid: $successfulImports succesvol, $failedImports mislukt."
