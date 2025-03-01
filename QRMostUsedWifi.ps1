# =================================================================
# WiFi QR Code Generator
# =================================================================
# Dit script genereert een QR code voor het meest gebruikte WiFi netwerk
# van de laatste 60 dagen. Als er geen geschiedenis beschikbaar is,
# wordt het huidige netwerk gebruikt.
#
# Gebruik:
# 1. Start PowerShell als Administrator
# 2. Navigeer naar de map met dit script
# 3. Voer uit: .\QRMostUsedWifi.ps1
#
# De QR code wordt opgeslagen op het bureaublad als 'WiFiQRCode.png'
# Deze QR code kan gescand worden met een telefoon/tablet om
# automatisch verbinding te maken met het WiFi netwerk.
# =================================================================

# Get the most used WiFi profile from last 60 days and generate QR code for it

try {
    Write-Host "WiFi verbindingsgeschiedenis van de laatste 60 dagen ophalen..." -ForegroundColor Cyan
    
    # Get event history for the last 60 days
    $startTime = (Get-Date).AddDays(-60)
    $events = Get-WinEvent -FilterHashtable @{
        LogName = "Microsoft-Windows-WLAN-AutoConfig/Operational"
        ID = @(8001)  # Only successful connections
        StartTime = $startTime
    } -ErrorAction Stop

    # Count connections per network
    $networkStats = @{}
    foreach ($event in $events) {
        $networkName = ($event.Message -split "SSID: ")[1] -split "`r`n" | Select-Object -First 1
        if (-not $networkStats.ContainsKey($networkName)) {
            $networkStats[$networkName] = @{
                Name = $networkName
                Connections = 0
            }
        }
        $networkStats[$networkName].Connections++
    }

    # Get the most connected network
    $mostUsedNetwork = $networkStats.Values | Sort-Object -Property Connections -Descending | Select-Object -First 1

    # If no connections found in history, use current connection
    if (-not $mostUsedNetwork) {
        Write-Host "Geen verbindingen gevonden in de laatste 60 dagen, huidige verbinding wordt gebruikt..." -ForegroundColor Yellow
        
        # Get current connection info
        $currentInterface = netsh wlan show interfaces | Out-String
        $currentSSID = if ($currentInterface -match "SSID\s*:\s*(.*)") { $matches[1].Trim() }
        
        if (-not $currentSSID) {
            Write-Host "Geen actieve WiFi verbinding gevonden." -ForegroundColor Red
            exit
        }
        
        $selectedNetwork = $currentSSID
        Write-Host "Huidige verbinding wordt gebruikt: $selectedNetwork" -ForegroundColor Cyan
    } else {
        $selectedNetwork = $mostUsedNetwork.Name
        Write-Host "Meest gebruikte netwerk in laatste 60 dagen: $selectedNetwork ($($mostUsedNetwork.Connections) verbindingen)" -ForegroundColor Cyan
    }

    # Get the password for this network
    $networkInfo = netsh wlan show profile name="$selectedNetwork" key=clear
    $password = ($networkInfo | Select-String "Key Content\s+:\s+(.+)").Matches.Groups[1].Value

    if (-not $password) {
        Write-Host "Kan het wachtwoord voor dit netwerk niet ophalen. Mogelijk zijn beheerrechten vereist." -ForegroundColor Red
        exit
    }

    # Generate WiFi QR Code
    $WifiString = "WIFI:T:WPA;S:$selectedNetwork;P:$password;;"
    Write-Host "QR Code aan het genereren voor WiFi netwerk..."

    # Create the QR code URL
    $qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + [uri]::EscapeDataString($WifiString)

    # Download the QR code to desktop
    $QRPath = Join-Path ([Environment]::GetFolderPath('Desktop')) "WiFiQRCode.png"
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($qrUrl, $QRPath)
    Write-Host "QR Code gegenereerd op locatie: $QRPath" -ForegroundColor Green

} catch {
    if ($_.Exception.Message -match "No events were found") {
        Write-Host "Geen WiFi gebeurtenissen gevonden in de laatste 60 dagen." -ForegroundColor Yellow
        # Try to use current connection as fallback
        & $MyInvocation.MyCommand.Path
    } else {
        Write-Host "Er is een fout opgetreden: $($_.Exception.Message)" -ForegroundColor Red
    }
}
