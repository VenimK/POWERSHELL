# WiFi Verbinding Script
# Vooraf ingestelde netwerkinformatie
$SSID = "Donk86"
$Password = "poiu1234"

# Genereer WiFi QR Code met QR Server API
$WifiString = "WIFI:T:WPA;S:$SSID;P:$Password;;"
Write-Host "QR Code aan het genereren voor WiFi netwerk..."

# Maak de QR code URL (met qrserver.com API)
$qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + [uri]::EscapeDataString($WifiString)

# Download de QR code
$QRPath = Join-Path $PSScriptRoot "WiFiQRCode.png"
$webClient = New-Object System.Net.WebClient
$webClient.DownloadFile($qrUrl, $QRPath)
Write-Host "QR Code gegenereerd op locatie: $QRPath"

Write-Host "Automatisch verbinden met $SSID..."

# Controleer of het netwerk bestaat
$network = netsh wlan show networks | Select-String -Pattern $SSID

if ($network) {
    Write-Host "Netwerk $SSID gevonden. Bezig met verbinden..."
    
    try {
        # Maak een tijdelijk profiel XML
        $profileXml = @"
<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>$SSID</name>
    <SSIDConfig>
        <SSID>
            <name>$SSID</name>
        </SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>WPA2PSK</authentication>
                <encryption>AES</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>$Password</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>
"@
        # Voeg het profiel toe
        $profileXml | Set-Content ".\WiFiProfile.xml"
        netsh wlan add profile filename=".\WiFiProfile.xml"
        Remove-Item ".\WiFiProfile.xml"

        # Probeer verbinding te maken met het netwerk
        netsh wlan connect name=$SSID
        
        # Wacht even om de verbindingsstatus te controleren
        Start-Sleep -Seconds 5
        
        # Controleer of we verbonden zijn
        $connectionStatus = netsh wlan show interfaces | Select-String -Pattern $SSID
        
        if ($connectionStatus) {
            Write-Host "Succesvol verbonden met $SSID"
        } else {
            Write-Host "Kan niet verbinden met $SSID. Controleer of het wachtwoord correct is."
        }
    }
    catch {
        Write-Host "Er is een fout opgetreden tijdens het verbinden: $_"
    }
} else {
    Write-Host "Netwerk $SSID niet gevonden. Controleer of het netwerk binnen bereik is."
}