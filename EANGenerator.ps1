# =================================================================
# EAN-13 Barcode Generator
# =================================================================
# Dit script genereert een EAN-13 barcode en slaat deze op als afbeelding.
# Het script valideert de invoer en berekent automatisch het controlegetal.
#
# Gebruik (kies een van deze methodes):
#
# Methode 1 (Tijdelijk bypass):
# 1. Start PowerShell als Administrator
# 2. Navigeer naar de map met dit script
# 3. Voer uit: powershell -ExecutionPolicy Bypass -File EANGenerator.ps1
#
# Methode 2 (Script ondertekenen):
# 1. Start PowerShell als Administrator
# 2. Voer uit: Set-ExecutionPolicy RemoteSigned
# 3. Navigeer naar de map met dit script
# 4. Voer uit: .\EANGenerator.ps1
#
# Methode 3 (Voor één keer):
# 1. Rechtsklik op het script
# 2. Kies 'Met PowerShell uitvoeren'
# =================================================================

function Calculate-EANChecksum {
    param (
        [string]$digits
    )
    
    # Controleer of we precies 12 cijfers hebben
    if ($digits.Length -ne 12 -or $digits -notmatch '^\d{12}$') {
        throw "EAN nummer moet exact 12 cijfers bevatten"
    }
    
    $sum = 0
    for ($i = 0; $i -lt 12; $i++) {
        $digit = [int]::Parse($digits[$i].ToString())
        # Even posities x 3, oneven posities x 1
        $multiplier = if ($i % 2 -eq 0) { 1 } else { 3 }
        $sum += $digit * $multiplier
    }
    
    $checksum = (10 - ($sum % 10)) % 10
    return $checksum
}

function Validate-EANNumber {
    param (
        [string]$number
    )
    
    # Verwijder spaties en streepjes
    $number = $number -replace '[-\s]', ''
    
    # Controleer of het nummer alleen cijfers bevat
    if ($number -notmatch '^\d+$') {
        Write-Host "Fout: EAN nummer mag alleen cijfers bevatten" -ForegroundColor Red
        return $null
    }
    
    # Als het nummer 13 cijfers heeft, controleer het controlegetal
    if ($number.Length -eq 13) {
        $expectedChecksum = [int]::Parse($number[12].ToString())
        $calculatedChecksum = Calculate-EANChecksum($number.Substring(0, 12))
        
        if ($expectedChecksum -ne $calculatedChecksum) {
            Write-Host "Waarschuwing: Ongeldig controlegetal. Wordt gecorrigeerd." -ForegroundColor Yellow
            return $number.Substring(0, 12)
        }
        return $number.Substring(0, 12)
    }
    
    # Als het nummer 12 cijfers heeft, gebruik het direct
    if ($number.Length -eq 12) {
        return $number
    }
    
    Write-Host "Fout: EAN nummer moet 12 of 13 cijfers bevatten" -ForegroundColor Red
    return $null
}

try {
    Clear-Host
    Write-Host "EAN-13 Barcode Generator" -ForegroundColor Cyan
    Write-Host "======================" -ForegroundColor Cyan
    
    # Vraag om EAN nummer
    $eanInput = Read-Host "Voer een EAN nummer in (12 of 13 cijfers)"
    
    # Valideer en corrigeer het nummer
    $eanNumber = Validate-EANNumber $eanInput
    if (-not $eanNumber) {
        exit
    }
    
    # Bereken het controlegetal
    $checksum = Calculate-EANChecksum $eanNumber
    $fullEAN = "$eanNumber$checksum"
    
    Write-Host "`nGevalideerd EAN-13 nummer: $fullEAN" -ForegroundColor Green
    
    # Genereer de barcode met Barcode Generator API
    Write-Host "Barcode aan het genereren..."
    
    # Gebruik een specifieke barcode API voor EAN-13
    $barcodeUrl = "https://barcodeapi.org/api/ean13/$fullEAN"
    
    # Download de barcode
    $barcodePath = Join-Path ([Environment]::GetFolderPath('Desktop')) "EANBarcode.png"
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($barcodeUrl, $barcodePath)
    
    Write-Host "Barcode gegenereerd en opgeslagen als: $barcodePath" -ForegroundColor Green
    Write-Host "Je kunt deze barcode nu gebruiken voor je product" -ForegroundColor Cyan
    
} catch {
    Write-Host "Er is een fout opgetreden: $($_.Exception.Message)" -ForegroundColor Red
}
