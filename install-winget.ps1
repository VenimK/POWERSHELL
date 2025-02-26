# Set output encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Controleer uitvoeringsbeleid op alle niveaus
Write-Host "Uitvoeringsbeleid controleren op alle niveaus..." -ForegroundColor Cyan
$policyList = Get-ExecutionPolicy -List
$policyList | Format-Table -AutoSize

# Aangezien er een beleidsoverschrijving is, slaan we wijzigingen over als het al Bypass of RemoteSigned is
$effectivePolicy = Get-ExecutionPolicy
Write-Host "Effectief Uitvoeringsbeleid: $effectivePolicy" -ForegroundColor Cyan

if ($effectivePolicy -notin @("RemoteSigned", "Bypass", "Unrestricted")) {
    try {
        Write-Host "Uitvoeringsbeleid instellen op RemoteSigned..." -ForegroundColor Yellow
        Set-ExecutionPolicy RemoteSigned -Force
        Write-Host "Uitvoeringsbeleid succesvol bijgewerkt naar RemoteSigned" -ForegroundColor Green
    } catch {
        Write-Host "Opmerking: Wijziging van uitvoeringsbeleid is geprobeerd maar kan worden overschreven door groepsbeleid." -ForegroundColor Yellow
        Write-Host "Huidig effectief beleid ($effectivePolicy) zou scripts nog steeds moeten kunnen uitvoeren." -ForegroundColor Yellow
    }
}

# Controleer of we met beheerderrechten draaien
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Start dit script als Administrator!"
    Exit
}

Write-Host "Start Winget installatie/reparatie proces..." -ForegroundColor Cyan

# Controleer of winget al is geinstalleerd
$hasWinget = Get-AppxPackage -Name Microsoft.DesktopAppInstaller

if ($hasWinget) {
    Write-Host "Bestaande Winget installatie verwijderen..." -ForegroundColor Yellow
    Get-AppxPackage *Microsoft.DesktopAppInstaller* | Remove-AppxPackage
    Start-Sleep -Seconds 2
}

# Download en installeer de nieuwste versie van winget
Write-Host "Nieuwste Winget installer downloaden..." -ForegroundColor Cyan
$progressPreference = 'silentlyContinue'
$wingetUrl = "https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
$installerPath = Join-Path $env:TEMP "Microsoft.DesktopAppInstaller.msixbundle"

try {
    Invoke-WebRequest -Uri $wingetUrl -OutFile $installerPath
    Write-Host "Winget installeren..." -ForegroundColor Cyan
    Add-AppxPackage $installerPath
} catch {
    Write-Error "Downloaden of installeren van Winget mislukt: $_"
    Exit 1
} finally {
    # Ruim het installatiebestand op
    if (Test-Path $installerPath) {
        Remove-Item $installerPath
    }
}

# Ververs omgevingsvariabelen
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

# Verifieer installatie
Write-Host "Winget installatie verifieren..." -ForegroundColor Cyan
try {
    $wingetVersion = winget --version
    Write-Host "Winget succesvol geinstalleerd!" -ForegroundColor Green
    Write-Host "Versie: $wingetVersion" -ForegroundColor Green
} catch {
    Write-Error "Winget installatie verificatie mislukt. Probeer het script opnieuw uit te voeren of herstart uw systeem."
    Exit 1
}

Write-Host "`nInstallatie voltooid! U moet mogelijk uw terminal opnieuw opstarten om de wijzigingen door te voeren." -ForegroundColor Cyan
