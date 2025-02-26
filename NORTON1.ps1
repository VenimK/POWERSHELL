# Script configuratie
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

Write-Host "Norton 360 installatie starten..." -ForegroundColor Yellow

try {
    Write-Host "Installeren van Norton 360..." -ForegroundColor Cyan
    # Install Norton 360 silently
    winget install -e --id XPFNZKWN35KD6Z --silent --accept-source-agreements --accept-package-agreements
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Norton 360 succesvol geinstalleerd!" -ForegroundColor Green
        Write-Host "Installatie voltooid. U kunt Norton later activeren via de Windows Start Menu." -ForegroundColor Yellow
    } else {
        throw "Installatie via winget is mislukt."
    }
} catch {
    Write-Host "Er is een fout opgetreden:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "`nProbeer Norton 360 handmatig te downloaden van:" -ForegroundColor Yellow
    Write-Host "https://nl.norton.com/products" -ForegroundColor Cyan
}