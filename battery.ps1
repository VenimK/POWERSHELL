# Batterij Status en Driver Update Script
# Dit script genereert een batterijrapport en handelt driver updates af

# Check en vraag om administrator rechten
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    try {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }
    catch {
        Write-Host "Fout bij het verkrijgen van administrator rechten: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Start het script opnieuw als administrator." -ForegroundColor Yellow
        pause
        exit
    }
}

# Functie voor Batterij Status
function Get-BatterijStatus {
    try {
        $batteryStatus = Get-WmiObject -Class Win32_Battery
        if ($batteryStatus) {
            Write-Host "`nBatterij Informatie:" -ForegroundColor Cyan
            Write-Host "-------------------"
            Write-Host "Batterij Status: $($batteryStatus.Status)"
            Write-Host "Resterend Percentage: $($batteryStatus.EstimatedChargeRemaining)%"
            Write-Host "Batterij Gezondheid: $(if ($batteryStatus.EstimatedChargeRemaining -gt 50) {'Goed'} else {'Redelijk'})"
        } else {
            Write-Host "Geen batterij gedetecteerd. Dit is waarschijnlijk een desktop PC." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Fout bij ophalen batterij informatie: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Functie voor Batterij Rapport
function Get-BatterijRapport {
    try {
        $reportPath = "$env:USERPROFILE\batterij-rapport.html"
        Write-Host "`nGenereren van batterijrapport..." -ForegroundColor Cyan
        powercfg /batteryreport /output $reportPath
        if (Test-Path $reportPath) {
            Write-Host "Batterijrapport succesvol gegenereerd op: $reportPath" -ForegroundColor Green
            Start-Process $reportPath
        }
    } catch {
        Write-Host "Fout bij genereren batterijrapport: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Functie voor Driver Updates
function Update-Drivers {
    try {
        Write-Host "`nStarten van driver update proces..." -ForegroundColor Cyan
        
        $UpdateSvc = New-Object -ComObject Microsoft.Update.ServiceManager
        $UpdateSvc.AddService2("7971f918-a847-4430-9279-4a52d1efe18d",7,"")
        
        $Session = New-Object -ComObject Microsoft.Update.Session
        $Searcher = $Session.CreateUpdateSearcher()
        
        $Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d'
        $Searcher.SearchScope = 1  # Alleen Machine
        $Searcher.ServerSelection = 3  # Derde Partij
        
        $Criteria = "IsInstalled=0 and Type='Driver'"
        Write-Host 'Zoeken naar driver updates...' -ForegroundColor Cyan
        
        $SearchResult = $Searcher.Search($Criteria)
        $Updates = $SearchResult.Updates

        if ($Updates.Count -eq 0) {
            Write-Host "Geen nieuwe driver updates gevonden." -ForegroundColor Green
            return
        }

        # Toon Beschikbare Drivers
        Write-Host "`nBeschikbare Driver Updates:" -ForegroundColor Cyan
        $Updates | Select-Object Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | Format-List

        # Download Updates
        $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
        $Updates | ForEach-Object { $UpdatesToDownload.Add($_) | Out-Null }
        
        Write-Host 'Downloaden van drivers...' -ForegroundColor Cyan
        $Downloader = $Session.CreateUpdateDownloader()
        $Downloader.Updates = $UpdatesToDownload
        $Downloader.Download()

        # Installeer Updates
        $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        $Updates | Where-Object { $_.IsDownloaded } | ForEach-Object { 
            $UpdatesToInstall.Add($_) | Out-Null 
        }

        Write-Host 'Installeren van drivers...' -ForegroundColor Cyan
        $Installer = $Session.CreateUpdateInstaller()
        $Installer.Updates = $UpdatesToInstall
        $InstallationResult = $Installer.Install()

        if ($InstallationResult.RebootRequired) {
            Write-Host 'Herstart benodigd! Herstart uw computer a.u.b.' -ForegroundColor Yellow
        } else {
            Write-Host 'Driver installatie succesvol afgerond.' -ForegroundColor Green
        }

    } catch {
        Write-Host "Fout tijdens driver update proces: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        # Opruimen
        try {
            $UpdateSvc.Services | Where-Object { 
                $_.IsDefaultAUService -eq $false -and 
                $_.ServiceID -eq "7971f918-a847-4430-9279-4a52d1efe18d" 
            } | ForEach-Object { 
                $UpdateSvc.RemoveService($_.ServiceID)
            }
        } catch {
            Write-Host "Waarschuwing bij opruimen: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# Hoofduitvoering
Clear-Host
Write-Host "Batterij en Driver Update Hulpprogramma" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green

# Voer alle functies uit
Get-BatterijStatus
Get-BatterijRapport
Update-Drivers