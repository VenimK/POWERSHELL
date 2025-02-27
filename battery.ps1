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

# Functie voor temperatuur conversie
function Convert-Temperature {
    param (
        [double]$Kelvin
    )
    # Converteer Kelvin naar Celsius
    $celsius = $Kelvin - 273.15
    return [math]::Round($celsius, 1)
}

# Functie voor Batterij Status
function Get-BatterijStatus {
    try {
        $batteryStatus = Get-WmiObject -Class Win32_Battery
        $batteryDetails = Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity
        $batteryStaticData = Get-WmiObject -Namespace root\wmi -Class BatteryStaticData
        
        # Probeer temperatuur informatie te krijgen
        $temperatureInfo = Get-WmiObject -Namespace "root\wmi" -Class MSAcpi_ThermalZoneTemperature -ErrorAction SilentlyContinue
        
        if ($batteryStatus) {
            Write-Host "`nBatterij Informatie:" -ForegroundColor Cyan
            Write-Host "-------------------"
            
            # Basis informatie
            Write-Host "Status: $($batteryStatus.Status)"
            Write-Host "Resterend Percentage: $($batteryStatus.EstimatedChargeRemaining)%"
            Write-Host "Batterij Gezondheid: $(if ($batteryStatus.EstimatedChargeRemaining -gt 50) {'Goed'} else {'Redelijk'})"
            
            # Temperatuur informatie
            if ($temperatureInfo) {
                Write-Host "`nTemperatuur:" -ForegroundColor Cyan
                Write-Host "------------"
                foreach ($zone in $temperatureInfo) {
                    $tempCelsius = Convert-Temperature -Kelvin ($zone.CurrentTemperature / 10)
                    $tempStatus = switch ($true) {
                        ($tempCelsius -gt 50) { "KRITIEK" }
                        ($tempCelsius -gt 45) { "HOOG" }
                        ($tempCelsius -gt 40) { "VERHOOGD" }
                        default { "NORMAAL" }
                    }
                    $tempColor = switch ($tempStatus) {
                        "KRITIEK" { "Red" }
                        "HOOG" { "Yellow" }
                        "VERHOOGD" { "DarkYellow" }
                        default { "Green" }
                    }
                    Write-Host "Batterij Temperatuur: $tempCelsius°C" -NoNewline
                    Write-Host " [$tempStatus]" -ForegroundColor $tempColor
                }
            }
            
            # Gedetailleerde informatie
            Write-Host "`nGedetailleerde Specificaties:" -ForegroundColor Cyan
            Write-Host "-------------------------"
            Write-Host "Naam: $($batteryStatus.Name)"
            Write-Host "Beschrijving: $($batteryStatus.Description)"
            Write-Host "Type Batterij: $($batteryStaticData.Chemistry)" -ForegroundColor Yellow
            
            # Vertaal batterij type
            $batteryType = switch ($batteryStaticData.Chemistry) {
                1 { "Andere" }
                2 { "Onbekend" }
                3 { "Lood Zuur" }
                4 { "Nikkel Cadmium" }
                5 { "Nikkel Metaal Hydride" }
                6 { "Lithium Ion" }
                7 { "Zink Lucht" }
                8 { "Lithium Polymeer" }
                Default { "Niet gespecificeerd" }
            }
            Write-Host "Batterij Technologie: $batteryType" -ForegroundColor Yellow
            
            # Capaciteit informatie
            if ($batteryDetails) {
                $designedCapacity = $batteryStaticData.DesignedCapacity
                $fullChargeCapacity = $batteryDetails.FullChargedCapacity
                $capacityPercentage = [math]::Round(($fullChargeCapacity / $designedCapacity) * 100, 2)
                
                Write-Host "`nCapaciteit Informatie:" -ForegroundColor Cyan
                Write-Host "---------------------"
                Write-Host "Ontwerp Capaciteit: $designedCapacity mWh"
                Write-Host "Huidige Max Capaciteit: $fullChargeCapacity mWh"
                Write-Host "Batterij Slijtage: $($100 - $capacityPercentage)%" -ForegroundColor $(if ($capacityPercentage -lt 70) {'Red'} elseif ($capacityPercentage -lt 85) {'Yellow'} else {'Green'})
            }
            
            # Voltage en stroomsterkte
            Write-Host "`nSpanning en Stroom:" -ForegroundColor Cyan
            Write-Host "-----------------"
            Write-Host "Voltage: $($batteryStatus.DesignVoltage) mV"
            
            # Geschatte resterende tijd
            if ($batteryStatus.EstimatedRunTime -ne 71582788) {
                $remainingTime = $batteryStatus.EstimatedRunTime
                $hours = [math]::Floor($remainingTime / 60)
                $minutes = $remainingTime % 60
                Write-Host "`nGeschatte Resterende Tijd: $hours uur en $minutes minuten"
            }
            
            # Aanbevelingen op basis van temperatuur
            if ($temperatureInfo) {
                $avgTemp = ($temperatureInfo | Measure-Object -Property CurrentTemperature -Average).Average / 10
                $avgTempCelsius = Convert-Temperature -Kelvin $avgTemp
                
                Write-Host "`nAanbevelingen:" -ForegroundColor Cyan
                Write-Host "--------------"
                if ($avgTempCelsius -gt 50) {
                    Write-Host "WAARSCHUWING: Batterij temperatuur is te hoog!" -ForegroundColor Red
                    Write-Host "- Sluit zware applicaties"
                    Write-Host "- Controleer ventilatie"
                    Write-Host "- Laat laptop afkoelen"
                }
                elseif ($avgTempCelsius -gt 45) {
                    Write-Host "Let op: Batterij temperatuur is verhoogd" -ForegroundColor Yellow
                    Write-Host "- Overweeg minder zware applicaties te gebruiken"
                    Write-Host "- Zorg voor goede ventilatie"
                }
                elseif ($avgTempCelsius -gt 40) {
                    Write-Host "Tip: Batterij temperatuur is aan de hoge kant" -ForegroundColor DarkYellow
                    Write-Host "- Houd de temperatuur in de gaten"
                }
                else {
                    Write-Host "Batterij temperatuur is normaal" -ForegroundColor Green
                }
            }
            
        } else {
            Write-Host "Geen batterij gedetecteerd. Dit is waarschijnlijk een desktop PC." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Fout bij het ophalen van batterij informatie: $($_.Exception.Message)" -ForegroundColor Red
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