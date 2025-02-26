# Controleer of het script met administrator rechten draait
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Als het script niet met admin rechten draait, herstart het met verhoogde rechten
if (-not $isAdmin) {
    Write-Host "Script wordt herstart met administrator rechten..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Zet de output encoding naar UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Definieer het graden symbool
$degreeSymbol = [char]176

# Gedetailleerde CPU-informatie ophalen
Write-Host "`nGedetailleerde CPU-informatie:" -ForegroundColor Green
Get-CimInstance -ClassName Win32_Processor | Select-Object @{Name="CPU Naam";Expression={$_.Name}},
    @{Name="Fabrikant";Expression={$_.Manufacturer}},
    @{Name="Beschrijving";Expression={$_.Description}},
    @{Name="Kernen";Expression={$_.NumberOfCores}},
    @{Name="Logische Processors";Expression={$_.NumberOfLogicalProcessors}},
    @{Name="Max Kloksnelheid (GHz)";Expression={[math]::Round($_.MaxClockSpeed/1000, 2)}},
    @{Name="Huidige Kloksnelheid (GHz)";Expression={[math]::Round($_.CurrentClockSpeed/1000, 2)}},
    @{Name="Socket";Expression={$_.SocketDesignation}},
    @{Name="L2 Cache (KB)";Expression={$_.L2CacheSize}},
    @{Name="L3 Cache (KB)";Expression={$_.L3CacheSize}},
    @{Name="Architectuur";Expression={$_.Architecture}},
    @{Name="Status";Expression={$_.Status}} | Format-Table -AutoSize

# CPU belasting informatie ophalen
Write-Host "`nCPU Belastingsinformatie:" -ForegroundColor Green
Get-CimInstance -ClassName Win32_Processor | Select-Object @{Name="CPU Belasting %";Expression={$_.LoadPercentage}},
    @{Name="Huidige Voltage";Expression={$_.CurrentVoltage}},
    @{Name="Energiebeheer Ondersteund";Expression={$_.PowerManagementSupported}} | Format-Table -AutoSize

# CPU temperatuur ophalen indien beschikbaar
Write-Host "`nCPU Temperatuurinformatie:" -ForegroundColor Green
Try {
    $tempOutput = Get-CimInstance MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" | 
        Select-Object @{Name="Temperatuur ($degreeSymbol C)";Expression={[math]::Round(($_.CurrentTemperature - 2732) / 10.0, 2)}}
    $tempOutput | Format-Table -AutoSize
} Catch {
    Write-Host "Temperatuurinformatie kon niet worden opgehaald" -ForegroundColor Yellow
}

# Extra processor details van ComputerInfo
Write-Host "`nAanvullende Processor Informatie:" -ForegroundColor Green
Get-ComputerInfo -Property "*processor*" | Format-List

# Wacht op gebruiker input voordat het venster sluit
Write-Host "`nDruk op een toets om af te sluiten..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")