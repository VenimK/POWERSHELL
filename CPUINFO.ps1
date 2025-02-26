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
# Controleer of het script met administrator rechten draait
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Temperatuurinformatie is niet beschikbaar - Administrator rechten zijn vereist" -ForegroundColor Yellow
} else {
    Try {
        Get-CimInstance MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" | 
        Select-Object @{Name="Temperatuur (°C)";Expression={[math]::Round(($_.CurrentTemperature - 2732) / 10.0, 2)}} | 
        Format-Table -AutoSize
    } Catch {
        Write-Host "Temperatuurinformatie kon niet worden opgehaald" -ForegroundColor Yellow
    }
}

# Extra processor details van ComputerInfo
Write-Host "`nAanvullende Processor Informatie:" -ForegroundColor Green
Get-ComputerInfo -Property "*processor*" | Format-List