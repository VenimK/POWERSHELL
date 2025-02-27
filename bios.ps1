# Gedetailleerde Systeem en BIOS Informatie Script
param (
    [switch]$ExportToHTML,
    [string]$ExportPath
)

# Controleer en vraag om administrator rechten
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    try {
        # Start een nieuwe PowerShell instantie als administrator met dezelfde parameters
        $arguments = ""
        if ($ExportToHTML) { $arguments += " -ExportToHTML" }
        if ($ExportPath) { $arguments += " -ExportPath '$ExportPath'" }
        
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"$arguments" -Verb RunAs
        exit
    }
    catch {
        Write-Host "Fout bij het verkrijgen van administrator rechten: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Start het script opnieuw als administrator." -ForegroundColor Yellow
        pause
        exit
    }
}

function Write-ColoredHeader {
    param(
        [string]$Text
    )
    Write-Host "`n$Text" -ForegroundColor Cyan
    Write-Host ("-" * $Text.Length) -ForegroundColor Cyan
}

function Get-FormattedSize {
    param (
        [uint64]$Bytes
    )
    $sizes = "Bytes", "KB", "MB", "GB", "TB"
    $order = 0
    while ($Bytes -ge 1024 -and $order -lt $sizes.Count) {
        $order++
        $Bytes = $Bytes/1024
    }
    return "{0:N2} {1}" -f $Bytes, $sizes[$order]
}

function Controleer-BIOSUpdates {
    param (
        [string]$Fabrikant,
        [string]$Model,
        [string]$HuidigeBIOSVersie
    )

    try {
        Write-ColoredHeader "BIOS UPDATE CONTROLE"
        
        switch -Wildcard ($Fabrikant.ToLower()) {
            "*asus*" {
                $modelFormatted = $Model.Replace(" ", "-")
                $supportUrl = "https://www.asus.com/supportonly/$modelFormatted/HelpDesk_BIOS/"
                Write-Host "BIOS Ondersteuningspagina: $supportUrl" -ForegroundColor Cyan
                Write-Host "Huidige BIOS Versie: $HuidigeBIOSVersie" -ForegroundColor Green
                Write-Host "`nGa naar de bovenstaande URL om te controleren op BIOS updates." -ForegroundColor Yellow
                Write-Host "Let op: Download en installeer BIOS updates alleen van de officiële website." -ForegroundColor Yellow
            }
            "*dell*" {
                $serviceTag = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
                $supportUrl = "https://www.dell.com/support/home/nl-nl/product-support/servicetag/$serviceTag/drivers"
                Write-Host "BIOS Ondersteuningspagina: $supportUrl" -ForegroundColor Cyan
                Write-Host "Huidige BIOS Versie: $HuidigeBIOSVersie" -ForegroundColor Green
                Write-Host "`nGa naar de bovenstaande URL om te controleren op BIOS updates." -ForegroundColor Yellow
            }
            "*hp*" {
                $supportUrl = "https://support.hp.com/nl-nl/drivers"
                Write-Host "BIOS Ondersteuningspagina: $supportUrl" -ForegroundColor Cyan
                Write-Host "Huidige BIOS Versie: $HuidigeBIOSVersie" -ForegroundColor Green
                Write-Host "`nGa naar de bovenstaande URL en voer uw productmodel ($Model) in." -ForegroundColor Yellow
            }
            "*lenovo*" {
                $supportUrl = "https://pcsupport.lenovo.com/nl/nl/downloads/ds012808"
                Write-Host "BIOS Ondersteuningspagina: $supportUrl" -ForegroundColor Cyan
                Write-Host "Huidige BIOS Versie: $HuidigeBIOSVersie" -ForegroundColor Green
                Write-Host "`nGa naar de bovenstaande URL en voer uw productmodel ($Model) in." -ForegroundColor Yellow
            }
            default {
                Write-Host "Automatische BIOS update controle niet beschikbaar voor $Fabrikant" -ForegroundColor Yellow
                Write-Host "Bezoek de website van uw fabrikant voor BIOS updates." -ForegroundColor Yellow
            }
        }
        
        Write-Host "`nWaarschuwing:" -ForegroundColor Red
        Write-Host "- Maak altijd een back-up van uw gegevens voor het updaten van de BIOS" -ForegroundColor Red
        Write-Host "- Zorg voor een stabiele stroomvoorziening tijdens de BIOS update" -ForegroundColor Red
        Write-Host "- Onderbreek het update proces NOOIT" -ForegroundColor Red
        Write-Host "- Sluit alle andere programma's tijdens de update" -ForegroundColor Red
    }
    catch {
        Write-Host "Fout bij het controleren van BIOS updates: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Systeem Informatie
Write-ColoredHeader "SYSTEEM INFORMATIE"
$computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$bios = Get-CimInstance -ClassName Win32_BIOS
$processor = Get-CimInstance -ClassName Win32_Processor
$videoCard = Get-CimInstance -ClassName Win32_VideoController
$networkAdapters = Get-CimInstance -ClassName Win32_NetworkAdapter | Where-Object { $_.PhysicalAdapter }
$diskDrives = Get-CimInstance -ClassName Win32_DiskDrive

Write-Host "Computernaam: " -NoNewline; Write-Host $computerSystem.Name -ForegroundColor Green
Write-Host "Fabrikant: " -NoNewline; Write-Host $computerSystem.Manufacturer -ForegroundColor Green
Write-Host "Model: " -NoNewline; Write-Host $computerSystem.Model -ForegroundColor Green
Write-Host "Serienummer: " -NoNewline; Write-Host $bios.SerialNumber -ForegroundColor Green

# Processor Details
Write-ColoredHeader "PROCESSOR INFORMATIE"
Write-Host "CPU: " -NoNewline; Write-Host $processor.Name.Trim() -ForegroundColor Green
Write-Host "Cores: " -NoNewline; Write-Host $processor.NumberOfCores -ForegroundColor Green
Write-Host "Logische Processors: " -NoNewline; Write-Host $processor.NumberOfLogicalProcessors -ForegroundColor Green
Write-Host "Max Kloksnelheid: " -NoNewline; Write-Host "$($processor.MaxClockSpeed) MHz" -ForegroundColor Green
Write-Host "Socket: " -NoNewline; Write-Host $processor.SocketDesignation -ForegroundColor Green
Write-Host "Virtualisatie Ingeschakeld: " -NoNewline
$virtualization = Get-CimInstance -ClassName Win32_Processor | Select-Object -ExpandProperty VirtualizationFirmwareEnabled
Write-Host $virtualization -ForegroundColor $(if ($virtualization) {"Green"} else {"Yellow"})

# Geheugen Details
Write-ColoredHeader "GEHEUGEN INFORMATIE"
$physicalMemory = Get-CimInstance -ClassName Win32_PhysicalMemory
$totalMemory = ($physicalMemory | Measure-Object -Property Capacity -Sum).Sum
Write-Host "Totaal Geïnstalleerd Geheugen: " -NoNewline; Write-Host "$(Get-FormattedSize -Bytes $totalMemory)" -ForegroundColor Green
Write-Host "`nGeheugen Modules:"
$physicalMemory | ForEach-Object {
    Write-Host "- Slot: $($_.DeviceLocator)" -ForegroundColor Yellow
    Write-Host "  Capaciteit: $(Get-FormattedSize -Bytes $_.Capacity)"
    Write-Host "  Snelheid: $($_.Speed) MHz"
    Write-Host "  Fabrikant: $($_.Manufacturer)"
    Write-Host "  Type: $($_.MemoryType)"
}

# Grafische Kaart
Write-ColoredHeader "GRAFISCHE KAART INFORMATIE"
foreach ($gpu in $videoCard) {
    Write-Host "GPU: " -NoNewline; Write-Host $gpu.Name -ForegroundColor Green
    Write-Host "Resolutie: " -NoNewline; Write-Host "$($gpu.CurrentHorizontalResolution) x $($gpu.CurrentVerticalResolution)" -ForegroundColor Green
    Write-Host "Driver Versie: " -NoNewline; Write-Host $gpu.DriverVersion -ForegroundColor Green
    Write-Host "Video Geheugen: " -NoNewline; Write-Host "$(Get-FormattedSize -Bytes $gpu.AdapterRAM)" -ForegroundColor Green
}

# Opslag Informatie
Write-ColoredHeader "OPSLAG INFORMATIE"
foreach ($disk in $diskDrives) {
    Write-Host "`nSchijf: $($disk.DeviceID)" -ForegroundColor Yellow
    Write-Host "Model: " -NoNewline; Write-Host $disk.Model -ForegroundColor Green
    Write-Host "Grootte: " -NoNewline; Write-Host "$(Get-FormattedSize -Bytes $disk.Size)" -ForegroundColor Green
    Write-Host "Interface: " -NoNewline; Write-Host $disk.InterfaceType -ForegroundColor Green
}

# BIOS Informatie
Write-ColoredHeader "BIOS INFORMATIE"
Write-Host "BIOS Versie: " -NoNewline; Write-Host $bios.Version -ForegroundColor Green
Write-Host "BIOS Fabrikant: " -NoNewline; Write-Host $bios.Manufacturer -ForegroundColor Green
Write-Host "BIOS Release Datum: " -NoNewline; Write-Host $bios.ReleaseDate.ToString("dd-MM-yyyy") -ForegroundColor Green
Write-Host "Serienummer: " -NoNewline; Write-Host $bios.SerialNumber -ForegroundColor Green
Write-Host "SMBIOS Versie: " -NoNewline; Write-Host $bios.SMBIOSBIOSVersion -ForegroundColor Green

Controleer-BIOSUpdates -Fabrikant $computerSystem.Manufacturer -Model $computerSystem.Model -HuidigeBIOSVersie $bios.Version

# Moederbord Informatie
Write-ColoredHeader "MOEDERBORD INFORMATIE"
$baseBoard = Get-CimInstance -ClassName Win32_BaseBoard
Write-Host "Fabrikant: " -NoNewline; Write-Host $baseBoard.Manufacturer -ForegroundColor Green
Write-Host "Model: " -NoNewline; Write-Host $baseBoard.Product -ForegroundColor Green
Write-Host "Serienummer: " -NoNewline; Write-Host $baseBoard.SerialNumber -ForegroundColor Green
Write-Host "Versie: " -NoNewline; Write-Host $baseBoard.Version -ForegroundColor Green

# Netwerk Adapters
Write-ColoredHeader "NETWERK ADAPTERS"
foreach ($adapter in $networkAdapters) {
    Write-Host "`nAdapter: $($adapter.Name)" -ForegroundColor Yellow
    Write-Host "MAC Adres: " -NoNewline; Write-Host $adapter.MACAddress -ForegroundColor Green
    Write-Host "Adapter Type: " -NoNewline; Write-Host $adapter.AdapterType -ForegroundColor Green
    Write-Host "Snelheid: " -NoNewline; Write-Host "$([math]::Round($adapter.Speed/1000000, 2)) Mbps" -ForegroundColor Green
}

# Boot Configuratie
Write-ColoredHeader "BOOT CONFIGURATIE"
$bootConfig = Get-CimInstance -ClassName Win32_BootConfiguration
Write-Host "Boot Directory: " -NoNewline; Write-Host $bootConfig.BootDirectory -ForegroundColor Green
Write-Host "Configuratie Pad: " -NoNewline; Write-Host $bootConfig.ConfigurationPath -ForegroundColor Green

# UEFI/Legacy Status
Write-ColoredHeader "BOOT MODUS"
$secureBootStatus = $null
try {
    $secureBootStatus = Confirm-SecureBootUEFI
} catch {
    $secureBootStatus = "Niet beschikbaar"
}

$bootMode = if ((Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\SecureBoot\State" -ErrorAction SilentlyContinue).UEFISecureBootEnabled) {
    "UEFI met Secure Boot"
} elseif ((Get-CimInstance -Class Win32_ComputerSystem).BootupState -match "Normal boot") {
    "UEFI zonder Secure Boot"
} else {
    "Legacy BIOS"
}

Write-Host "Boot Type: " -NoNewline; Write-Host $bootMode -ForegroundColor Green
Write-Host "Secure Boot Status: " -NoNewline; Write-Host $secureBootStatus -ForegroundColor $(if ($secureBootStatus -eq $true) {"Green"} else {"Yellow"})

# TPM Status
Write-ColoredHeader "TPM STATUS"
try {
    $tpm = Get-Tpm
    if ($tpm) {
        Write-Host "TPM Aanwezig: " -NoNewline; Write-Host "Ja" -ForegroundColor Green
        Write-Host "TPM Geactiveerd: " -NoNewline; Write-Host $tpm.TpmEnabled -ForegroundColor $(if ($tpm.TpmEnabled) {"Green"} else {"Yellow"})
        Write-Host "TPM Eigenaar: " -NoNewline; Write-Host $tpm.TpmOwned -ForegroundColor $(if ($tpm.TpmOwned) {"Green"} else {"Yellow"})
        Write-Host "TPM Versie: " -NoNewline; Write-Host $tpm.ManufacturerVersion -ForegroundColor Green
        Write-Host "TPM Fabrikant: " -NoNewline; Write-Host $tpm.ManufacturerId -ForegroundColor Green
    } else {
        Write-Host "TPM Status: Niet gevonden" -ForegroundColor Yellow
    }
} catch {
    Write-Host "TPM Status kon niet worden bepaald" -ForegroundColor Yellow
}

# Firmware Updates
Write-ColoredHeader "FIRMWARE UPDATES"
try {
    $firmwareUpdates = Get-HotFix | Where-Object { $_.Description -match "Firmware|BIOS|System" } | Sort-Object -Property InstalledOn -Descending
    if ($firmwareUpdates) {
        Write-Host "Laatste Firmware/BIOS Updates:"
        $firmwareUpdates | Select-Object -First 5 | ForEach-Object {
            Write-Host "`nUpdate ID: " -NoNewline; Write-Host $_.HotFixID -ForegroundColor Green
            Write-Host "Geïnstalleerd: " -NoNewline; Write-Host $_.InstalledOn.ToString("dd-MM-yyyy") -ForegroundColor Green
            Write-Host "Beschrijving: " -NoNewline; Write-Host $_.Description -ForegroundColor Green
        }
    } else {
        Write-Host "Geen recente firmware updates gevonden" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Kon firmware update geschiedenis niet ophalen" -ForegroundColor Yellow
}

# Aanbevelingen
Write-ColoredHeader "SYSTEEM AANBEVELINGEN"
$recommendations = @()

if ($bootMode -ne "UEFI met Secure Boot") {
    $recommendations += "- Overweeg UEFI met Secure Boot in te schakelen voor betere beveiliging"
}

if ($tpm -and -not $tpm.TpmEnabled) {
    $recommendations += "- Activeer TPM voor verbeterde systeembeveiliging"
}

if ($virtualization -eq $false) {
    $recommendations += "- Overweeg virtualisatie in te schakelen in BIOS/UEFI voor betere prestaties met virtuele machines"
}

if ($recommendations.Count -gt 0) {
    Write-Host "Aanbevolen acties:"
    $recommendations | ForEach-Object {
        Write-Host $_ -ForegroundColor Yellow
    }
} else {
    Write-Host "Geen specifieke aanbevelingen. Systeem configuratie lijkt optimaal." -ForegroundColor Green
}

# Export naar HTML als gevraagd
if ($ExportToHTML) {
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Systeem Rapport - $($computerSystem.Name)</title>
    <meta charset="UTF-8">
    <style>
        body { 
            font-family: Arial, sans-serif; 
            margin: 20px;
            background-color: #f5f6fa;
            color: #2c3e50;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        h1 { 
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
        }
        h2 { 
            color: #2c3e50; 
            border-bottom: 2px solid #3498db; 
            padding-bottom: 5px;
            margin-top: 30px;
        }
        .info-group { 
            margin: 15px 0; 
            padding: 20px; 
            background-color: white; 
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .label { 
            font-weight: bold; 
            color: #34495e;
            display: inline-block;
            width: 200px;
        }
        .value { 
            color: #27ae60;
        }
        .warning { 
            color: #e67e22;
        }
        .error { 
            color: #c0392b;
        }
        .success {
            color: #27ae60;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
        }
        th, td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #f8f9fa;
        }
        .timestamp {
            color: #7f8c8d;
            font-size: 0.9em;
            margin-bottom: 20px;
        }
        .recommendations {
            background-color: #fff3cd;
            border-left: 5px solid #ffc107;
            padding: 15px;
            margin: 20px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Systeem Rapport - $($computerSystem.Name)</h1>
        <div class="timestamp">Gegenereerd op: $(Get-Date -Format "dd-MM-yyyy HH:mm:ss")</div>

        <div class="info-group">
            <h2>Systeem Informatie</h2>
            <p><span class="label">Computer:</span> <span class="value">$($computerSystem.Name)</span></p>
            <p><span class="label">Fabrikant:</span> <span class="value">$($computerSystem.Manufacturer)</span></p>
            <p><span class="label">Model:</span> <span class="value">$($computerSystem.Model)</span></p>
            <p><span class="label">Serienummer:</span> <span class="value">$($bios.SerialNumber)</span></p>
            <p><span class="label">Besturingssysteem:</span> <span class="value">$($os.Caption)</span></p>
            <p><span class="label">OS Versie:</span> <span class="value">$($os.Version)</span></p>
            <p><span class="label">OS Build:</span> <span class="value">$($os.BuildNumber)</span></p>
        </div>

        <div class="info-group">
            <h2>Processor Informatie</h2>
            <p><span class="label">CPU:</span> <span class="value">$($processor.Name.Trim())</span></p>
            <p><span class="label">Cores:</span> <span class="value">$($processor.NumberOfCores)</span></p>
            <p><span class="label">Logische Processors:</span> <span class="value">$($processor.NumberOfLogicalProcessors)</span></p>
            <p><span class="label">Max Kloksnelheid:</span> <span class="value">$($processor.MaxClockSpeed) MHz</span></p>
            <p><span class="label">Socket:</span> <span class="value">$($processor.SocketDesignation)</span></p>
            <p><span class="label">Virtualisatie:</span> <span class="$(if($virtualization){'success'}else{'warning'})">$(if($virtualization){'Ingeschakeld'}else{'Uitgeschakeld'})</span></p>
        </div>

        <div class="info-group">
            <h2>Geheugen Informatie</h2>
            <p><span class="label">Totaal Geheugen:</span> <span class="value">$(Get-FormattedSize -Bytes $totalMemory)</span></p>
            <table>
                <tr>
                    <th>Slot</th>
                    <th>Capaciteit</th>
                    <th>Snelheid</th>
                    <th>Fabrikant</th>
                    <th>Type</th>
                </tr>
                $(
                    $physicalMemory | ForEach-Object {
                        "<tr>
                            <td>$($_.DeviceLocator)</td>
                            <td>$(Get-FormattedSize -Bytes $_.Capacity)</td>
                            <td>$($_.Speed) MHz</td>
                            <td>$($_.Manufacturer)</td>
                            <td>$($_.MemoryType)</td>
                        </tr>"
                    }
                )
            </table>
        </div>

        <div class="info-group">
            <h2>Grafische Kaart Informatie</h2>
            $(
                $videoCard | ForEach-Object {
                    "<p><span class='label'>GPU:</span> <span class='value'>$($_.Name)</span></p>
                    <p><span class='label'>Resolutie:</span> <span class='value'>$($_.CurrentHorizontalResolution) x $($_.CurrentVerticalResolution)</span></p>
                    <p><span class='label'>Driver Versie:</span> <span class='value'>$($_.DriverVersion)</span></p>
                    <p><span class='label'>Video Geheugen:</span> <span class='value'>$(Get-FormattedSize -Bytes $_.AdapterRAM)</span></p>"
                }
            )
        </div>

        <div class="info-group">
            <h2>Opslag Informatie</h2>
            <table>
                <tr>
                    <th>Schijf</th>
                    <th>Model</th>
                    <th>Grootte</th>
                    <th>Interface</th>
                </tr>
                $(
                    $diskDrives | ForEach-Object {
                        "<tr>
                            <td>$($_.DeviceID)</td>
                            <td>$($_.Model)</td>
                            <td>$(Get-FormattedSize -Bytes $_.Size)</td>
                            <td>$($_.InterfaceType)</td>
                        </tr>"
                    }
                )
            </table>
        </div>

        <div class="info-group">
            <h2>BIOS Informatie</h2>
            <p><span class="label">BIOS Versie:</span> <span class="value">$($bios.Version)</span></p>
            <p><span class="label">BIOS Fabrikant:</span> <span class="value">$($bios.Manufacturer)</span></p>
            <p><span class="label">Release Datum:</span> <span class="value">$($bios.ReleaseDate.ToString("dd-MM-yyyy"))</span></p>
            <p><span class="label">SMBIOS Versie:</span> <span class="value">$($bios.SMBIOSBIOSVersion)</span></p>
        </div>

        <div class="info-group">
            <h2>BIOS Update Informatie</h2>
            <p><span class="label">Ondersteuningspagina:</span> <span class="value"><a href="$(
                switch -Wildcard ($computerSystem.Manufacturer.ToLower()) {
                    "*asus*" { "https://www.asus.com/supportonly/$($computerSystem.Model.Replace(' ', '-'))/HelpDesk_BIOS/" }
                    "*dell*" { "https://www.dell.com/support/home/nl-nl/product-support/servicetag/$($bios.SerialNumber)/drivers" }
                    "*hp*" { "https://support.hp.com/nl-nl/drivers" }
                    "*lenovo*" { "https://pcsupport.lenovo.com/nl/nl/downloads/ds012808" }
                    default { "#" }
                }
            )" target="_blank">BIOS Downloads</a></span></p>
            <div class="warning" style="margin-top: 15px; padding: 10px; border-left: 4px solid #e74c3c;">
                <strong>Waarschuwingen voor BIOS Updates:</strong>
                <ul>
                    <li>Maak altijd een back-up van uw gegevens voor het updaten van de BIOS</li>
                    <li>Zorg voor een stabiele stroomvoorziening tijdens de BIOS update</li>
                    <li>Onderbreek het update proces NOOIT</li>
                    <li>Download BIOS updates alleen van de officiële website van de fabrikant</li>
                    <li>Sluit alle andere programma's tijdens de update</li>
                </ul>
            </div>
        </div>

        <div class="info-group">
            <h2>Moederbord Informatie</h2>
            <p><span class="label">Fabrikant:</span> <span class="value">$($baseBoard.Manufacturer)</span></p>
            <p><span class="label">Model:</span> <span class="value">$($baseBoard.Product)</span></p>
            <p><span class="label">Versie:</span> <span class="value">$($baseBoard.Version)</span></p>
        </div>

        <div class="info-group">
            <h2>Netwerk Adapters</h2>
            <table>
                <tr>
                    <th>Adapter</th>
                    <th>MAC Adres</th>
                    <th>Type</th>
                    <th>Snelheid</th>
                </tr>
                $(
                    $networkAdapters | ForEach-Object {
                        "<tr>
                            <td>$($_.Name)</td>
                            <td>$($_.MACAddress)</td>
                            <td>$($_.AdapterType)</td>
                            <td>$([math]::Round($_.Speed/1000000, 2)) Mbps</td>
                        </tr>"
                    }
                )
            </table>
        </div>

        <div class="info-group">
            <h2>Boot Configuratie</h2>
            <p><span class="label">Boot Type:</span> <span class="value">$bootMode</span></p>
            <p><span class="label">Secure Boot:</span> <span class="$(if($secureBootStatus -eq $true){'success'}else{'warning'})">$secureBootStatus</span></p>
            <p><span class="label">Boot Directory:</span> <span class="value">$($bootConfig.BootDirectory)</span></p>
        </div>

        <div class="info-group">
            <h2>TPM Status</h2>
            $(if ($tpm) {
                "<p><span class='label'>TPM Aanwezig:</span> <span class='success'>Ja</span></p>
                <p><span class='label'>TPM Geactiveerd:</span> <span class='$(if($tpm.TpmEnabled){'success'}else{'warning'})'>$($tpm.TpmEnabled)</span></p>
                <p><span class='label'>TPM Eigenaar:</span> <span class='$(if($tpm.TpmOwned){'success'}else{'warning'})'>$($tpm.TpmOwned)</span></p>
                <p><span class='label'>TPM Versie:</span> <span class='value'>$($tpm.ManufacturerVersion)</span></p>
                <p><span class='label'>TPM Fabrikant:</span> <span class='value'>$($tpm.ManufacturerId)</span></p>"
            } else {
                "<p><span class='label'>TPM Status:</span> <span class='warning'>Niet gevonden</span></p>"
            })
        </div>

        $(if ($firmwareUpdates) {
            "<div class='info-group'>
                <h2>Laatste Firmware Updates</h2>
                <table>
                    <tr>
                        <th>Update ID</th>
                        <th>Geïnstalleerd</th>
                        <th>Beschrijving</th>
                    </tr>
                    $(
                        $firmwareUpdates | Select-Object -First 5 | ForEach-Object {
                            "<tr>
                                <td>$($_.HotFixID)</td>
                                <td>$($_.InstalledOn.ToString('dd-MM-yyyy'))</td>
                                <td>$($_.Description)</td>
                            </tr>"
                        }
                    )
                </table>
            </div>"
        })

        $(if ($recommendations.Count -gt 0) {
            "<div class='recommendations'>
                <h2>Systeem Aanbevelingen</h2>
                <ul>
                    $(
                        $recommendations | ForEach-Object {
                            "<li>$_</li>"
                        }
                    )
                </ul>
            </div>"
        })
    </div>
</body>
</html>
"@

    $exportFilePath = if ($ExportPath) {
        Join-Path $ExportPath "SystemReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    } else {
        Join-Path $PSScriptRoot "SystemReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    }

    $htmlContent | Out-File -FilePath $exportFilePath -Encoding UTF8
    Write-Host "`nRapport geëxporteerd naar: $exportFilePath" -ForegroundColor Green
}