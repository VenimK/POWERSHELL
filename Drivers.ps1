# Check voor Administrator rechten
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

try {
    Write-Host "Windows Update Service voorbereiden..." -ForegroundColor Cyan
    
    # Start de Windows Update service als deze niet draait
    $wuauserv = Get-Service -Name "wuauserv"
    if ($wuauserv.Status -ne "Running") {
        Write-Host "Windows Update service starten..." -ForegroundColor Yellow
        Start-Service -Name "wuauserv"
        Start-Sleep -Seconds 5
    }

    # Controleer of de service nu draait
    $wuauserv.Refresh()
    if ($wuauserv.Status -ne "Running") {
        throw "Windows Update service kon niet worden gestart."
    }

    Write-Host "Zoeken naar driver updates..." -ForegroundColor Cyan
    
    # Maak Windows Update objecten aan
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    
    # Zoek naar driver updates
    $SearchResult = $UpdateSearcher.Search("Type='Driver' AND IsInstalled=0")
    
    if ($SearchResult.Updates.Count -eq 0) {
        Write-Host "Geen nieuwe driver updates gevonden." -ForegroundColor Green
    } else {
        Write-Host "`nBeschikbare Driver Updates:" -ForegroundColor Cyan
        
        # Toon beschikbare updates
        $SearchResult.Updates | ForEach-Object {
            Write-Host "`nNaam: $($_.Title)"
            Write-Host "Beschrijving: $($_.Description)"
            Write-Host "-----------------------------------"
        }
        
        # Download en installeer updates
        $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
        $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        
        $SearchResult.Updates | ForEach-Object {
            if (-not $_.IsDownloaded) {
                Write-Host "Downloaden: $($_.Title)" -ForegroundColor Yellow
                $UpdatesToDownload.Add($_) | Out-Null
            } else {
                $UpdatesToInstall.Add($_) | Out-Null
            }
        }
        
        if ($UpdatesToDownload.Count -gt 0) {
            Write-Host "`nDownloaden van $($UpdatesToDownload.Count) updates..." -ForegroundColor Cyan
            $Downloader = $UpdateSession.CreateUpdateDownloader()
            $Downloader.Updates = $UpdatesToDownload
            $Downloader.Download()
        }
        
        $SearchResult.Updates | ForEach-Object {
            if ($_.IsDownloaded) {
                $UpdatesToInstall.Add($_) | Out-Null
            }
        }
        
        if ($UpdatesToInstall.Count -gt 0) {
            Write-Host "`nInstalleren van $($UpdatesToInstall.Count) updates..." -ForegroundColor Cyan
            $Installer = $UpdateSession.CreateUpdateInstaller()
            $Installer.Updates = $UpdatesToInstall
            $InstallationResult = $Installer.Install()
            
            if ($InstallationResult.RebootRequired) {
                Write-Host "`nHerstart vereist om de installatie te voltooien!" -ForegroundColor Yellow
            } else {
                Write-Host "`nAlle updates zijn succesvol geïnstalleerd." -ForegroundColor Green
            }
        }
    }
} catch {
    Write-Host "Er is een fout opgetreden: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    Write-Host "`nDruk op een toets om af te sluiten..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}