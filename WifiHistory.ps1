# Function to convert time to readable format
function Format-TimeSpan {
    param (
        [TimeSpan]$TimeSpan
    )
    if ($TimeSpan.Days -gt 0) {
        return "$($TimeSpan.Days) dagen, $($TimeSpan.Hours) uren, $($TimeSpan.Minutes) minuten"
    }
    elseif ($TimeSpan.Hours -gt 0) {
        return "$($TimeSpan.Hours) uren, $($TimeSpan.Minutes) minuten"
    }
    else {
        return "$($TimeSpan.Minutes) minuten"
    }
}

try {
    Write-Host "WiFi verbindingsgeschiedenis van de laatste 365 dagen:" -ForegroundColor Cyan
    Write-Host "------------------------------------------------" -ForegroundColor Cyan

    # Get current network information
    $currentInterface = netsh wlan show interfaces | Out-String
    $currentSSID = if ($currentInterface -match "SSID\s*:\s*(.*)") { $matches[1].Trim() }
    $currentSignal = if ($currentInterface -match "Signal\s*:\s*(.*)") { $matches[1].Trim() }
    $currentChannel = if ($currentInterface -match "Channel\s*:\s*(.*)") { $matches[1].Trim() }
    $currentRadio = if ($currentInterface -match "Radio type\s*:\s*(.*)") { $matches[1].Trim() }
    
    # Get all saved WiFi profiles
    $profiles = netsh wlan show profiles | Select-String "All User Profile\s+:\s+(.+)" | ForEach-Object {
        $_.Matches.Groups[1].Value.Trim()
    }
    
    Write-Host "`nOpgeslagen WiFi netwerken gevonden: $($profiles.Count)" -ForegroundColor Cyan
    
    # Get event history
    $startTime = (Get-Date).AddDays(-365)
    $events = Get-WinEvent -FilterHashtable @{
        LogName = "Microsoft-Windows-WLAN-AutoConfig/Operational"
        ID = @(8001, 8002, 8003)
        StartTime = $startTime
    } -ErrorAction Stop

    # Display current connection info
    if ($currentSSID) {
        Write-Host "`nHuidige verbinding:" -ForegroundColor Green
        Write-Host "  Netwerk: $currentSSID" -ForegroundColor Green
        Write-Host "  Signaalsterkte: $currentSignal" -ForegroundColor Green
        Write-Host "  Kanaal: $currentChannel" -ForegroundColor Green
        Write-Host "  Radio type: $currentRadio" -ForegroundColor Green
    }

    # Initialize stats for all profiles
    $networkStats = @{}
    foreach ($profile in $profiles) {
        $networkStats[$profile] = @{
            Name = $profile
            Connections = 0
            Disconnections = 0
            Failed = 0
            LastSeen = $null
            FirstSeen = $null
        }
    }

    # Count events per network
    foreach ($event in $events) {
        $networkName = ($event.Message -split "SSID: ")[1] -split "`r`n" | Select-Object -First 1
        if (-not $networkStats.ContainsKey($networkName)) {
            $networkStats[$networkName] = @{
                Name = $networkName
                Connections = 0
                Disconnections = 0
                Failed = 0
                LastSeen = $event.TimeCreated
                FirstSeen = $event.TimeCreated
            }
        }
        
        switch ($event.Id) {
            8001 { 
                $networkStats[$networkName].Connections++
                if ($null -eq $networkStats[$networkName].FirstSeen -or $event.TimeCreated -lt $networkStats[$networkName].FirstSeen) {
                    $networkStats[$networkName].FirstSeen = $event.TimeCreated
                }
                if ($null -eq $networkStats[$networkName].LastSeen -or $event.TimeCreated -gt $networkStats[$networkName].LastSeen) {
                    $networkStats[$networkName].LastSeen = $event.TimeCreated
                }
            }
            8002 { $networkStats[$networkName].Disconnections++ }
            8003 { $networkStats[$networkName].Failed++ }
        }
    }

    # Convert to array and sort by number of connections
    $sortedNetworks = $networkStats.Values | Sort-Object -Property Connections -Descending

    # Display network statistics
    Write-Host "`nNetwerk statistieken (gesorteerd op aantal verbindingen):" -ForegroundColor Cyan
    Write-Host "------------------------------------------------" -ForegroundColor Cyan
    
    $maxConnections = ($sortedNetworks | Where-Object { $_.Connections -gt 0 } | Select-Object -First 1).Connections
    if ($maxConnections -eq 0) { $maxConnections = 1 }
    
    foreach ($stats in $sortedNetworks) {
        $barLength = [math]::Round(($stats.Connections / $maxConnections) * 30)
        $bar = "#" * $barLength + "." * (30 - $barLength)
        
        $differenceFromTop = $maxConnections - $stats.Connections
        $color = if ($stats.Name -eq $currentSSID) { "Green" } else { "White" }
        
        Write-Host "`nNetwerk: $($stats.Name)" -ForegroundColor $color
        Write-Host "  Verbindingen: [$bar] $($stats.Connections)x" -ForegroundColor $color
        if ($differenceFromTop -gt 0 -and $stats.Connections -gt 0) {
            Write-Host "  Verschil met meest gebruikte: -$differenceFromTop verbindingen" -ForegroundColor Yellow
        }
        Write-Host "  Verbroken verbindingen: $($stats.Disconnections)"
        Write-Host "  Mislukte verbindingen: $($stats.Failed)"
        
        if ($stats.FirstSeen) {
            Write-Host "  Eerste keer gezien: $($stats.FirstSeen.ToString('dd-MM-yyyy HH:mm:ss'))"
        }
        if ($stats.LastSeen) {
            Write-Host "  Laatst gezien: $($stats.LastSeen.ToString('dd-MM-yyyy HH:mm:ss'))"
        }
        
        # Calculate success rate
        $totalAttempts = $stats.Connections + $stats.Failed
        if ($totalAttempts -gt 0) {
            $successRate = ($stats.Connections / $totalAttempts) * 100
            Write-Host "  Succes percentage: $([math]::Round($successRate, 1))%"
        }
        elseif ($stats.Connections -eq 0 -and $stats.Failed -eq 0) {
            Write-Host "  Status: Geen verbindingen in de laatste 365 dagen" -ForegroundColor Yellow
        }
    }
}
catch {
    if ($_.Exception.Message -match "No events were found") {
        Write-Host "Geen WiFi gebeurtenissen gevonden in de opgegeven periode." -ForegroundColor Yellow
    }
    else {
        Write-Host "Fout bij ophalen van WiFi geschiedenis: $($_.Exception.Message)" -ForegroundColor Red
    }
}
