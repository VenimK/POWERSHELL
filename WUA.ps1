[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,
    [Parameter(Mandatory=$false)]
    [switch]$ExportResults,
    [Parameter(Mandatory=$false)]
    [switch]$CleanupAfterScan,
    [Parameter(Mandatory=$false)]
    [ValidateSet("All", "Required", "Optional")]
    [string]$UpdateType = "All",
    [Parameter(Mandatory=$false)]
    [switch]$InstallUpdates,
    [Parameter(Mandatory=$false)]
    [switch]$GenerateHtmlReport,
    [Parameter(Mandatory=$false)]
    [switch]$SendEmail,
    [Parameter(Mandatory=$false)]
    [string]$SmtpServer,
    [Parameter(Mandatory=$false)]
    [string]$EmailFrom,
    [Parameter(Mandatory=$false)]
    [string]$EmailTo,
    [Parameter(Mandatory=$false)]
    [string]$EmailSubject = "Windows Update Report"
)

# Check for administrator privileges and relaunch if needed
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Relaunching script with administrator privileges..." -ForegroundColor Yellow
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $MyInvocation.MyCommand.Path)
    $argList += $PSBoundParameters.GetEnumerator() | ForEach-Object { "-$($_.Key)", $_.Value }
    Start-Process powershell.exe -ArgumentList $argList -Verb RunAs -Wait
    return
}

# Continue with the rest of the script if we are running as admin
Write-Host "Running with administrator privileges." -ForegroundColor Green

# Get the script's current path and set default output path if not provided
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $OutputPath) {
    $OutputPath = Join-Path $scriptPath "WinUpdate"
}

# Function to format file size
function Format-FileSize {
    param([long]$size)
    $sizes = 'B','KB','MB','GB','TB'
    $index = 0
    while ($size -gt 1kb -and $index -lt ($sizes.Count - 1)) {
        $size = $size / 1kb
        $index++
    }
    return "{0:N2} {1}" -f $size, $sizes[$index]
}

# Function to show progress
function Show-DownloadProgress {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Activity,
        [Parameter(Mandatory=$true)]
        [long]$BytesReceived,
        [Parameter(Mandatory=$true)]
        [long]$TotalBytes
    )
    
    $percentComplete = ($BytesReceived / $TotalBytes) * 100
    $downloadedSize = Format-FileSize -size $BytesReceived
    $totalSize = Format-FileSize -size $TotalBytes
    
    Write-Progress -Activity $Activity -Status "$downloadedSize of $totalSize" -PercentComplete $percentComplete
}

# Function to generate HTML report
function Generate-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Updates,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        [Parameter(Mandatory=$true)]
        [datetime]$ScanTime
    )

    $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows Update Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #2c3e50; }
        h2 { color: #34495e; margin-top: 20px; }
        .summary { background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin: 10px 0; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #34495e; color: white; }
        tr:nth-child(even) { background-color: #f8f9fa; }
        tr:hover { background-color: #f1f2f3; }
        .required { color: #e74c3c; }
        .optional { color: #7f8c8d; }
        .severity-critical { color: #c0392b; font-weight: bold; }
        .severity-important { color: #d35400; }
        .severity-moderate { color: #f39c12; }
        .footer { margin-top: 30px; color: #7f8c8d; font-size: 0.9em; }
    </style>
</head>
<body>
    <h1>Windows Update Report</h1>
    <div class="summary">
        <p><strong>Scan Date:</strong> $($ScanTime.ToString("yyyy-MM-dd HH:mm:ss"))</p>
        <p><strong>Computer Name:</strong> $env:COMPUTERNAME</p>
        <p><strong>Total Updates Found:</strong> $($Updates.Count)</p>
        <p><strong>Required Updates:</strong> $($Updates.Where({-not $_.IsOptional}).Count)</p>
        <p><strong>Optional Updates:</strong> $($Updates.Where({$_.IsOptional}).Count)</p>
    </div>
"@

    $htmlUpdates = "<h2>Available Updates</h2>`n<table>`n<tr><th>Title</th><th>KB</th><th>Size</th><th>Type</th><th>Severity</th><th>Reboot Required</th></tr>`n"
    
    foreach ($update in $Updates) {
        $typeClass = if ($update.IsOptional) { "optional" } else { "required" }
        $severityClass = switch ($update.MsrcSeverity) {
            "Critical" { "severity-critical" }
            "Important" { "severity-important" }
            "Moderate" { "severity-moderate" }
            default { "" }
        }
        
        $htmlUpdates += "<tr>
            <td>$($update.Title)</td>
            <td>$(if ($update.KBArticleIDs) { "KB" + ($update.KBArticleIDs -join ", KB") } else { "N/A" })</td>
            <td>$(if ($update.MaxDownloadSize) { Format-FileSize $update.MaxDownloadSize } else { "Unknown" })</td>
            <td class='$typeClass'>$(if ($update.IsOptional) { "Optional" } else { "Required" })</td>
            <td class='$severityClass'>$($update.MsrcSeverity)</td>
            <td>$($update.RebootRequired)</td>
        </tr>`n"
    }

    $htmlFooter = @"
    </table>
    <div class="footer">
        <p>Report generated on $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))</p>
        <p>Windows Update Analysis Script</p>
    </div>
</body>
</html>
"@

    $htmlContent = $htmlHeader + $htmlUpdates + $htmlFooter
    $htmlPath = Join-Path $OutputPath "UpdateReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
    return $htmlPath
}

# Function to send email
function Send-UpdateReport {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SmtpServer,
        [Parameter(Mandatory=$true)]
        [string]$From,
        [Parameter(Mandatory=$true)]
        [string]$To,
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        [Parameter(Mandatory=$true)]
        [string]$HtmlReport,
        [Parameter(Mandatory=$true)]
        [array]$Updates
    )

    $requiredCount = ($Updates | Where-Object { -not $_.IsOptional }).Count
    $optionalCount = ($Updates | Where-Object { $_.IsOptional }).Count
    
    $emailBody = @"
<html>
<body>
<h2>Windows Update Summary</h2>
<p>A Windows Update scan was performed on $env:COMPUTERNAME.</p>
<p><strong>Results:</strong></p>
<ul>
    <li>Total Updates Found: $($Updates.Count)</li>
    <li>Required Updates: $requiredCount</li>
    <li>Optional Updates: $optionalCount</li>
</ul>
<p>Please see the attached report for detailed information.</p>
</body>
</html>
"@

    $attachment = $HtmlReport

    try {
        $securePassword = ConvertTo-SecureString "YourEmailPassword" -AsPlainText -Force
        $credentials = New-Object System.Management.Automation.PSCredential ($From, $securePassword)
        
        Send-MailMessage -SmtpServer $SmtpServer `
                        -From $From `
                        -To $To `
                        -Subject $Subject `
                        -Body $emailBody `
                        -BodyAsHtml `
                        -Attachments $attachment `
                        -Encoding UTF8 `
                        -ErrorAction Stop
        
        Write-Host "Email report sent successfully to $To" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to send email: $_" -ForegroundColor Red
    }
}

# Initialize language strings
$Messages = @{
    StartTranscript = "Start logboek registratie..."
    CreateOutputDir = "Uitvoermap aangemaakt: {0}"
    StartAnalysis = "Windows Update Analyse wordt gestart..."
    AdminRequired = "Fout: Voer dit script uit als administrator."
    DownloadingCab = "Nieuwe wsusscn2.cab bestand wordt gedownload..."
    CabUpToDate = "Lokale wsusscn2.cab is up-to-date."
    DownloadComplete = "Download succesvol voltooid!"
    InitSession = "Update sessie wordt gestart..."
    SearchingUpdates = "Zoeken naar updates..."
    SearchProgress = "Zoeken naar updates... Dit kan enkele minuten duren."
    SearchCriteria = "Zoekcriteria: {0}"
    NoUpdatesFound = "Geen beschikbare updates gevonden."
    FoundUpdates = "{0} beschikbare updates gevonden:"
    UpdateTitle = "Titel: {0}"
    UpdateKB = "KB-nummer: {0}"
    UpdateSize = "Grootte: {0}"
    UpdateSeverity = "Ernst: {0}"
    UpdateType = "Type: {0}"
    UpdateCategories = "Categorieën: {0}"
    ExportComplete = "Resultaten geëxporteerd naar: {0}"
    CleaningUp = "Tijdelijke bestanden worden opgeruimd..."
    CleanupComplete = "Opruimen succesvol voltooid."
    CleanupError = "Fout tijdens opruimen: {0}"
    ScriptComplete = "Script succesvol voltooid!"
    Required = "Vereist"
    Optional = "Optioneel"
    ErrorInit = "Fout bij initialiseren update sessie: {0}"
    ErrorSearch = "Fout bij zoeken naar updates: {0}"
    ErrorDownload = "Fout bij downloaden wsusscn2.cab: {0}"
    InstallPrompt = "Wilt u deze updates installeren? (J/N): "
    StartingInstall = "Start met installeren van updates..."
    DownloadingUpdate = "Downloaden van update: {0}"
    InstallingUpdate = "Installeren van update: {0}"
    InstallComplete = "Installatie voltooid."
    InstallError = "Fout bij installeren: {0}"
    RebootRequired = "Computer moet opnieuw worden opgestart om de updates te voltooien."
    InstallProgress = "Voortgang: {0}%"
    Phase = "=== FASE: {0} ==="
    Progress = "Voortgang: {0}"
}

# Check and set execution policy
$currentPolicy = Get-ExecutionPolicy
if ($currentPolicy -eq "Restricted" -or $currentPolicy -eq "AllSigned") {
    try {
        Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
        Write-Host "Execution policy set to Bypass for current process" -ForegroundColor Yellow
    } catch {
        Write-Host "Failed to set execution policy. Script may not run correctly: $_" -ForegroundColor Red
        Exit 1
    }
}

# Set output encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

# Start Transcript for logging
$logPath = Join-Path $OutputPath "WUA_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $logPath -Force
Write-Host $Messages.StartTranscript -ForegroundColor Cyan

# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Define the URL for the wsusscn2.cab file
$cabUrl = "https://catalog.s.download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab"
# Specify the local path where the file should be stored
$localPath = Join-Path $OutputPath "wsusscn2.cab"

# Create output directory if it doesn't exist
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Host ($Messages.CreateOutputDir -f $OutputPath) -ForegroundColor Green
}

Write-Host $Messages.StartAnalysis -ForegroundColor Cyan

# Check for administrator privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host $Messages.AdminRequired -ForegroundColor Red
    Stop-Transcript
    Exit 1
}

# Download the wsusscn2.cab file with progress
try {
    $needsDownload = $false
    
    # First check if local file exists
    if (-not (Test-Path $localPath)) {
        $needsDownload = $true
        Write-Host $Messages.DownloadingCab -ForegroundColor Yellow
    } else {
        # File exists, only check server if we can reach it
        try {
            $webRequest = [System.Net.WebRequest]::Create($cabUrl)
            $webRequest.Method = "HEAD"
            $webRequest.Timeout = 5000  # 5 second timeout
            $webResponse = $webRequest.GetResponse()
            $webLastModified = [datetime]::Parse($webResponse.Headers["Last-Modified"])
            $totalBytes = [long]$webResponse.Headers["Content-Length"]
            $webResponse.Close()

            $localFileInfo = Get-Item $localPath
            if ($webLastModified -gt $localFileInfo.LastWriteTime) {
                $needsDownload = $true
                Write-Host $Messages.DownloadingCab -ForegroundColor Yellow
            } else {
                Write-Host $Messages.CabUpToDate -ForegroundColor Green
            }
        } catch {
            # Can't reach server, use existing file
            Write-Host "Cannot check for newer version. Using existing local file." -ForegroundColor Yellow
        }
    }

    if ($needsDownload) {
        Write-Host $Messages.DownloadingCab -ForegroundColor Yellow
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell Script")
        
        $webClient.DownloadFileAsync($cabUrl, $localPath)
        
        Register-ObjectEvent -InputObject $webClient -EventName DownloadProgressChanged -Action {
            Show-DownloadProgress -Activity $Messages.DownloadingCab `
                                -BytesReceived $EventArgs.BytesReceived `
                                -TotalBytes $EventArgs.TotalBytesToReceive
        } | Out-Null
        
        Register-ObjectEvent -InputObject $webClient -EventName DownloadFileCompleted -Action {
            Write-Host $Messages.DownloadComplete -ForegroundColor Green
            $Global:DownloadComplete = $true
        } | Out-Null

        while (-not $Global:DownloadComplete) {
            Start-Sleep -Milliseconds 100
        }
    }
} catch {
    Write-Host ($Messages.ErrorDownload -f $_) -ForegroundColor Red
    if (Test-Path $localPath) {
        Write-Host "Using existing local file instead." -ForegroundColor Yellow
    } else {
        Stop-Transcript
        Exit 1
    }
}

# Initialize update session
try {
    Write-Host ($Messages.Phase -f "INITIALISATIE") -ForegroundColor Cyan
    Write-Host $Messages.InitSession -ForegroundColor Cyan
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager
    $UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", $localPath)
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
} catch {
    Write-Host ($Messages.ErrorInit -f $_) -ForegroundColor Red
    Stop-Transcript
    Exit 1
}

# Search for updates
Write-Host ($Messages.Phase -f "ZOEKEN NAAR UPDATES") -ForegroundColor Cyan
Write-Host $Messages.SearchingUpdates -ForegroundColor Cyan
$UpdateSearcher.ServerSelection = 3 # ssOthers
$UpdateSearcher.ServiceID = [string] $UpdateService.ServiceID

try {
    # Build search criteria based on update type
    $searchCriteria = "IsInstalled=0"
    
    # Get all updates first
    Write-Host ($Messages.SearchCriteria -f $searchCriteria) -ForegroundColor Gray
    Write-Host $Messages.SearchProgress -ForegroundColor Yellow
    
    # Create a simple progress indicator
    $spin = @('|', '/', '-', '\')
    $spinIndex = 0
    $progressTimer = [System.Diagnostics.Stopwatch]::StartNew()
    
    # Start a background job just for the progress display
    $progressJob = Start-Job -ScriptBlock {
        $spin = @('|', '/', '-', '\')
        $spinIndex = 0
        while ($true) {
            $spinIndex = ($spinIndex + 1) % 4
            Write-Host "`r$($spin[$spinIndex]) Searching... " -NoNewline
            Start-Sleep -Milliseconds 200
        }
    }
    
    # Perform the search
    $SearchResult = $UpdateSearcher.Search($searchCriteria)
    
    # Stop the progress display
    Stop-Job -Job $progressJob
    Remove-Job -Job $progressJob -Force
    Write-Host "`r" -NoNewline  # Clear the progress line
    
    $progressTimer.Stop()
    $searchTime = [math]::Round($progressTimer.Elapsed.TotalSeconds)
    Write-Host "Search completed in $searchTime seconds." -ForegroundColor Green
    
    $AllUpdates = $SearchResult.Updates
    
    # Filter updates based on type
    if ($UpdateType -eq "Required") {
        $Updates = $AllUpdates | Where-Object { -not $_.IsOptional }
    } elseif ($UpdateType -eq "Optional") {
        $Updates = $AllUpdates | Where-Object { $_.IsOptional }
    } else {
        $Updates = $AllUpdates
    }

    if ($Updates.Count -eq 0) {
        Write-Host $Messages.NoUpdatesFound -ForegroundColor Green
    } else {
        Write-Host ($Messages.FoundUpdates -f $Updates.Count) -ForegroundColor Cyan
        Write-Host "`nUpdate Details:" -ForegroundColor Yellow
        
        # Create an array to store update details
        $updateDetails = @()
        
        foreach ($Update in $Updates) {
            $updateInfo = [PSCustomObject]@{
                Title = $Update.Title
                KB = $(if ($Update.KBArticleIDs) { "KB" + ($Update.KBArticleIDs -join ", KB") } else { "N/A" })
                Size = $(if ($Update.MaxDownloadSize) { Format-FileSize $Update.MaxDownloadSize } else { "Unknown" })
                Type = $(if ($Update.IsOptional) { $Messages.Optional } else { $Messages.Required })
                Severity = $Update.MsrcSeverity
                Categories = $($Update.Categories | ForEach-Object { $_.Name } | Join-String -Separator ", ")
                RebootRequired = $Update.RebootRequired
                Description = $Update.Description
            }
            $updateDetails += $updateInfo
        }

        # Display updates in a formatted table
        $updateDetails | Format-Table -Property Title, KB, Size, Type, Severity, RebootRequired -AutoSize -Wrap

        # Export results if requested
        if ($ExportResults) {
            $csvPath = Join-Path $OutputPath "UpdateResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $updateDetails | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host ($Messages.ExportComplete -f $csvPath) -ForegroundColor Green
        }

        # Generate HTML report if requested
        if ($GenerateHtmlReport) {
            Write-Host ($Messages.Phase -f "HTML RAPPORT GENEREREN") -ForegroundColor Cyan
            try {
                $htmlReportPath = Generate-HTMLReport -Updates $Updates -OutputPath $OutputPath -ScanTime (Get-Date)
                Write-Host "HTML report generated successfully at: $htmlReportPath" -ForegroundColor Green
                
                # Verify the file exists
                if (Test-Path $htmlReportPath) {
                    Write-Host "Report file verified at location." -ForegroundColor Green
                } else {
                    Write-Host "Warning: Report file not found at expected location!" -ForegroundColor Red
                }
            }
            catch {
                Write-Host "Error generating HTML report: $_" -ForegroundColor Red
            }
        }

        # Send email report if requested
        if ($SendEmail) {
            Send-UpdateReport -SmtpServer $SmtpServer -From $EmailFrom -To $EmailTo -Subject $EmailSubject -HtmlReport $htmlReportPath -Updates $Updates
        }

        # Install updates if requested
        if ($InstallUpdates) {
            Write-Host ($Messages.Phase -f "UPDATES INSTALLEREN") -ForegroundColor Cyan
            
            $totalUpdates = $Updates.Count
            $currentUpdate = 0
            $rebootRequired = $false
            
            foreach ($Update in $Updates) {
                $currentUpdate++
                $percentComplete = [math]::Round(($currentUpdate / $totalUpdates) * 100)
                
                # Display update progress header
                Write-Host "`n=== Update $currentUpdate/$totalUpdates ($percentComplete%) ===" -ForegroundColor Yellow
                Write-Host "Title: $($Update.Title)" -ForegroundColor Cyan
                
                # Download phase
                Write-Host "`nDownloading update..." -ForegroundColor Yellow
                $downloadJob = Start-Job -ScriptBlock {
                    $spin = @('|', '/', '-', '\')
                    $spinIndex = 0
                    while ($true) {
                        $spinIndex = ($spinIndex + 1) % 4
                        Write-Host "`r$($spin[$spinIndex]) Downloading... " -NoNewline
                        Start-Sleep -Milliseconds 200
                    }
                }
                
                try {
                    if (-not $Update.IsDownloaded) {
                        $session = New-Object -ComObject Microsoft.Update.Session
                        $downloader = $session.CreateUpdateDownloader()
                        $downloader.Updates.Add($Update) | Out-Null
                        $downloadResult = $downloader.Download()
                        
                        # Stop the progress spinner
                        Stop-Job -Job $downloadJob
                        Remove-Job -Job $downloadJob -Force
                        Write-Host "`rDownload completed successfully" -ForegroundColor Green
                    } else {
                        Stop-Job -Job $downloadJob
                        Remove-Job -Job $downloadJob -Force
                        Write-Host "`rUpdate already downloaded" -ForegroundColor Green
                    }
                    
                    # Installation phase
                    Write-Host "`nInstalling update..." -ForegroundColor Yellow
                    
                    # Accept EULA if needed
                    if (-not $Update.EulaAccepted) {
                        $Update.AcceptEula()
                    }
                    
                    $installJob = Start-Job -ScriptBlock {
                        $spin = @('|', '/', '-', '\')
                        $spinIndex = 0
                        while ($true) {
                            $spinIndex = ($spinIndex + 1) % 4
                            Write-Host "`r$($spin[$spinIndex]) Installing... " -NoNewline
                            Start-Sleep -Milliseconds 200
                        }
                    }
                    
                    $session = New-Object -ComObject Microsoft.Update.Session
                    $installer = $session.CreateUpdateInstaller()
                    $installer.Updates.Add($Update) | Out-Null
                    $installResult = $installer.Install()
                    
                    # Stop the progress spinner
                    Stop-Job -Job $installJob
                    Remove-Job -Job $installJob -Force
                    
                    if ($installResult.ResultCode -eq 2) { # orcSucceeded
                        Write-Host "`rInstallation completed successfully" -ForegroundColor Green
                        if ($installResult.RebootRequired) {
                            $rebootRequired = $true
                        }
                    } else {
                        Write-Host "`rInstallation failed with code: $($installResult.ResultCode)" -ForegroundColor Red
                    }
                    
                } catch {
                    Write-Host "`rError during update process: $_" -ForegroundColor Red
                    # Ensure jobs are cleaned up in case of error
                    Get-Job | Where-Object { $_.Name -match 'download|install' } | Remove-Job -Force
                }
            }
            
            Write-Host "`n=== Installation Summary ===" -ForegroundColor Cyan
            Write-Host "Total updates processed: $totalUpdates" -ForegroundColor Green
            if ($rebootRequired) {
                Write-Host $Messages.RebootRequired -ForegroundColor Yellow
            }
        }
    }
} catch {
    Write-Host ($Messages.ErrorSearch -f $_) -ForegroundColor Red
    Stop-Transcript
    Exit 1
}

# Cleanup if requested
if ($CleanupAfterScan) {
    Write-Host "`n$($Messages.CleaningUp)" -ForegroundColor Cyan
    try {
        Remove-Item $localPath -Force
        Write-Host $Messages.CleanupComplete -ForegroundColor Green
    } catch {
        Write-Host ($Messages.CleanupError -f $_) -ForegroundColor Red
    }
}

Write-Host "`n$($Messages.ScriptComplete)" -ForegroundColor Green
Stop-Transcript