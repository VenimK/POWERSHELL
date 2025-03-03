# Chrome Bookmarks Export Script
param (
    [ValidateSet('CSV', 'HTML', 'JSON')]
    [string]$ExportFormat = 'CSV',
    [switch]$SortByDate,
    [switch]$SortByName,
    [string]$FilterFolder,
    [string]$ExportPath
)

# Set encoding for proper Dutch character display
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

$UserName = $env:USERNAME
$BookmarksPath = "$Env:systemdrive\Users\$UserName\AppData\Local\Google\Chrome\User Data\Default\Bookmarks"

if (-not (Test-Path -Path $BookmarksPath)) {
    Write-Warning "Kan Chrome Bookmarks niet vinden voor gebruiker: $UserName"
    exit
}

# Als er geen exportpad is opgegeven, toon een mapkeuze dialoog
if (-not $ExportPath) {
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.Description = "Kies een map om de bookmarks te exporteren"
    $FolderBrowser.ShowNewFolderButton = $true
    
    if ($FolderBrowser.ShowDialog() -eq 'OK') {
        $ExportPath = $FolderBrowser.SelectedPath
    } else {
        Write-Warning "Geen exportlocatie geselecteerd. Gebruik huidige map."
        $ExportPath = $PSScriptRoot
    }
}

# Controleer of het exportpad bestaat
if (-not (Test-Path -Path $ExportPath)) {
    try {
        New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
        Write-Host "Map aangemaakt: $ExportPath"
    } catch {
        Write-Error "Kon de map niet aanmaken: $ExportPath"
        exit
    }
}

try {
    # Lees het bookmarks bestand
    $BookmarksJson = Get-Content -Path $BookmarksPath -Raw | ConvertFrom-Json

    # Functie om recursief door bookmarks te gaan
    function Get-Bookmarks {
        param (
            [Parameter(Mandatory = $true)]
            $Node,
            [string]$FolderPath = ""
        )

        if ($Node.type -eq "folder") {
            $FolderPath = if ($FolderPath) {
                "$FolderPath\$($Node.name)"
            } else {
                $Node.name
            }

            foreach ($Child in $Node.children) {
                Get-Bookmarks -Node $Child -FolderPath $FolderPath
            }
        }
        elseif ($Node.type -eq "url") {
            [PSCustomObject]@{
                Map = $FolderPath
                Naam = $Node.name
                URL = $Node.url
                DatumToegevoegd = [DateTime]::FromFileTimeUtc($Node.date_added).ToLocalTime()
                LaatsteBezoek = if ($Node.last_visited) {
                    [DateTime]::FromFileTimeUtc($Node.last_visited).ToLocalTime()
                } else { $null }
            }
        }
    }

    # Verzamel alle bookmarks
    $Results = @()
    $Results += Get-Bookmarks -Node $BookmarksJson.roots.bookmark_bar
    $Results += Get-Bookmarks -Node $BookmarksJson.roots.other

    # Filter op map indien opgegeven
    if ($FilterFolder) {
        $Results = $Results | Where-Object { $_.Map -like "*$FilterFolder*" }
    }

    # Sorteer resultaten indien gewenst
    if ($SortByDate) {
        $Results = $Results | Sort-Object DatumToegevoegd -Descending
    }
    elseif ($SortByName) {
        $Results = $Results | Sort-Object Naam
    }

    # Toon aantal bookmarks
    Write-Host "`nTotaal aantal bookmarks gevonden: $($Results.Count)"

    # Bepaal bestandsnaam op basis van datum en tijd
    $TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $FileName = "ChromeBookmarks_$TimeStamp"

    # Export op basis van gekozen formaat
    $ExportFile = Join-Path $ExportPath $FileName
    switch ($ExportFormat) {
        'CSV' {
            $ExportFile = "$ExportFile.csv"
            $Results | Export-Csv -Path $ExportFile -NoTypeInformation -Encoding UTF8
        }
        'HTML' {
            $ExportFile = "$ExportFile.html"
            $HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Chrome Bookmarks Export</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tr:hover { background-color: #f5f5f5; }
    </style>
</head>
<body>
    <h1>Chrome Bookmarks</h1>
    <p>Geëxporteerd op $(Get-Date -Format "dd-MM-yyyy HH:mm")</p>
    <table>
        <tr>
            <th>Map</th>
            <th>Naam</th>
            <th>URL</th>
            <th>Datum Toegevoegd</th>
            <th>Laatste Bezoek</th>
        </tr>
"@
            foreach ($bookmark in $Results) {
                $HTML += @"
        <tr>
            <td>$($bookmark.Map)</td>
            <td>$($bookmark.Naam)</td>
            <td><a href="$($bookmark.URL)">$($bookmark.URL)</a></td>
            <td>$($bookmark.DatumToegevoegd)</td>
            <td>$($bookmark.LaatsteBezoek)</td>
        </tr>
"@
            }
            $HTML += @"
    </table>
</body>
</html>
"@
            $HTML | Out-File -FilePath $ExportFile -Encoding UTF8
        }
        'JSON' {
            $ExportFile = "$ExportFile.json"
            $Results | ConvertTo-Json -Depth 10 | Out-File -FilePath $ExportFile -Encoding UTF8
        }
    }

    # Toon resultaten in console
    $Results | Format-Table -AutoSize -Wrap

    Write-Host "`nBookmarks zijn geexporteerd naar: $ExportFile" -Encoding UTF8
    Write-Host "Gebruik parameters voor verschillende export opties:" -Encoding UTF8
    Write-Host " -ExportFormat: 'CSV', 'HTML', of 'JSON'" -Encoding UTF8
    Write-Host " -SortByDate: Sorteer op datum toegevoegd" -Encoding UTF8
    Write-Host " -SortByName: Sorteer op naam" -Encoding UTF8
    Write-Host " -FilterFolder: Filter op mapnaam" -Encoding UTF8
    Write-Host " -ExportPath: Pad waar de bestanden worden opgeslagen" -Encoding UTF8

} catch {
    Write-Error "Fout bij het lezen van Chrome bookmarks: $_"
}
