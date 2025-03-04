# Edge Bookmarks Export Script
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
Write-Host "Zoeken naar Edge bookmarks voor gebruiker: $UserName"

# Check alle mogelijke locaties voor Edge bookmarks
$PossiblePaths = @(
    "$Env:systemdrive\Users\$UserName\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks",
    "$Env:systemdrive\Users\$UserName\AppData\Local\Microsoft\Edge Beta\User Data\Default\Bookmarks",
    "$Env:systemdrive\Users\$UserName\AppData\Local\Microsoft\Edge Dev\User Data\Default\Bookmarks"
)

$BookmarksPath = $null
foreach ($Path in $PossiblePaths) {
    Write-Host "Controleren pad: $Path"
    if (Test-Path -Path $Path) {
        $BookmarksPath = $Path
        Write-Host "Bookmarks gevonden op: $Path"
        break
    }
}

if (-not $BookmarksPath) {
    Write-Warning "Kan Edge Bookmarks niet vinden voor gebruiker: $UserName"
    Write-Host "Gecontroleerde locaties:"
    $PossiblePaths | ForEach-Object { Write-Host " - $_" }
    exit
}

# Controleer of Edge actief is
$edgeProcess = Get-Process msedge -ErrorAction SilentlyContinue
if ($edgeProcess) {
    Write-Warning "Sluit Edge browser voordat je dit script uitvoert"
    exit
}

# Als er geen exportpad is opgegeven, toon een mapkeuze dialoog
if (-not $ExportPath) {
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.Description = "Kies een map om de Edge bookmarks te exporteren"
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
    Write-Host "Bezig met lezen van bookmarks bestand..."
    # Lees het bookmarks bestand
    $BookmarksContent = Get-Content -Path $BookmarksPath -Raw
    Write-Host "Bookmarks bestand gelezen, converteren naar JSON..."
    $BookmarksJson = $BookmarksContent | ConvertFrom-Json
    Write-Host "JSON conversie succesvol"

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
            Write-Host "Verwerken van map: $FolderPath"

            foreach ($Child in $Node.children) {
                Get-Bookmarks -Node $Child -FolderPath $FolderPath
            }
        }
        elseif ($Node.type -eq "url") {
            Write-Host "Gevonden bookmark: $($Node.name)"
            [PSCustomObject]@{
                Map = $FolderPath
                Naam = $Node.name
                URL = $Node.url
                DatumToegevoegd = [DateTime]::FromFileTimeUtc($Node.date_added).ToLocalTime()
            }
        }
    }

    Write-Host "Start verwerken van bookmarks..."
    # Verzamel alle bookmarks
    $Results = @()
    
    if ($BookmarksJson.roots.bookmark_bar) {
        Write-Host "Verwerken van bookmarks balk..."
        $Results += Get-Bookmarks -Node $BookmarksJson.roots.bookmark_bar
    }
    
    if ($BookmarksJson.roots.other) {
        Write-Host "Verwerken van andere bookmarks..."
        $Results += Get-Bookmarks -Node $BookmarksJson.roots.other
    }

    # Filter op map indien opgegeven
    if ($FilterFolder) {
        Write-Host "Filteren op map: $FilterFolder"
        $Results = $Results | Where-Object { $_.Map -like "*$FilterFolder*" }
    }

    # Sorteer resultaten indien gewenst
    if ($SortByDate) {
        Write-Host "Sorteren op datum..."
        $Results = $Results | Sort-Object DatumToegevoegd -Descending
    }
    elseif ($SortByName) {
        Write-Host "Sorteren op naam..."
        $Results = $Results | Sort-Object Naam
    }

    # Toon aantal resultaten
    Write-Host "`nTotaal aantal bookmarks gevonden: $($Results.Count)"

    # Bepaal bestandsnaam op basis van datum en tijd
    $TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $FileName = "EdgeBookmarks_$TimeStamp"

    # Export op basis van gekozen formaat
    $ExportFile = Join-Path $ExportPath $FileName
    Write-Host "Exporteren naar: $ExportFile"
    
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
    <title>Edge Bookmarks Export</title>
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
    <h1>Microsoft Edge Bookmarks</h1>
    <p>Geëxporteerd op $(Get-Date -Format "dd-MM-yyyy HH:mm")</p>
    <table>
        <tr>
            <th>Map</th>
            <th>Naam</th>
            <th>URL</th>
            <th>Datum Toegevoegd</th>
        </tr>
"@
            foreach ($item in $Results) {
                $HTML += @"
        <tr>
            <td>$($item.Map)</td>
            <td>$($item.Naam)</td>
            <td><a href="$($item.URL)">$($item.URL)</a></td>
            <td>$($item.DatumToegevoegd)</td>
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

    Write-Host "`nBookmarks zijn geëxporteerd naar: $ExportFile" -Encoding UTF8
    Write-Host "Gebruik parameters voor verschillende export opties:" -Encoding UTF8
    Write-Host " -ExportFormat: 'CSV', 'HTML', of 'JSON'" -Encoding UTF8
    Write-Host " -SortByDate: Sorteer op datum toegevoegd" -Encoding UTF8
    Write-Host " -SortByName: Sorteer op naam" -Encoding UTF8
    Write-Host " -FilterFolder: Filter op mapnaam" -Encoding UTF8
    Write-Host " -ExportPath: Pad waar de bestanden worden opgeslagen" -Encoding UTF8

} catch {
    Write-Error "Fout bij het lezen van Edge bookmarks: $_"
    Write-Host "Stack Trace:"
    Write-Host $_.ScriptStackTrace
}