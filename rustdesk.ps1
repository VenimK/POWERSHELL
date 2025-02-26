$ErrorActionPreference= 'silentlycontinue'

# Assign the value random password to the password variable
$rustdesk_pw='MusicLover1'

# Get your config string from your Web portal and Fill Below
$rustdesk_cfg="==Qfi0TUxRlW6FUciVUVOZ0NCp0VSBVVYtEO0QzT3xGZ1ETYMplTOFUUKZWWHVnWiojI5V2aiwiIiojIpBXYiwiIiojI5FGblJnIsISdl5iN4MXYu5yazVGZiojI0N3boJye"

################################### Please Do Not Edit Below This Line #########################################

# Run as administrator and stays in the current directory
if (-Not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000)
    {
        Start-Process PowerShell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"cd '$pwd'; & '$PSCommandPath';`"";
        Exit;
    }
}

# This function will return the latest version and download link as an object
function getLatest()
{
    $Page = Invoke-WebRequest -Uri 'https://github.com/rustdesk/rustdesk/releases/latest' -UseBasicParsing
    $HTML = New-Object -Com "HTMLFile"
    try
    {
        $HTML.IHTMLDocument2_write($Page.Content)
    }
    catch
    {
        $src = [System.Text.Encoding]::Unicode.GetBytes($Page.Content)
        $HTML.write($src)
    }

    # Current example link: https://github.com/rustdesk/rustdesk/releases/download/1.2.3/rustdesk-1.2.3-x86_64.exe
    $Downloadlink = ($HTML.Links | Where {$_.href -match '(.)+\/rustdesk\/rustdesk\/releases\/download\/\d{1}.\d{1,2}.\d{1,2}(.{0,3})\/rustdesk(.)+x86_64.exe'} | select -first 1).href

    # bugfix - sometimes you need to replace "about:"
    $Downloadlink = $Downloadlink.Replace('about:', 'https://github.com')

    $Version = "unknown"
    if ($Downloadlink -match './rustdesk/rustdesk/releases/download/(?<content>.*)/rustdesk-(.)+x86_64.exe')
    {
        $Version = $matches['content']
    }

    if ($Version -eq "unknown" -or $Downloadlink -eq "")
    {
        Write-Output "ERROR: Version or download link not found."
        Exit
    }

    # Create object to return
    $params += @{Version = $Version}
    $params += @{Downloadlink = $Downloadlink}
    $Result = New-Object PSObject -Property $params
   
    return($Result)
}

$RustDeskOnGitHub = getLatest


$rdver = ((Get-ItemProperty  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\RustDesk\").Version)

if ($rdver -eq $RustDeskOnGitHub.Version)
{
    Write-Output "RustDesk $rdver is the newest version."
    Exit
}

if (!(Test-Path C:\Temp))
{
    New-Item -ItemType Directory -Force -Path C:\Temp > null
}

cd C:\Temp

Invoke-WebRequest $RustDeskOnGitHub.Downloadlink -Outfile "Rustdesk.exe"
Start-Process .\Rustdesk.exe --silent-install
Start-Sleep -seconds 20

$ServiceName = 'Rustdesk'
$arrService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

if ($arrService -eq $null)
{
    Write-Output "Installing service"
    cd $env:ProgramFiles\Rustdesk
    Start-Process .\Rustdesk.exe --install-service
    Start-Sleep -seconds 20
}

while ($arrService.Status -ne 'Running')
{
    Start-Service $ServiceName
    Start-Sleep -seconds 5
    $arrService.Refresh()
}

cd $env:ProgramFiles\RustDesk\
.\rustdesk.exe --get-id | Write-Output -OutVariable rustdesk_id

.\rustdesk.exe --config $rustdesk_cfg

.\rustdesk.exe --password $rustdesk_pw

Write-Output "..............................................."
# Show the value of the ID Variable
Write-Output "RustDesk ID: $rustdesk_id"

# Show the value of the Password Variable
Write-Output "Password: $rustdesk_pw"
Write-Output "..............................................."