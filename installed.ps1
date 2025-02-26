Get-CimInstance Win32_OperatingSystem | Select-Object  Caption, InstallDate, ServicePackMajorVersion, OSArchitecture, BootDevice,  BuildNumber, CSName | FL


CLS

  Write-Host -NoNewLine "OS Version: "

  Get-CimInstance Win32_OperatingSystem | Select-Object  Caption | ForEach{ $_.Caption }

  Write-Host ""
Write-Host -NoNewLine "Geinstalleerd Op: "

  Get-CimInstance Win32_OperatingSystem | Select-Object  InstallDate | ForEach{ $_.InstallDate }

  Write-Host ""


Write-Host -NoNewLine "Service Pack Version: "

  Get-CimInstance Win32_OperatingSystem | Select-Object  ServicePackMajorVersion | ForEach{ $_.ServicePackMajorVersion }

  Write-Host ""


Write-Host -NoNewLine "OS Architecture: "

  Get-CimInstance Win32_OperatingSystem | Select-Object  OSArchitecture | ForEach{ $_.OSArchitecture }

  Write-Host ""


Write-Host -NoNewLine "Boot Device: "

  Get-CimInstance Win32_OperatingSystem | Select-Object  BootDevice | ForEach{ $_.BootDevice }

  Write-Host ""


Write-Host -NoNewLine "Build Number: "

  Get-CimInstance Win32_OperatingSystem | Select-Object  BuildNumber | ForEach{ $_.BuildNumber }

  Write-Host ""


Write-Host -NoNewLine "Host Name: "

  Get-CimInstance Win32_OperatingSystem | Select-Object  CSName | ForEach{ $_.CSName }

  Write-Host ""




  Write-Host ""
