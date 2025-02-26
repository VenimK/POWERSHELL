Write-host "PC INFO"
Get-CimInstance -ClassName Win32_ComputerSystem

Write-host "BIOS INFO"
Get-CimInstance -ClassName Win32_BIOS