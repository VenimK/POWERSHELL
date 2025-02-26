Write-Ascii  "MusicLover -  Controle Windows Systeem" -fore yellow

sfc /scannow

Write-Ascii  "MusicLover - Controle Image Windows" -fore blue
DISM /Online /Cleanup-Image /CheckHealth

Write-Ascii  "MusicLover - Scan Image Windows" -fore red
DISM /Online /Cleanup-Image /ScanHealth


Write-Ascii  "MusicLover - Herstel Image Windows" -fore green
DISM /Online /Cleanup-Image /RestoreHealth


Write-Ascii  "MusicLover - Windows CHECK OK" -fore cyan