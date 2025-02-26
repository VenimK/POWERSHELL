#3# Schijfopruiming gebruiken  
# Toon een bericht dat aangeeft dat de Schijfopruiming wordt gebruikt
Write-Host "Schijfopruiming wordt gebruikt" -ForegroundColor Yellow  
# Start de Schijfopruiming met de opgegeven sagerun parameter
cleanmgr /sagerun:1 | out-Null  
# Geef een pieptoon weer met ASCII-code 7
Write-Host "$([char]7)"  
# Pauzeer het script voor 5 seconden
Sleep 5  
# Toon een succesbericht dat de Schijfopruiming is voltooid
Write-Host "Schijfopruiming is succesvol voltooid" -ForegroundColor Green  
# Pauzeer het script voor 10 seconden
Sleep 10  