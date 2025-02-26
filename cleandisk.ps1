#3# Using Disk cleanup Tool  
# Display a message indicating the usage of the Disk Cleanup tool
write-Host "Using Disk cleanup Tool" -ForegroundColor Yellow  
# Run the Disk Cleanup tool with the specified sagerun parameter
cleanmgr /sagerun:1 | out-Null  
# Emit a beep sound using ASCII code 7
Write-Host "$([char]7)"  
# Pause the script for 5 seconds
Sleep 5  
# Display a success message indicating that Disk Cleanup was successfully done
write-Host "Disk Cleanup Successfully done" -ForegroundColor Green  
# Pause the script for 10 seconds
Sleep 10  