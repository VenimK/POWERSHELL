function Log-Message {
    param (
        [string]$Message
    )
    # Write messages to a log file or the console. 
    # For simplicity, we're just writing to the console here.
    Write-Host $Message -ForegroundColor White
}

function Install-WingetPackage {
    param (
        [string]$PackageName,
        [string]$PackageDisplayName
    )

    Log-Message "Checking installation status of $PackageDisplayName."
    Write-Host "Checking installation status of $PackageDisplayName..." -ForegroundColor Yellow

    # Check if the package is installed
    $installedPackage = winget list --exact --id $PackageName -q

    if ($installedPackage) {
        Log-Message "$PackageDisplayName is already installed. Checking for updates."

        # Check for updates
        $updateAvailable = winget upgrade --id $PackageName --silent -q

        if ($updateAvailable) {
            Log-Message "An update for $PackageDisplayName is available. Updating..."
            Write-Host "An update for $PackageDisplayName is available. Updating..." -ForegroundColor Yellow

            try {
                winget upgrade $PackageName --silent --accept-source-agreements --accept-package-agreements
                Log-Message "$PackageDisplayName updated successfully."
                Write-Host "$PackageDisplayName updated successfully." -ForegroundColor Green
            } catch {
                $errorMessage = "Failed to update package ${PackageDisplayName}: $($_.ToString())"
                Log-Message $errorMessage
                Write-Output $errorMessage -ForegroundColor Red
            }
        } else {
            Log-Message "$PackageDisplayName is already up to date. No action needed."
            Write-Host "$PackageDisplayName is already up to date. No action needed." -ForegroundColor Cyan
        }
    } else {
        Log-Message "Starting installation of $PackageDisplayName."
        Write-Host "Installing $PackageDisplayName..." -ForegroundColor Cyan

        try {
            # Run the installation command
            $installResult = winget install $PackageName --silent --accept-source-agreements --accept-package-agreements

            if ($installResult -like "*successfully*") {
                Log-Message "$PackageDisplayName installed successfully."
                Write-Host "$PackageDisplayName installed successfully." -ForegroundColor Green
            } else {
                throw "Installation command for $PackageDisplayName did not complete as expected."
            }
        } catch {
            $errorMessage = "Failed to install package ${PackageDisplayName}: $($_.ToString())"
            Log-Message $errorMessage
            Write-Output $errorMessage -ForegroundColor Red
        }
    }
}

# Example usage to install 7-Zip
# First, check the available packages for 7-Zip to get the correct ID
$packageId = (winget search 7-Zip | Select-String -Pattern '7-Zip' | Select-Object -First 1).Line.Split('|')[0].Trim()

# Call the Install-WingetPackage function with the correct ID.
if ($packageId) {
    $packageDisplayName = "7Zip"
    Install-WingetPackage -PackageName $packageId -PackageDisplayName $packageDisplayName
} else {
    Write-Host "7Zip package not found." -ForegroundColor Red
}
