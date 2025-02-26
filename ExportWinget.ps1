# Load necessary assemblies for WinForms
Add-Type -AssemblyName System.Windows.Forms

# Function to check and adjust execution policy
function Check-ExecutionPolicy {
    $currentPolicy = Get-ExecutionPolicy
    if ($currentPolicy -eq 'Restricted') {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "The current execution policy is set to 'Restricted'.`n`n" +
            "This prevents scripts from running. Would you like to change the execution policy to 'RemoteSigned'?",
            "Execution Policy Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
                [System.Windows.Forms.MessageBox]::Show("Execution policy changed to 'RemoteSigned'.", "Policy Changed")
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to change execution policy: $_", "Error")
                exit
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Operation canceled. Script will exit.", "Canceled")
            exit
        }
    }
}

# Function to show a Save File dialog
function Show-SaveFileDialog {
    # Create a SaveFileDialog object
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog

    # Set properties for the dialog
    $saveFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $saveFileDialog.FileName = "output.json"  # Default file name
    $saveFileDialog.Title = "Save Output File"

    # Show the dialog and get the result
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $saveFileDialog.FileName  # Return the selected file path
    } else {
        return $null  # If canceled, return null
    }
}

# Main script execution
try {
    # Check and adjust execution policy if needed
    Check-ExecutionPolicy

    # Call the function to show the save file dialog
    $outputFilePath = Show-SaveFileDialog

    # Check if a file path was selected
    if ($null -ne $outputFilePath) {
        # Execute the winget export command 
        $command = "winget export -o `"$outputFilePath`""
        
        # Run the command
        Invoke-Expression $command

        # Inform the user that the file has been saved
        [System.Windows.Forms.MessageBox]::Show("Bestand succesvol geexporteerd naar $outputFilePath", "Success")
    } else {
        [System.Windows.Forms.MessageBox]::Show("Operation was canceled.", "Canceled")
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error")
}