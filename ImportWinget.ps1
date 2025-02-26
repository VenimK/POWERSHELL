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

# Function to show an Open File dialog
function Show-OpenFileDialog {
    # Create an OpenFileDialog object
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog

    # Set properties for the dialog
    $openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $openFileDialog.Title = "Select Winget Import File"
    $openFileDialog.InitialDirectory = [System.Environment]::GetFolderPath('Desktop')

    # Show the dialog and get the result
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName  # Return the selected file path
    } else {
        return $null  # If canceled, return null
    }
}

# Main script execution
try {
    # Check and adjust execution policy if needed
    Check-ExecutionPolicy

    # Call the function to show the open file dialog
    $inputFilePath = Show-OpenFileDialog

    # Check if a file path was selected
    if ($null -ne $inputFilePath) {
        # Show confirmation dialog
        $result = [System.Windows.Forms.MessageBox]::Show(
            "This will install all applications listed in the selected file.`n`nDo you want to continue?",
            "Confirm Import",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Execute the winget import command
            $command = "winget import -i `"$inputFilePath`""
            
            # Run the command
            Invoke-Expression $command

            # Inform the user that the import has completed
            [System.Windows.Forms.MessageBox]::Show("Applications have been successfully imported from $inputFilePath", "Success")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Import operation was canceled.", "Canceled")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No file was selected.", "Canceled")
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error")
}
