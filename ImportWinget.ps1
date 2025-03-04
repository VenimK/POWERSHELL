# Laad benodigde assemblies voor WinForms
Add-Type -AssemblyName System.Windows.Forms

# Get system language/culture
$systemCulture = (Get-Culture).Name
$translations = @{
    'nl-NL' = @{
        'ExecutionPolicyRestricted' = "Het huidige uitvoeringsbeleid is ingesteld op 'Restricted'.`n`nDit voorkomt dat scripts kunnen worden uitgevoerd. Wilt u het uitvoeringsbeleid wijzigen naar 'RemoteSigned'?"
        'ExecutionPolicyWarning' = "Uitvoeringsbeleid Waarschuwing"
        'PolicyChanged' = "Uitvoeringsbeleid is gewijzigd naar 'RemoteSigned'."
        'PolicyChangedTitle' = "Beleid Gewijzigd"
        'PolicyChangeError' = "Kon het uitvoeringsbeleid niet wijzigen: {0}"
        'ErrorTitle' = "Fout"
        'OperationCanceled' = "Bewerking geannuleerd. Script wordt afgesloten."
        'CanceledTitle' = "Geannuleerd"
        'JsonFiles' = "JSON bestanden (*.json)|*.json|Alle bestanden (*.*)|*.*"
        'SelectWingetFile' = "Selecteer Winget Import Bestand"
        'ConfirmImport' = "Dit zal alle applicaties installeren die in het geselecteerde bestand staan.`n`nWilt u doorgaan?"
        'ConfirmImportTitle' = "Bevestig Import"
        'ImportSuccess' = "Applicaties zijn succesvol geimporteerd van {0}"
        'SuccessTitle' = "Succes"
        'ImportCanceled' = "Import bewerking is geannuleerd."
        'NoFileSelected' = "Geen bestand geselecteerd."
        'ErrorOccurred' = "Er is een fout opgetreden: {0}"
        'DutchLanguageNotAvailable' = "Nederlands taalpakket is niet geïnstalleerd in Windows.`n`nWinget zal in het Engels worden weergegeven.`n`nInstalleer het Nederlandse taalpakket via Windows-instellingen > Tijd en taal > Taal om Winget in het Nederlands te gebruiken."
        'LanguageWarning' = "Taal Waarschuwing"
    }
    'en-US' = @{
        'ExecutionPolicyRestricted' = "The current execution policy is set to 'Restricted'.`n`nThis prevents scripts from running. Would you like to change the execution policy to 'RemoteSigned'?"
        'ExecutionPolicyWarning' = "Execution Policy Warning"
        'PolicyChanged' = "Execution policy changed to 'RemoteSigned'."
        'PolicyChangedTitle' = "Policy Changed"
        'PolicyChangeError' = "Failed to change execution policy: {0}"
        'ErrorTitle' = "Error"
        'OperationCanceled' = "Operation canceled. Script will exit."
        'CanceledTitle' = "Canceled"
        'JsonFiles' = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        'SelectWingetFile' = "Select Winget Import File"
        'ConfirmImport' = "This will install all applications listed in the selected file.`n`nDo you want to continue?"
        'ConfirmImportTitle' = "Confirm Import"
        'ImportSuccess' = "Applications have been successfully imported from {0}"
        'SuccessTitle' = "Success"
        'ImportCanceled' = "Import operation was canceled."
        'NoFileSelected' = "No file was selected."
        'ErrorOccurred' = "An error occurred: {0}"
    }
    'de-DE' = @{
        'ExecutionPolicyRestricted' = "Die aktuelle Ausfuehrungsrichtlinie ist auf 'Restricted' festgelegt.`n`nDies verhindert die Ausfuehrung von Skripten. Moechten Sie die Ausfuehrungsrichtlinie auf 'RemoteSigned' aendern?"
        'ExecutionPolicyWarning' = "Ausfuehrungsrichtlinie Warnung"
        'PolicyChanged' = "Ausfuehrungsrichtlinie wurde auf 'RemoteSigned' geaendert."
        'PolicyChangedTitle' = "Richtlinie Geaendert"
        'PolicyChangeError' = "Fehler beim Aendern der Ausfuehrungsrichtlinie: {0}"
        'ErrorTitle' = "Fehler"
        'OperationCanceled' = "Vorgang abgebrochen. Skript wird beendet."
        'CanceledTitle' = "Abgebrochen"
        'JsonFiles' = "JSON Dateien (*.json)|*.json|Alle Dateien (*.*)|*.*"
        'SelectWingetFile' = "Winget Import Datei auswaehlen"
        'ConfirmImport' = "Dies installiert alle Anwendungen aus der ausgewaehlten Datei.`n`nMoechten Sie fortfahren?"
        'ConfirmImportTitle' = "Import Bestaetigen"
        'ImportSuccess' = "Anwendungen wurden erfolgreich von {0} importiert"
        'SuccessTitle' = "Erfolg"
        'ImportCanceled' = "Import-Vorgang wurde abgebrochen."
        'NoFileSelected' = "Keine Datei ausgewaehlt."
        'ErrorOccurred' = "Ein Fehler ist aufgetreten: {0}"
    }
}

# Default to Dutch if system language is not supported
if (-not $translations.ContainsKey($systemCulture)) {
    $systemCulture = 'nl-NL'
}

# Get translation function
function Get-Translation {
    param (
        [string]$Key,
        [array]$Parameters = @()
    )
    
    $text = $translations[$systemCulture][$Key]
    if ($Parameters.Count -gt 0) {
        $text = [string]::Format($text, $Parameters)
    }
    return $text
}

# Function to check and adjust execution policy
function Check-ExecutionPolicy {
    $currentPolicy = Get-ExecutionPolicy
    if ($currentPolicy -eq 'Restricted') {
        $result = [System.Windows.Forms.MessageBox]::Show(
            (Get-Translation 'ExecutionPolicyRestricted'),
            (Get-Translation 'ExecutionPolicyWarning'),
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
                [System.Windows.Forms.MessageBox]::Show(
                    (Get-Translation 'PolicyChanged'),
                    (Get-Translation 'PolicyChangedTitle')
                )
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    (Get-Translation 'PolicyChangeError' $_),
                    (Get-Translation 'ErrorTitle')
                )
                exit
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Translation 'OperationCanceled'),
                (Get-Translation 'CanceledTitle')
            )
            exit
        }
    }
}

# Function to check if Dutch language is available
function Test-DutchLanguage {
    $dutchLanguage = Get-WinSystemLocale | Where-Object { $_.Name -eq 'nl-NL' }
    if (-not $dutchLanguage) {
        [System.Windows.Forms.MessageBox]::Show(
            (Get-Translation 'DutchLanguageNotAvailable'),
            (Get-Translation 'LanguageWarning'),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return $false
    }
    return $true
}

# Function to show an Open File dialog
function Show-OpenFileDialog {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = (Get-Translation 'JsonFiles')
    $openFileDialog.Title = (Get-Translation 'SelectWingetFile')
    $openFileDialog.InitialDirectory = [System.Environment]::GetFolderPath('Desktop')

    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    } else {
        return $null
    }
}

# Main script execution
try {
    Check-ExecutionPolicy
    Test-DutchLanguage
    $inputFilePath = Show-OpenFileDialog

    if ($null -ne $inputFilePath) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            (Get-Translation 'ConfirmImport'),
            (Get-Translation 'ConfirmImportTitle'),
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Set Windows language preference to Dutch before running winget
            $env:LANG = "nl-NL"
            [System.Threading.Thread]::CurrentThread.CurrentUICulture = 'nl-NL'
            [System.Threading.Thread]::CurrentThread.CurrentCulture = 'nl-NL'
            
            $command = "winget import -i `"$inputFilePath`""
            Invoke-Expression $command

            [System.Windows.Forms.MessageBox]::Show(
                (Get-Translation 'ImportSuccess' $inputFilePath),
                (Get-Translation 'SuccessTitle')
            )
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Translation 'ImportCanceled'),
                (Get-Translation 'CanceledTitle')
            )
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            (Get-Translation 'NoFileSelected'),
            (Get-Translation 'CanceledTitle')
        )
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        (Get-Translation 'ErrorOccurred' $_),
        (Get-Translation 'ErrorTitle')
    )
}
