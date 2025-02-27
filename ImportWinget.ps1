# Laad benodigde assemblies voor WinForms
Add-Type -AssemblyName System.Windows.Forms

# Functie om het uitvoeringsbeleid te controleren en aan te passen
function Controleer-UitvoerBeleid {
    $huidigBeleid = Get-ExecutionPolicy
    if ($huidigBeleid -eq 'Restricted') {
        $resultaat = [System.Windows.Forms.MessageBox]::Show(
            "Het huidige uitvoeringsbeleid is ingesteld op 'Restricted'.`n`n" +
            "Dit voorkomt dat scripts kunnen worden uitgevoerd. Wilt u het uitvoeringsbeleid wijzigen naar 'RemoteSigned'?",
            "Uitvoeringsbeleid Waarschuwing",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($resultaat -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
                [System.Windows.Forms.MessageBox]::Show("Uitvoeringsbeleid is gewijzigd naar 'RemoteSigned'.", "Beleid Gewijzigd")
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Kon het uitvoeringsbeleid niet wijzigen: $_", "Fout")
                exit
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Bewerking geannuleerd. Script wordt afgesloten.", "Geannuleerd")
            exit
        }
    }
}

# Functie om een Bestand Openen dialoog te tonen
function Toon-BestandKiezer {
    # Maak een OpenFileDialog object
    $bestandKiezer = New-Object System.Windows.Forms.OpenFileDialog

    # Stel eigenschappen in voor de dialoog
    $bestandKiezer.Filter = "JSON bestanden (*.json)|*.json|Alle bestanden (*.*)|*.*"
    $bestandKiezer.Title = "Selecteer Winget Import Bestand"
    $bestandKiezer.InitialDirectory = [System.Environment]::GetFolderPath('Desktop')

    # Toon de dialoog en krijg het resultaat
    if ($bestandKiezer.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $bestandKiezer.FileName  # Geef het geselecteerde bestandspad terug
    } else {
        return $null  # Als geannuleerd, geef null terug
    }
}

# Hoofdscript uitvoering
try {
    # Controleer en pas uitvoeringsbeleid aan indien nodig
    Controleer-UitvoerBeleid

    # Roep de functie aan om de bestandkiezer te tonen
    $invoerBestandPad = Toon-BestandKiezer

    # Controleer of er een bestandspad is geselecteerd
    if ($null -ne $invoerBestandPad) {
        # Toon bevestigingsdialoog
        $resultaat = [System.Windows.Forms.MessageBox]::Show(
            "Dit zal alle applicaties installeren die in het geselecteerde bestand staan.`n`nWilt u doorgaan?",
            "Bevestig Import",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($resultaat -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Voer het winget import commando uit
            $commando = "winget import -i `"$invoerBestandPad`""
            
            # Voer het commando uit
            Invoke-Expression $commando

            # Informeer de gebruiker dat de import is voltooid
            [System.Windows.Forms.MessageBox]::Show("Applicaties zijn succesvol geimporteerd van $invoerBestandPad", "Succes")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Import bewerking is geannuleerd.", "Geannuleerd")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Geen bestand geselecteerd.", "Geannuleerd")
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show("Er is een fout opgetreden: $_", "Fout")
}
