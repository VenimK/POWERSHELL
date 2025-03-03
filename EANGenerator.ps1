# =================================================================
# Word Document EAN Barcode Generator
# =================================================================
# This script reads EAN numbers from a Word document and adds their
# corresponding barcodes next to them.
# =================================================================

function Calculate-EANChecksum {
    param (
        [string]$digits
    )
    
    # Controleer of we precies 12 cijfers hebben
    if ($digits.Length -ne 12 -or $digits -notmatch '^\d{12}$') {
        throw "EAN nummer moet exact 12 cijfers bevatten"
    }
    
    $sum = 0
    for ($i = 0; $i -lt 12; $i++) {
        $digit = [int]::Parse($digits[$i].ToString())
        # Even posities x 3, oneven posities x 1
        $multiplier = if ($i % 2 -eq 0) { 1 } else { 3 }
        $sum += $digit * $multiplier
    }
    
    $checksum = (10 - ($sum % 10)) % 10
    return $checksum
}

function Validate-EANNumber {
    param (
        [string]$number
    )
    
    # Verwijder spaties en streepjes
    $number = $number -replace '[-\s]', ''
    
    # Controleer of het nummer alleen cijfers bevat
    if ($number -notmatch '^\d+$') {
        Write-Host "Fout: EAN nummer mag alleen cijfers bevatten" -ForegroundColor Red
        return $null
    }
    
    # Als het nummer 13 cijfers heeft, controleer het controlegetal
    if ($number.Length -eq 13) {
        $expectedChecksum = [int]::Parse($number[12].ToString())
        $calculatedChecksum = Calculate-EANChecksum($number.Substring(0, 12))
        
        if ($expectedChecksum -ne $calculatedChecksum) {
            Write-Host "Waarschuwing: Ongeldig controlegetal. Wordt gecorrigeerd." -ForegroundColor Yellow
            return $number.Substring(0, 12)
        }
        return $number.Substring(0, 12)
    }
    
    # Als het nummer 12 cijfers heeft, gebruik het direct
    if ($number.Length -eq 12) {
        return $number
    }
    
    Write-Host "Fout: EAN nummer moet 12 of 13 cijfers bevatten" -ForegroundColor Red
    return $null
}

function Process-WordDocument {
    param (
        [string]$documentPath
    )
    
    $word = $null
    $sourceDoc = $null
    $newDoc = $null
    $tempPath = $null
    
    # Force close any existing Word processes at start
    Write-Host "Cleaning up any existing Word processes..." -ForegroundColor Yellow
    Get-Process | Where-Object { $_.ProcessName -eq "WINWORD" } | ForEach-Object { 
        try {
            $_.CloseMainWindow()
            Start-Sleep -Seconds 1
            if (!$_.HasExited) { $_.Kill() }
        } catch { }
    }
    
    try {
        # Create Word COM object with timeout
        Write-Host "Starting Word..." -ForegroundColor Cyan
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        # Create new filename
        $directory = [System.IO.Path]::GetDirectoryName($documentPath)
        $filename = [System.IO.Path]::GetFileNameWithoutExtension($documentPath)
        $extension = [System.IO.Path]::GetExtension($documentPath)
        $newPath = Join-Path $directory ($filename + "_met_barcodes" + $extension)
        
        # Create temporary path
        $tempPath = [System.IO.Path]::GetTempFileName()
        Remove-Item $tempPath -Force
        $tempPath = $tempPath + $extension
        
        # Open source document with timeout
        Write-Host "Opening document: $documentPath" -ForegroundColor Cyan
        $sourceDoc = $word.Documents.Open($documentPath)
        Start-Sleep -Seconds 2  # Give Word time to fully open the document
        
        # Create new document
        $newDoc = $word.Documents.Add()
        
        $found = $false
        Write-Host "`nAnalyzing document content..." -ForegroundColor Cyan
        
        # First, collect all EAN numbers and their locations
        $eanLocations = @()
        $range = $sourceDoc.Content
        $paragraphs = $range.Paragraphs
        
        Write-Host "`nCollecting EAN numbers..." -ForegroundColor Yellow
        
        for ($i = 1; $i -le $paragraphs.Count; $i++) {
            $paragraph = $paragraphs.Item($i)
            $text = $paragraph.Range.Text.Trim()
            
            # Extract numbers
            $numbers = $text -replace '[^\d]', ''
            if ($numbers.Length -ge 12) {
                $offset = 0
                while ($offset + 12 -le $numbers.Length) {
                    $potentialEAN = $numbers.Substring($offset, 12)
                    
                    # Validate the number
                    $eanNumber = Validate-EANNumber $potentialEAN
                    if ($eanNumber) {
                        $found = $true
                        $eanLocations += @{
                            EAN = $eanNumber
                            Text = $text
                        }
                        break  # Only process first EAN in paragraph
                    }
                    $offset += 12
                }
            }
        }
        
        Write-Host "Found $($eanLocations.Count) valid EAN numbers" -ForegroundColor Green
        
        # Process each EAN and add to new document
        foreach ($loc in $eanLocations) {
            $eanNumber = $loc.EAN
            
            # Calculate checksum and get full EAN
            $checksum = Calculate-EANChecksum $eanNumber
            $fullEAN = "$eanNumber$checksum"
            
            Write-Host "Processing EAN: $fullEAN" -ForegroundColor Green
            
            # Generate and download barcode
            Write-Host "Generating barcode..." -ForegroundColor Cyan
            $barcodeUrl = "https://barcodeapi.org/api/ean13/$fullEAN"
            $tempBarcodeFile = [System.IO.Path]::GetTempFileName() + ".png"
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($barcodeUrl, $tempBarcodeFile)
            
            Write-Host "Inserting barcode into document..." -ForegroundColor Cyan
            
            # Get end of document
            $range = $newDoc.Range($newDoc.Content.End - 1, $newDoc.Content.End - 1)
            
            # Insert barcode
            $shape = $newDoc.InlineShapes.AddPicture($tempBarcodeFile, $False, $True, $range)
            
            # Add EAN number below barcode
            $range.InsertAfter("`n$fullEAN`n`n")
            $range.Paragraphs.Alignment = 1  # Center alignment
            
            # Clean up temp file
            Remove-Item $tempBarcodeFile -Force
            
            Write-Host "Barcode inserted successfully" -ForegroundColor Green
        }
        
        if (-not $found) {
            Write-Host "`nNo valid EAN numbers found in the document." -ForegroundColor Yellow
            Write-Host "Make sure your document contains EAN numbers that are 12 or 13 digits long." -ForegroundColor Yellow
            Write-Host "Numbers can contain spaces, dots, or hyphens, but must have the correct number of digits." -ForegroundColor Yellow
            return
        }
        
        # Save new document to temp location first
        Write-Host "`nSaving document to temporary location..." -ForegroundColor Cyan
        
        # Word constants
        $wdFormatDocumentDefault = 16
        
        # Save with explicit format to temp location
        $newDoc.SaveAs2($tempPath, $wdFormatDocumentDefault)
        
        # Close documents with timeout
        Write-Host "Closing documents..." -ForegroundColor Yellow
        if ($sourceDoc) { 
            $sourceDoc.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sourceDoc) | Out-Null
            $sourceDoc = $null
        }
        if ($newDoc) { 
            $newDoc.Close()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($newDoc) | Out-Null
            $newDoc = $null
        }
        
        # Close Word with timeout
        Write-Host "Closing Word..." -ForegroundColor Yellow
        if ($word) {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            $word = $null
        }
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Final cleanup of any remaining Word processes
        Get-Process | Where-Object { $_.ProcessName -eq "WINWORD" } | ForEach-Object { 
            try {
                $_.CloseMainWindow()
                Start-Sleep -Seconds 1
                if (!$_.HasExited) { $_.Kill() }
            } catch { }
        }
        
        # Move the temp file to final location
        Write-Host "`nMoving file to final location..." -ForegroundColor Yellow
        Move-Item -Path $tempPath -Destination $newPath -Force
        
        Write-Host "`nDocument processing completed successfully!" -ForegroundColor Green
        Write-Host "New document saved as: $newPath" -ForegroundColor Green
        
    } catch {
        Write-Host "Error processing document: $($_.Exception.Message)" -ForegroundColor Red
        
        # Emergency cleanup
        if ($sourceDoc) { 
            try { 
                $sourceDoc.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sourceDoc) | Out-Null
            } catch { }
        }
        if ($newDoc) { 
            try { 
                $newDoc.Close()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($newDoc) | Out-Null
            } catch { }
        }
        if ($word) { 
            try { 
                $word.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            } catch { }
        }
        
        # Clean up temp file if it exists
        if ($tempPath -and (Test-Path $tempPath)) {
            try {
                Remove-Item $tempPath -Force
            } catch { }
        }
        
        # Force close any remaining Word processes
        Get-Process | Where-Object { $_.ProcessName -eq "WINWORD" } | ForEach-Object { 
            try { 
                $_.CloseMainWindow()
                Start-Sleep -Seconds 1
                if (!$_.HasExited) { $_.Kill() }
            } catch { }
        }
        
        throw
    }
}

function Show-FilePickerDialog {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $dialog.Filter = "Word Documents (*.doc;*.docx)|*.doc;*.docx|All Files (*.*)|*.*"
    $dialog.FilterIndex = 1
    $dialog.Title = "Select a Word document containing EAN numbers"
    
    if ($dialog.ShowDialog() -eq 'OK') {
        return $dialog.FileName
    }
    return $null
}

try {
    Clear-Host
    Write-Host "Word Document EAN Barcode Generator" -ForegroundColor Cyan
    Write-Host "=================================" -ForegroundColor Cyan
    
    # Ask for Word document path
    Write-Host "Select a Word document containing EAN numbers..." -ForegroundColor Cyan
    $documentPath = Show-FilePickerDialog

    if (-not $documentPath) {
        Write-Host "No file selected. Exiting..." -ForegroundColor Yellow
        return
    }

    if (-not (Test-Path $documentPath)) {
        Write-Host "Selected file does not exist. Exiting..." -ForegroundColor Red
        return
    }

    Process-WordDocument $documentPath
    
} catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
