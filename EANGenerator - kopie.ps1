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
    
    try {
        # Create Word COM object
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        # Open the document
        Write-Host "Opening document: $documentPath"
        $doc = $word.Documents.Open($documentPath)
        
        # Find all potential EAN numbers (12 or 13 digits, possibly with spaces or hyphens)
        $content = $doc.Content
        $pattern = "\b\d[\d\s-]{10,15}\d\b"
        
        $found = $false
        $selection = $word.Selection
        $selection.Find.ClearFormatting()
        
        # Configure find parameters
        $selection.Find.Forward = $true
        $selection.Find.Text = $pattern
        $selection.Find.MatchWildcards = $true
        
        # Process each found number
        While ($selection.Find.Execute()) {
            $found = $true
            $originalText = $selection.Text
            
            # Clean and validate the number
            $eanNumber = Validate-EANNumber $originalText
            if ($eanNumber) {
                # Calculate checksum and get full EAN
                $checksum = Calculate-EANChecksum $eanNumber
                $fullEAN = "$eanNumber$checksum"
                
                # Generate and download barcode
                $barcodeUrl = "https://barcodeapi.org/api/ean13/$fullEAN"
                $tempPath = [System.IO.Path]::GetTempFileName() + ".png"
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($barcodeUrl, $tempPath)
                
                # Move to end of selection and insert barcode
                $selection.Collapse(0) # Collapse to end
                $selection.InsertBreak(6) # Line break
                $doc.InlineShapes.AddPicture($tempPath)
                $selection.InsertBreak(6) # Line break
                
                # Clean up temp file
                Remove-Item $tempPath -Force
                
                Write-Host "Processed EAN: $fullEAN"
            }
        }
        
        if (-not $found) {
            Write-Host "No valid EAN numbers found in the document." -ForegroundColor Yellow
        }
        
        # Save and close
        $doc.Save()
        $doc.Close()
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        
        Write-Host "`nDocument processing completed successfully!" -ForegroundColor Green
        
    } catch {
        Write-Host "Error processing document: $($_.Exception.Message)" -ForegroundColor Red
        if ($doc) { $doc.Close($false) }
        if ($word) { $word.Quit() }
    }
}

try {
    Clear-Host
    Write-Host "Word Document EAN Barcode Generator" -ForegroundColor Cyan
    Write-Host "=================================" -ForegroundColor Cyan
    
    # Ask for Word document path
    $docPath = Read-Host "Enter the path to your Word document"
    
    # Validate file exists and is a Word document
    if (-not (Test-Path $docPath)) {
        throw "Document not found: $docPath"
    }
    if (-not $docPath.ToLower().EndsWith('.doc') -and -not $docPath.ToLower().EndsWith('.docx')) {
        throw "File must be a Word document (.doc or .docx)"
    }
    
    # Process the document
    Process-WordDocument $docPath
    
} catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
