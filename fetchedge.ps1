$UserName = "VenimK" 
$UserName        
$Path = "$Env:systemdrive\Users\$UserName\AppData\Local\Microsoft\Edge\User Data\Default\History" 
        if (-not (Test-Path -Path $Path)) { 
            Write-Verbose "[!] Could not find Edge History for username: $UserName" 
        } 
        $Regex = '(htt(p|s))://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)*?' 
        $Value = Get-Content -Path "$Env:systemdrive\Users\$UserName\AppData\Local\Microsoft\Edge\User Data\Default\History"|Select-String -AllMatches $regex |% {($_.Matches).Value} |Sort -Unique 
        $Value | ForEach-Object { 
            $Key = $_ 
            if ($Key -match $Search){ 
                New-Object -TypeName PSObject -Property @{ 
                    User = $UserName 
                    Browser = 'Edge' 
                    DataType = 'History' 
                    Data = $_ 
                } 
            } 
        } 