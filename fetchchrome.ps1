$UserName = $env:USERNAME
$Search = ".*"  
$Path = "$Env:systemdrive\Users\$UserName\AppData\Local\Google\Chrome\User Data\Default\History"
if (-not (Test-Path -Path $Path)) {
    Write-Warning "Could not find Chrome History for username: $UserName"
    exit
}
$chromeProcess = Get-Process chrome -ErrorAction SilentlyContinue
if ($chromeProcess) {
    Write-Warning "Please close Chrome browser before running this script"
    exit
}
try {
    $Regex = '(htt(p|s))://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)*?'
    $Value = Get-Content -Path $Path -ErrorAction Stop | 
        Select-String -AllMatches $regex | 
        ForEach-Object { ($_.Matches).Value } | 
        Sort-Object -Unique
    $Value | ForEach-Object {
        $Key = $_
        if ($Key -match $Search) {
            New-Object -TypeName PSObject -Property @{
                User = $UserName
                Browser = 'Chrome'
                DataType = 'History'
                Data = $_
            }
        }
    }
}
catch {
    Write-Error "Error accessing Chrome history: $_"
}