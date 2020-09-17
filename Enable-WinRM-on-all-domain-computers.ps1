$ErrorActionPreference = "SilentlyContinue"

. .\Enable-WinRM.ps1

$Computers = (Get-ADComputer -Filter {(Enabled -eq "True") -and (OperatingSystem -like "*windows*") -and (OperatingSystem -notlike "*server*")} | Sort-Object Name).Name

foreach ($Computer in $Computers){
    if (Test-Connection -Count 1 -ErrorAction SilentlyContinue -ComputerName $Computer) {
        Enable-WinRM -ComputerName $Computer
    } else {
        Write-Host $Computer -ForegroundColor Red
    }
}

Invoke-Item $Global:LogSummary