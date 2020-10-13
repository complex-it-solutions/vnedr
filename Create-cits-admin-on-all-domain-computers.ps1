. .\Write-Log.ps1
. .\New-LocalUser.ps1

$Computers = (Get-ADComputer -Filter {(Enabled -eq "True") -and (OperatingSystem -like "*windows*") -and (OperatingSystem -notlike "*server*")} | Sort-Object Name).Name

foreach ($Computer in $Computers){
    if (Test-Connection -Count 1 -ErrorAction SilentlyContinue -ComputerName $Computer) {
        Write-Host $Computer -ForegroundColor Green
        New-LocalUserLegacy -ComputerName $Computer -UserName "localadmin" -Password "P@$$w0rd" -BecomeAdmin -Description "Локальный администратор localadmin"
    } else {
        Write-Host $Computer -ForegroundColor Red
    }
}