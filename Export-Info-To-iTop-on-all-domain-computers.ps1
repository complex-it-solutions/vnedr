. .\Export-To-Itop.ps1

###########################################################################
$OrganizationName = "Экопродукт"
$OrganizationLocation = "Экопродукт Москва"
###########################################################################
$LogSummary = "Export-To-iTop-Summary.txt"
$KE_List = "Export-To-iTop-KE.txt"
$KE_MB_List = "Export-To-iTop-Motherboard.txt"
###########################################################################

Remove-Item $LogSummary -ErrorAction SilentlyContinue

Write-Host "Started collecting full KE info"
"Started collecting full KE info" | Write-Log -LogFilename $LogSummary

$Computers = (Get-ADComputer -Filter {Enabled -eq "True" -and OperatingSystem -notlike "*Server*"} -Properties OperatingSystem).Name | Sort-Object
$Computers | ForEach-Object {
    if (Test-Connection $_ -Count 1 -Quiet -ErrorAction SilentlyContinue) {
        Test-WSMan -ComputerName $_ -ErrorAction SilentlyContinue | Out-Null
        if ($?) {
            Write-Host $_ -ForegroundColor Green            
            Export-KEToItop -ComputerName $_ -OrganizationName $OrganizationName -OrganizationLocation $OrganizationLocation
            "$_ | Data collected successfully" | Write-Log -LogFilename $LogSummary
        } else {
            Write-Host $_ -ForegroundColor Red
            "$_ | Can't connect through WinRM [FAILURE]" | Write-Log -LogFilename $LogSummary
        }
    } else {
        Write-Host $_ -ForegroundColor Red
        "$_ | Can't ping [ERROR]" | Write-Log -LogFilename $LogSummary
    }
} | ConvertTo-Csv -NoTypeInformation | Out-File $KE_List -Encoding utf8

Write-Host "Finished collecting full KE info"
"Finished collecting full KE info" | Write-Log -LogFilename $LogSummary
Write-Host "------------------------------------------------------------"
Write-Host ""
Write-Host "Started collecting short KE info"
"Started collecting short KE info" | Write-Log -LogFilename $LogSummary

$computers = (Get-ADComputer -Filter {Enabled -eq "True" -and OperatingSystem -notlike "*Server*"} -Properties OperatingSystem).Name | Sort-Object
$computers | ForEach-Object {
    if (Test-Connection $_ -Count 1 -Quiet -ErrorAction SilentlyContinue) {
        Test-WSMan -ComputerName $_ -ErrorAction SilentlyContinue | Out-Null
        if ($?) {
            Write-Host $_ -ForegroundColor Green
            Export-MotherboardToItop -ComputerName $_
            "$_ | Data collected successfully" | Write-Log -LogFilename $LogSummary
        } else {
            Write-Host $_ -ForegroundColor Red
            "$_ | Can't connect through WinRM [FAILURE]" | Write-Log -LogFilename $LogSummary
        }
    } else {
        Write-Host $_ -ForegroundColor Red
        "$_ | Can't ping [ERROR]" | Write-Log -LogFilename $LogSummary
    }
} | ConvertTo-Csv -NoTypeInformation | Out-File $KE_MB_List -Encoding utf8

Write-Host "Finished collecting short KE info"
"Finished collecting short KE info" | Write-Log -LogFilename $LogSummary
Write-Host "------------------------------------------------------------"

Invoke-Item .\$LogSummary