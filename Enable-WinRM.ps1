. .\Write-Log.ps1
. .\Download-File.ps1

###########################################################################
$Global:LogSummary = "Enable-WinRM-Summary.txt"
###########################################################################

#Download pre-requisite PS-tools
Download-File -URI "https://live.sysinternals.com/psexec.exe" -ForceOverwrite $true
Download-File -URI "https://live.sysinternals.com/psservice.exe" -ForceOverwrite $true

Remove-Item $Global:LogSummary -ErrorAction SilentlyContinue

Function Enable-WinRM {

    param (
        [Parameter(ValueFromPipeline)]
        [string]$ComputerName
    )
    process {
        $Result = winrm identify -remote:$ComputerName 2>$null
       
        if ($LastExitCode -eq 0) {
            Write-Host "$ComputerName | WinRM already enabled" -ForegroundColor Green
            "$ComputerName | WinRM already enabled" | Write-Log -LogFilename $Global:LogSummary
        } else {
            Write-Host "$ComputerName | Configuring WinRM" -ForegroundColor Yellow
            .\psexec.exe -accepteula \\$ComputerName -s C:\Windows\system32\winrm.cmd quickconfig -quiet
            if ($LastExitCode -eq 0) {
                .\psservice.exe -accepteula \\$ComputerName restart WinRM
                $Result = winrm id -r:$ComputerName 2>$null
                if ($LastExitCode -eq 0) {
                    Write-Host "$ComputerName | WinRM successfully configured" -ForegroundColor Green
                    "$ComputerName | WinRM successfully configured" | Write-Log -LogFilename $Global:LogSummary
                } else {
                    "$ComputerName | WinRM service restart error [FAILURE]" | Write-Log -LogFilename $Global:LogSummary
                    #Exit 1
                }
            } else {
                "$ComputerName | WinRM configure failed [FAILURE]" | Write-Log -LogFilename $Global:LogSummary
                #Exit 1
            }
        }
    }
}