Function Write-Log {
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline)]
        [string]$LogEntry,
        [string]$LogPath = ".",
        [string]$LogFilename = "log.txt"
    )

    Process {
        Add-Content "$LogPath\$LogFilename" "[$(Get-Date -Format "dd.MM.yyyy HH:mm:ss")] - $LogEntry"
    }
}