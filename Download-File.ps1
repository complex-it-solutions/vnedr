function Download-File {
    Param (
        [string]$URI, #URI файла
        [string]$Path = ".", #папка для сохранения
        [boolean]$ForceOverwrite = $False #перезаписать существующий файл?
    )
        $Filename = switch -Regex ($URI) {"^\S+\/(?<Name>\S+)" {$Matches.Name}}
        if (!(Test-Path $Path\$Filename) -or ($ForceOverwrite)) {
            if (Test-Path $Path) {
                Invoke-WebRequest -Uri $URI -OutFile $Path\$Filename
            } else {
                $PromptFolder = Read-Host -Prompt 'Cannot find destination folder. Create it? (Y/N)'
                if ($PromptFolder -eq "y" -or $PromptFolder -eq "Y") {
                    New-Item -ItemType Directory -Path $Path
                    Invoke-WebRequest -Uri $URI -OutFile $Path\$Filename
                }
            }
        } else {
            $PromptFile = Read-Host 'File already exists, Overwrite it? (Y/N)'
            if ($PromptFile -eq "y" -or $PromptFolder -eq "Y") {
                Invoke-WebRequest -Uri $URI -OutFile $Path\$Filename
            }
        }
}
