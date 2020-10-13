Function New-RandomPassword {
    param (
        [int]$Length=8
    )
    $RandomPassword = (-join ([char[]](48..57) + [char[]](65..90) + [char[]](97..122) | Get-Random -Count $Length))
    Write-Output $RandomPassword
}

Function New-LocalUserLegacy {

    param(
        [Parameter(ValueFromPipeline)]
        [string]$ComputerName = $env:COMPUTERNAME,
        [Parameter(Mandatory=$true)]
        [string]$UserName,
        [string]$GroupName = "Пользователи",
        [Parameter(Mandatory=$true)]
        [string]$Password,
        [string]$Description = $UserName,
        [switch]$BecomeAdmin
    )

    try {
        $Computer = [ADSI]"WinNT://$ComputerName"
        if ((($Computer.Children | Where-Object {$_.SchemaClassName  -eq 'user'} | Where-Object {$_.Path -like "*$Username*"}).Path).Count -eq 0) {
            $User = $Computer.Create("User", $UserName)
            $User.SetPassword($Password)
            $User.SetInfo()
            $User.Description = $Description
            $User.SetInfo()
            $User = [ADSI]("WinNT://$ComputerName/$Username")
            $User.UserFlags.Value = $User.UserFlags.Value -bor 0x10000
            $User.CommitChanges()
            if ($BecomeAdmin -eq $true) {
                $GroupName = "Администраторы"
            } 
            $Group = [ADSI]("WinNT://$ComputerName/$GroupName")
            $Group.PSBase.Invoke("Add",$User.PSBase.Path)
            Write-Host "[$ComputerName] Пользователь $Username добавлен в группу $GroupName"
            Write-Host "Пароль: $Password"
        } else {
            Write-Host "[$ComputerName] Пользователь $Username уже существует"
        }
    }
    catch {
        Write-Host "Что-то пошло не так" -ForegroundColor Red
    }
}
