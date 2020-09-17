. .\Write-Log.ps1

#Requires -Version 2.0
function Export-KEToItop {
    <#
    .SYNOPSIS
        Эта функция используется для сбора информации о ПК, которую затем можно импортировать в iTop как конфигурационную единицу (КЕ).
        Рекомендуется использовать вместе с ConvertTo-CSV, результатом будет CSV-список.
    
    .DESCRIPTION
        Собирается информация о начинке компьютера. Такая как: материнская плата, процессор, оперативная память, версия операционной системы.
        Фунцию можно использовать на нескольких компьютерах.
    
    .PARAMETER ComputerType
        Тип компьютера: настольный или ноутбук. По умолчанию, настольный.
    
    .PARAMETER Criticality
        Критичность компьютера: Низкая, средняя или высокая. По умолчанию, низкая.
    
    .PARAMETER Description
        Описание компьютера. Разумно указывать тут данные для подключения через TeamViewer (или им подобных) и информацию о дополнительных учётных записях администраторов.
    
    .PARAMETER OrganizationName
        Название организации, символ-в-символ, как в айтопе. Например "БИЛ (БрэндИмпортЛоджистик)". По умолчанию, <Название компании>.
    
    .PARAMETER OrganizationLocation
        Расположение организации, символ-в-символ, как в айтопе. Например "Офис БИЛ Химки". По умолчанию, <Расположение компании>.
    
    .PARAMETER UseWMI
        Использовать WMI вместо CIM. Необходимо для работы с Powershell 2.0. По умолчанию, CIM
    
    .EXAMPLE
        PS C:\> Export-KEToItop
        Собрать информацию на локальном компьютере
        
    .EXAMPLE
        PS C:\> Export-KEToItop -ComputerName MANAGER01
        Собрать информацию на удалённом компьютере MANAGER01
        
    .EXAMPLE
        PS C:\> Export-KEToItop -ComputerName MANAGER01 -UseWMI
        Собрать информацию на удалённом компьютере MANAGER01 используя WMI.
    
    .EXAMPLE
        PS C:\> 
        $computers = (Get-ADComputer -SearchBase "OU=Moscow,OU=Workstations,DC=mege,DC=ru" -Filter {Enabled -eq "True" -and Name -like "*NOTE*"}).Name
        $computers | foreach {
            if (Test-Connection $_ -Count 1 -Quiet) {
                Test-WSMan -ComputerName $_ | Out-Null
                if ($?) {
                    Export-KEToItop -ComputerName $_ -OrganizationName "Миг Электро" -OrganizationLocation "Офис Щербаковская"
                }
            }
        } | ConvertTo-Csv -NoTypeInformation
        
        Собрать информацию на удалённых ноутбуках в Московском офисе компании Миг Электро и сконвертировать информацию в CSV
    
    .NOTES
        Автор:            Максим Кулагин <m.kulagin@cits.ru>
        Последняя правка: 2020-09-15
        Версия 1.0 - Первый релиз
        Версия 1.1 - Добавлен параметр UseWMI. Понижена минимальная версия Powershell
        Версия 1.2 - Добавлено автоматическое определение типа компьютера
        Версия 1.3 - Убран параметр UseWMI, переделано всё на CIM-классы
        
        TODO:
        Есть проблема с определением железок от НР, если это не ноут. Например вместо модели HP 8184 должна быть HP 260 G2 DM
    #>
    
    param (
        [Parameter(ValueFromPipeline)]
        [string]$ComputerName = "localhost",
        [ValidateSet("Низкая","Средняя","Высокая")]
        [string]$Criticality = "Низкая",
        [string]$Description,
        [string]$OrganizationName = "<Название компании>",
        [string]$OrganizationLocation = "<Расположение компании>",
        [string]$StatusMark = "Выдано",
        [string]$OSFamily = "Windows"
    )
    process {

        (Test-WSMan -ComputerName $ComputerName -ErrorAction SilentlyContinue).ProductVersion -match "(\d).\d+$" | Out-Null
        $WSManVersion = $Matches[1]

        switch ($WSManVersion) {
            3 {
                $CIMOption = New-CimSessionOption -Protocol Wsman
            }
            default {
                $CIMOption = New-CimSessionOption -Protocol Dcom
            }
        }

        $CIMSession = New-CimSession -ComputerName $ComputerName -SessionOption $CIMOption
        $MB = Get-CimInstance -ClassName CIM_Card -CimSession $CIMSession
        $OS = Get-CimInstance -ClassName CIM_OperatingSystem -CimSession $CIMSession
        $CPU = Get-CimInstance -ClassName CIM_Processor -CimSession $CIMSession
        $Comp = Get-CimInstance -ClassName CIM_ComputerSystem -CimSession $CIMSession

        $RAMInfo = ""
        Get-CimInstance -ClassName CIM_PhysicalMemory -CimSession $CIMSession | ForEach-Object {
            $RAMInfo += [string]($_.Capacity / 1Gb) + " Gb (" + $_.PartNumber + "); "
        }
        $RAMInfo = $RAMInfo -replace "..$"

        switch ($Comp.PCSystemType) {
            1 {
                $ComputerType = "Настольный"
                $ComputerModel = $MB.Product
            }
            2 {
                $ComputerType = "Ноутбук"
                $ComputerModel = $Comp.Model
            }
        }

        [PSCustomObject]@{

            "Название*"                                         = $OS.CSName
            "Описание"                                          = $Description
            "Организация->Название"                             = $OrganizationName
            "Критичность"                                       = $Criticality
            "Статус"                                            = $StatusMark
            "Бренд->Название"                                   = $MB.Manufacturer -replace "Micro-Star International Co., Ltd.", "MSI" -replace "Hewlett-Packard", "HP"
            "Модель->Название"                                  = $ComputerModel
            "Семейство ОС->Название"                            = $OSFamily
            "Версия ОС->Название"                               = $OS.Caption -replace "Microsoft " -replace "Майкрософт " -replace "Профессиональная", "Pro" -replace " \(Registered Trademark\)" -replace " для рабочих станций"
            "Процессор"                                         = $CPU.Name
            "ОЗУ"                                               = $RAMInfo
            "Тип"                                               = $ComputerType
            "Организация->Полное название"                      = $OrganizationName
            "Расположение->Название"                            = $OrganizationLocation
            "Бренд->Полное название"                            = $MB.Manufacturer -replace "Micro-Star International Co., Ltd.", "MSI" -replace "Hewlett-Packard", "HP"
            "Модель->Полное название"                           = $ComputerModel
            "Семейство ОС->Полное название"                     = $OSFamily
            "Версия ОС->Полное название"                        = $OS.Caption -replace "Microsoft " -replace "Майкрософт " -replace "Профессиональная", "Pro" -replace " \(Registered Trademark\)" -replace " для рабочих станций"
        }
    Get-CimSession | Remove-CimSession
    }
}

function Export-MotherboardToItop {
    param (
        [Parameter(ValueFromPipeline)]
        [string]$ComputerName = "localhost"
    )
    process {
            
        (Test-WSMan -ComputerName $ComputerName -ErrorAction SilentlyContinue).ProductVersion -match "(\d).\d+$" | Out-Null
        $WSManVersion = $Matches[1]
    
        switch ($WSManVersion) {
            3 {
                $CIMOption = New-CimSessionOption -Protocol Wsman
            }
            default {
                $CIMOption = New-CimSessionOption -Protocol Dcom
            }
        }
    
        $CIMSession = New-CimSession -ComputerName $ComputerName -SessionOption $CIMOption
        $MB = Get-CimInstance -ClassName CIM_Card -CimSession $CIMSession
        $Comp = Get-CimInstance  -ClassName CIM_ComputerSystem -CimSession $CIMSession

        switch ($Comp.PCSystemType) {
            1 {
                $ComputerModel = $MB.Product
            }
            2 {
                $ComputerModel = $Comp.Model
            }
        }
    
        [PSCustomObject]@{
            "Название*"                                         = $ComputerModel
            "Бренд->Название"                                   = $MB.Manufacturer -replace "Micro-Star International Co., Ltd.", "MSI" -replace "Hewlett-Packard", "HP"
            "Тип устройства*"                                   = "Персональный компьютер"
            "Бренд->Полное название"                            = $MB.Manufacturer -replace "Micro-Star International Co., Ltd.", "MSI" -replace "Hewlett-Packard", "HP"
        }
    }
}