function Export-PersonToItop {

    param (
        [Parameter(ValueFromPipeline)]
        [string]$FromCSV = ".\Сотрудники шаблон.csv",
        [string]$ToCSV = ".\Сотрудники $($OrganizationName).csv",
        [string]$OrganizationName = "<Название компании>",
        [string]$OrganizationLocation = "<Расположение компании>",
        [string]$Status = "Активный",
        [string]$ContactType = "Персона",
        [string]$Notify = "Да"
    )

    process {
            
        $Persons = Import-Csv $FromCSV -Delimiter ";"
        $PersonsInfo = @()
        foreach ($Person in $Persons) {
            
            $PersonsInfo += [PSCustomObject]@{
                "Имя"                    = $Person.Имя
                "Фамилия"                = $Person.Фамилия
                "Полное название"        = "$($Person.Имя) $($Person.Фамилия)"
                "Организация->Название"  = $OrganizationName
                "Статус"                 = $Status
                "Расположение->Название" = $OrganizationLocation
                "Email"                  = $Person.Почта
                "Телефон"                = $Person.'Добавочный номер'
                "Функция"                = $Person.Должность
                "Уведомлять"             = $Notify
                "Тип контакта"           = $ContactType
                "Мобильный телефон"      = $Person.'Мобильный телефон'
            }
        }
        $PersonsInfo | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" | Out-File $ToCSV -Encoding unicode
        Invoke-Item $ToCSV
    }
}

Export-PersonToItop -OrganizationName "Компания" -OrganizationLocation "Компания Город"
