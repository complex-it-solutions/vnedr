function Export-PersonsToItop {

    param (
        [Parameter(ValueFromPipeline)]
        [string]$FromCSV = ".\Сотрудники шаблон.csv",
        [string]$ToCSV = ".\Persons $($OrganizationName).csv",
        [string]$OrganizationName = "<Название компании>",
        [string]$OrganizationLocation = "<Расположение компании>",
        [string]$Status = "Активный",
        [string]$Notify = "да"
    )

    process {
            
        $Persons = Import-Csv $FromCSV -Delimiter ";"
        $PersonsInfo = @()
        foreach ($Person in $Persons) {
            
            $PersonsInfo += [PSCustomObject]@{
                "Имя"                    = $Person.Имя
                "Фамилия"                = $Person.Фамилия
                "Организация->Название"  = $OrganizationName
                "Статус"                 = $Status
                "Расположение->Название" = $OrganizationLocation
                "Email"                  = $Person.Почта
                "Телефон"                = $Person.'Добавочный номер'
                "Функция"                = $Person.Должность
                "Уведомлять"             = $Notify
                "Мобильный телефон"      = $Person.'Мобильный телефон'
            }
        }
        $PersonsInfo | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" | Out-File $ToCSV -Encoding default
    }
}

Export-PersonsToItop -OrganizationName "Компания" -OrganizationLocation "Расположение"
