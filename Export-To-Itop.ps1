. .\Write-Log.ps1

#Requires -Version 2.0
function Export-KEToItop {
    <#
    .SYNOPSIS
        ��� ������� ������������ ��� ����� ���������� � ��, ������� ����� ����� ������������� � iTop ��� ���������������� ������� (��).
        ������������� ������������ ������ � ConvertTo-CSV, ����������� ����� CSV-������.
    
    .DESCRIPTION
        ���������� ���������� � ������� ����������. ����� ���: ����������� �����, ���������, ����������� ������, ������ ������������ �������.
        ������ ����� ������������ �� ���������� �����������.
    
    .PARAMETER ComputerType
        ��� ����������: ���������� ��� �������. �� ���������, ����������.
    
    .PARAMETER Criticality
        ����������� ����������: ������, ������� ��� �������. �� ���������, ������.
    
    .PARAMETER Description
        �������� ����������. ������� ��������� ��� ������ ��� ����������� ����� TeamViewer (��� �� ��������) � ���������� � �������������� ������� ������� ���������������.
    
    .PARAMETER OrganizationName
        �������� �����������, ������-�-������, ��� � ������. �������� "��� (��������������������)". �� ���������, <�������� ��������>.
    
    .PARAMETER OrganizationLocation
        ������������ �����������, ������-�-������, ��� � ������. �������� "���� ��� �����". �� ���������, <������������ ��������>.
    
    .PARAMETER UseWMI
        ������������ WMI ������ CIM. ���������� ��� ������ � Powershell 2.0. �� ���������, CIM
    
    .EXAMPLE
        PS C:\> Export-KEToItop
        ������� ���������� �� ��������� ����������
        
    .EXAMPLE
        PS C:\> Export-KEToItop -ComputerName MANAGER01
        ������� ���������� �� �������� ���������� MANAGER01
        
    .EXAMPLE
        PS C:\> Export-KEToItop -ComputerName MANAGER01 -UseWMI
        ������� ���������� �� �������� ���������� MANAGER01 ��������� WMI.
    
    .EXAMPLE
        PS C:\> 
        $computers = (Get-ADComputer -SearchBase "OU=Moscow,OU=Workstations,DC=mege,DC=ru" -Filter {Enabled -eq "True" -and Name -like "*NOTE*"}).Name
        $computers | foreach {
            if (Test-Connection $_ -Count 1 -Quiet) {
                Test-WSMan -ComputerName $_ | Out-Null
                if ($?) {
                    Export-KEToItop -ComputerName $_ -OrganizationName "��� �������" -OrganizationLocation "���� ������������"
                }
            }
        } | ConvertTo-Csv -NoTypeInformation
        
        ������� ���������� �� �������� ��������� � ���������� ����� �������� ��� ������� � ��������������� ���������� � CSV
    
    .NOTES
        �����:            ������ ������� <m.kulagin@cits.ru>
        ��������� ������: 2020-09-15
        ������ 1.0 - ������ �����
        ������ 1.1 - �������� �������� UseWMI. �������� ����������� ������ Powershell
        ������ 1.2 - ��������� �������������� ����������� ���� ����������
        ������ 1.3 - ����� �������� UseWMI, ���������� �� �� CIM-������
        
        TODO:
        ���� �������� � ������������ ������� �� ��, ���� ��� �� ����. �������� ������ ������ HP 8184 ������ ���� HP 260 G2 DM
    #>
    
    param (
        [Parameter(ValueFromPipeline)]
        [string]$ComputerName = "localhost",
        [ValidateSet("������","�������","�������")]
        [string]$Criticality = "������",
        [string]$Description,
        [string]$OrganizationName = "<�������� ��������>",
        [string]$OrganizationLocation = "<������������ ��������>",
        [string]$StatusMark = "������",
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
                $ComputerType = "����������"
                $ComputerModel = $MB.Product
            }
            2 {
                $ComputerType = "�������"
                $ComputerModel = $Comp.Model
            }
        }

        [PSCustomObject]@{

            "��������*"                                         = $OS.CSName
            "��������"                                          = $Description
            "�����������->��������"                             = $OrganizationName
            "�����������"                                       = $Criticality
            "������"                                            = $StatusMark
            "�����->��������"                                   = $MB.Manufacturer -replace "Micro-Star International Co., Ltd.", "MSI" -replace "Hewlett-Packard", "HP"
            "������->��������"                                  = $ComputerModel
            "��������� ��->��������"                            = $OSFamily
            "������ ��->��������"                               = $OS.Caption -replace "Microsoft " -replace "���������� " -replace "����������������", "Pro" -replace " \(Registered Trademark\)" -replace " ��� ������� �������"
            "���������"                                         = $CPU.Name
            "���"                                               = $RAMInfo
            "���"                                               = $ComputerType
            "�����������->������ ��������"                      = $OrganizationName
            "������������->��������"                            = $OrganizationLocation
            "�����->������ ��������"                            = $MB.Manufacturer -replace "Micro-Star International Co., Ltd.", "MSI" -replace "Hewlett-Packard", "HP"
            "������->������ ��������"                           = $ComputerModel
            "��������� ��->������ ��������"                     = $OSFamily
            "������ ��->������ ��������"                        = $OS.Caption -replace "Microsoft " -replace "���������� " -replace "����������������", "Pro" -replace " \(Registered Trademark\)" -replace " ��� ������� �������"
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
            "��������*"                                         = $ComputerModel
            "�����->��������"                                   = $MB.Manufacturer -replace "Micro-Star International Co., Ltd.", "MSI" -replace "Hewlett-Packard", "HP"
            "��� ����������*"                                   = "������������ ���������"
            "�����->������ ��������"                            = $MB.Manufacturer -replace "Micro-Star International Co., Ltd.", "MSI" -replace "Hewlett-Packard", "HP"
        }
    }
}