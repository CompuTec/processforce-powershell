#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Protocols Templates
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# THIS VERSION CAN BE USED ONLY FOR VERSIONS LOWER THAN ProcessForce 9.3 PL: 08 
# Description:
#      Import Protocols Templates. Script add new or will update existing Templates.
#      You need to have all requred files for import.
# Warning:
#   It's recommended run script when all users all disconnected.
#   Before running this script please do database backup.
# Troubleshooting:
#   https://connect.computec.pl/display/PF930EN/FAQ+PowerShell
#   https://connect.computec.pl/display/PF920EN/FAQ+PowerShell
# Script source:
#   https://connect.computec.pl/display/PF930EN/Quality+Control+scripts
########################################################################
Write-Host -backgroundcolor Yellow -foregroundcolor DarkBlue ("Script Version:" + $SCRIPT_VERSION)
#endregion

#region #PF API library usage
# You need to check in what architecture PowerShell ISE is running (x64 or x86),
# you need run PowerShell ISE in the same architecture like PF API is installed (check in Windows -> Programs & Features)
# Examples:
#     SAP Client x64 + PF x64 installed on DB/Company => PF API x64 => Windows PowerShell ISE
#     SAP Client x86 + PF x86 installed on DB/Company => PF API x86 => Windows PowerShell ISE x86

[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#endregion

#region Script parameters

$csvImportCatalog = $PSScriptRoot + "\"
#If you are using lower version of PowerShell than PowerShell 4.0 you can use static path
# $csvImportCatalog = "C:\PS\PF\";

$csvQualityTemplatesForTestProtocolsPath = -join ($csvImportCatalog, "QualityTemplatesForTestProtocols.csv")
$csvQualityTemplatesForTestProtocolsPropertiesPath = -join ($csvImportCatalog, "QualityTemplatesForTestProtocolsProperties.csv")

#endregion

#region #Datbase/Company connection settings
#configuration xml
$configurationXMLFilePath = -join ($csvImportCatalog, "configuration.xml");
if (!(Test-Path $configurationXMLFilePath -PathType Leaf)) {
    Write-Host -BackgroundColor Red ([string]::Format("File: {0} don't exists.", $configurationXMLFilePath));
    return;
}
[xml] $configurationXml = Get-Content -Encoding UTF8 $configurationXMLFilePath
$xmlConnection = $configurationXml.SelectSingleNode("/configuration/connection");

$connectionConfirmation = [string]::Format('You are connecting to Database: {0} on Server: {1} as User: {2}. Do you want to continue [y/n]?:', $xmlConnection.Database, $xmlConnection.SQLServer, $xmlConnection.UserName);
Write-Host $connectionConfirmation -backgroundcolor Yellow -foregroundcolor DarkBlue -NoNewline
$confirmation = Read-Host
if (($confirmation -ne 'y') -and ($confirmation -ne 'Y')) {
    return;
}

$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.LicenseServer = $xmlConnection.LicenseServer;
$pfcCompany.SQLServer = $xmlConnection.SQLServer;
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::[string]$xmlConnection.DbServerType;
$pfcCompany.Databasename = $xmlConnection.Database;
$pfcCompany.UserName = $xmlConnection.UserName;
$pfcCompany.Password = $xmlConnection.Password;
 
#endregion

#region #Connect to company
 
write-host -backgroundcolor yellow -foregroundcolor black  "Trying connect..."
$version = [CompuTec.Core.CoreConfiguration+DatabaseSetup]::AddonVersion
write-host -backgroundcolor green -foregroundcolor black "PF API Library:" $version';' 'Host:'(Get-WmiObject Win32_OperatingSystem).CSName';' 'OSArchitecture:' (Get-WmiObject Win32_OperatingSystem).OSArchitecture
 
try {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'code')]
    $code = $pfcCompany.Connect()
 
    write-host -backgroundcolor green -foregroundcolor black "Connected to:" $pfcCompany.SapCompany.CompanyName "/ " $pfcCompany.SapCompany.CompanyDB"" "Sap Company version: " $pfcCompany.SapCompany.Version
}
catch {
    #Show error messages & stop the script
    write-host "Connection Failure: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
 
    write-host "LicenseServer:" $pfcCompany.LicenseServer
    write-host "SQLServer:" $pfcCompany.SQLServer
    write-host "DbServerType:" $pfcCompany.DbServerType
    write-host "Databasename" $pfcCompany.Databasename
    write-host "UserName:" $pfcCompany.UserName
}

#If company is not connected - stops the script
if (-not $pfcCompany.IsConnected) {
    write-host -backgroundcolor yellow -foregroundcolor black "Company is not connected"
    return 
}

try {

    [array] $csvProtocolsList = Import-Csv -Delimiter ';' $csvQualityTemplatesForTestProtocolsPath;
	
    if ((Test-Path -Path $csvQualityTemplatesForTestProtocolsPropertiesPath -PathType leaf) -eq $true) {
        [array] $csvProtocolsProp = Import-Csv -Delimiter ';' $csvQualityTemplatesForTestProtocolsPropertiesPath;
    }
    else {
        write-host "Qualiti Templates For Test protocls Properties - csv not available."
    }

    write-Host 'Preparing data: '
    $totalRows = $csvProtocolsList.Count + $csvProtocolsProp.Count;
    $protocolsList = New-Object 'System.Collections.Generic.List[array]'
    $protocolsPropDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
	
    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $csvProtocolsList) {
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        $protocolsList.Add([array]$row);
    }

    foreach ($row in $csvProtocolsProp) {
        $key = $row.TemplateCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
	
        if ($protocolsPropDict.ContainsKey($key) -eq $false) {
            $protocolsPropDict[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $protocolsPropDict[$key];
		
        $list.Add([array]$row);
    }

    Write-Host '';
    Write-Host 'Adding/updating data: ' -NoNewLine;

    if ($protocolsList.Count -gt 1) {
        $total = $protocolsList.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;

    #Checking that Template already exist 
    foreach ($csvTemplate in $protocolsList) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
     
            $rs.DoQuery([string]::Format( "SELECT ""U_TemplateCode"", ""Code"" FROM ""@CT_PF_OTPT"" WHERE ""U_TemplateCode"" = N'{0}'", $csvTemplate.TemplateCode))
            $exists = $false;
            if ($rs.RecordCount -gt 0) {
                $exists = $true
            }
  
            #Creating Template
            $tmpl = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"TestProtocolTemplate")
            $rs.MoveFirst();
    
            if ($exists -eq $true) {
                $dummy = $tmpl.getByKey($rs.Fields.Item('Code').Value);
            }
            else {
                $tmpl.U_TemplateCode = $csvTemplate.TemplateCode;
                $tmpl.U_TemplateName = $csvTemplate.TemplateName;
                if ($csvTemplate.ValidFrom -ne "") {
                    $tmpl.U_ValidFromDate = $csvTemplate.ValidFrom;
                }
                else {
                    $tmpl.U_ValidFromDate = [DateTime]::MinValue 
                }
                if ($csvTemplate.ValidTo -ne "") {
                    $tmpl.U_ValidToDate = $csvTemplate.ValidTo;
                }
                else {
                    $tmpl.U_ValidToDate = [DateTime]::MinValue 
                }
                if ($csvTemplate.GroupCode -ne "") {
                    $tmpl.U_GrpCode = $csvTemplate.GroupCode;
                }
                if ($csvTemplate.Remarks -ne "") {
                    $tmpl.U_Remarks = $csvTemplate.Remarks;
                }
            }
            #Data loading from the csv file - Rows for templates from Quality_TemplatesForTestProtocolsProperties.csv file
            [array]$Properties = $protocolsPropDict[$csvTemplate.TemplateCode];
            if ($Properties.count -gt 0) {
                #Deleting all exisitng Phrases
                $count = $tmpl.Properties.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $tmpl.Properties.DelRowAtPos(0);
                }
                $tmpl.Properties.SetCurrentLine($tmpl.Properties.Count - 1);
         
                #Adding Properties
                foreach ($prop in $Properties) {
                    try {
                        $tmpl.Properties.U_PrpCode = $prop.PropertyCode;
                    } catch {
                        $err = [string]::Format("Property with Code:{0} don't exists. Details: {1}",$prop.PropertyCode,[string]$_.Exception.Message);
                        Throw [System.Exception] ($err)
                    }
                    $tmpl.Properties.U_Expression = $prop.Expression;
			
                    if ($prop.RangeFrom -ne "") {
                        $tmpl.Properties.U_RangeValueFrom = $prop.RangeFrom;
                    }
                    else {
                        $tmpl.Properties.U_RangeValueFrom = 0;
                    }
                    $tmpl.Properties.U_RangeValueTo = $prop.RangeTo;
			
                    if ($prop.UoM -ne "") {
                        $tmpl.Properties.U_UnitOfMeasure = $prop.UoM;
                    }
			
                    if ($prop.ReferenceCode -ne "") {
                        $tmpl.Properties.U_RefCode = $prop.ReferenceCode;
                    }
			
                    if ($prop.ValidFrom -ne "") {
                        $tmpl.Properties.U_ValidFromDate = $prop.ValidFrom;
                    }
                    else {
                        $tmpl.Properties.U_ValidFromDate = [DateTime]::MinValue 
                    }
                    if ($prop.ValidTo -ne "") {
                        $tmpl.Properties.U_ValidToDate = $prop.ValidTo;
                    }
                    else {
                        $tmpl.Properties.U_ValidToDate = [DateTime]::MinValue 
                    }
			
                    $dummy = $tmpl.Properties.Add()
                }
            }
  
            $message = 0
     
            #Adding or updating Template depends on exists in the database
            if ($exists -eq $true) {
                $message = $tmpl.Update()
            }
            else {
                $message = $tmpl.Add()
            }
     
            if ($message -lt 0) {    
                $err = $pfcCompany.GetLastErrorDescription()
                Throw [System.Exception] ($err)
            }
        }
        Catch {
            $err = $_.Exception.Message;
            if ($exists -eq $true) {
                $taskMsg = "adding";
            }
            else {
                $taskMsg = "updating"
            }
            $ms = [string]::Format("Error when {0} Protocol Template with Code {1} Details: {2}", $taskMsg, $csvTemplate.TemplateCode, $err);
            Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
            if ($pfcCompany.InTransaction) {
                $pfcCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
            } 
        }		 
    }
}
Catch {
    $err = $_.Exception.Message;
    $ms = [string]::Format("Exception occured: {0}", $err);
    Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
    if ($pfcCompany.InTransaction) {
        $pfcCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
    } 
}
Finally {
    #region Close connection
    if ($pfcCompany.IsConnected) {
        $pfcCompany.Disconnect()
        Write-Host '';
        write-host  –backgroundcolor green –foregroundcolor black "Disconnected from the company"
    }
    #endregion
}