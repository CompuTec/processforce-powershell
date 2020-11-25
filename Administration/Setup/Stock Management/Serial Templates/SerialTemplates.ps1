#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Serial Templates
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.170) PL: 07 R1 (32-bit)
# Description:
#      Import Serial Templates. Script add new or will update existing Serial Templates.
#      You need to have all requred files for import.
# Warning:
#   It's recommended run script when all users all disconnected.
#   Before running this script please do database backup.
# Troubleshooting:
#   https://connect.computec.pl/display/PF100EN/PowerShell+FAQ
#   https://connect.computec.pl/display/PF930EN/PowerShell+FAQ
# Script source:
#   https://code.computec.pl/repos?visibility=public
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

$csvSerialTemplatesPath = -join ($csvImportCatalog, "SerialTemplates.csv")

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

$pfCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfCompany.LicenseServer = $xmlConnection.LicenseServer;
$pfCompany.SQLServer = $xmlConnection.SQLServer;
$pfCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::[string]$xmlConnection.DbServerType;
$pfCompany.Databasename = $xmlConnection.Database;
$pfCompany.UserName = $xmlConnection.UserName;
$pfCompany.Password = $xmlConnection.Password;
 
#endregion

#region #Connect to company
 
write-host -backgroundcolor yellow -foregroundcolor black  "Trying connect..."
$version = [CompuTec.Core.CoreConfiguration+DatabaseSetup]::AddonVersion
write-host -backgroundcolor green -foregroundcolor black "PF API Library:" $version';' 'Host:'(Get-WmiObject Win32_OperatingSystem).CSName';' 'OSArchitecture:' (Get-WmiObject Win32_OperatingSystem).OSArchitecture
 
try {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'code')]
    $code = $pfCompany.Connect()
 
    write-host -backgroundcolor green -foregroundcolor black "Connected to:" $pfCompany.SapCompany.CompanyName "/ " $pfCompany.SapCompany.CompanyDB"" "Sap Company version: " $pfCompany.SapCompany.Version
}
catch {
    #Show error messages & stop the script
    write-host "Connection Failure: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
 
    write-host "LicenseServer:" $pfCompany.LicenseServer
    write-host "SQLServer:" $pfCompany.SQLServer
    write-host "DbServerType:" $pfCompany.DbServerType
    write-host "Databasename" $pfCompany.Databasename
    write-host "UserName:" $pfCompany.UserName
}

#If company is not connected - stops the script
if (-not $pfCompany.IsConnected) {
    write-host -backgroundcolor yellow -foregroundcolor black "Company is not connected"
    return 
}

try {
    #Data loading from a csv file
    [array] $csvSerialTemplates = Import-Csv -Delimiter ';' $csvSerialTemplatesPath;
	
    write-Host 'Preparing data: ' -NoNewline
    $totalRows = $csvSerialTemplates.Count;
    $serialTemplatesList = New-Object 'System.Collections.Generic.List[array]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
	
    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $csvSerialTemplates) {
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        $serialTemplatesList.Add([array]$row);
    }

    Write-Host '';
    Write-Host 'Adding/updating data: ' -NoNewline;

    $rs = $pfCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
    if ($serialTemplatesList.Count -gt 1) {
        $total = $serialTemplatesList.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;
    foreach ($template in $serialTemplatesList) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $rs.DoQuery([string]::Format("SELECT ""Code"" AS ""Code"" FROM ""@CT_PF_OSTM"" WHERE ""Code"" = N'{0}'", $template.TemplateCode));
	
            #Creating Item Property object
            $templateObj = $pfCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::SerialTemplate);
            #Checking that the property already exist
            if ($rs.RecordCount -gt 0) {
                $key = $rs.Fields.Item("Code").Value;
                $dummy = $templateObj.GetByKey($key);
                $exists = $true
            }
            else {
                $templateObj.Code = $template.TemplateCode;
                $exists = $false
            }
            
            $templateObj.U_TemplateName = $template.TemplateName;
            if (![string]::IsNullOrWhiteSpace($template.DistDateFormat)) {
                $templateObj.U_DistDtFr = $template.DistDateFormat;
            }
            if (![string]::IsNullOrWhiteSpace($template.DistCounter)) {
                $templateObj.U_DistCntr = $template.DistCounter;
            }
            if (![string]::IsNullOrWhiteSpace($template.MnfDateFormat)) {
                $templateObj.U_MnfDtFr = $template.MnfDateFormat;
            }
            if (![string]::IsNullOrWhiteSpace($template.MnfCounter)) {
                $templateObj.U_MnfCntr = $template.MnfCounter;
            }
            if (![string]::IsNullOrWhiteSpace($template.LotDateFormat)) {
                $templateObj.U_LotDtFrmt = $template.LotDateFormat;
            }
            if (![string]::IsNullOrWhiteSpace($template.LotCounter)) {
                $templateObj.U_LotCntr = $template.LotCounter;
            }
            if (![string]::IsNullOrWhiteSpace($template.DistFormula)) {
                $templateObj.U_DistTempl = $template.DistFormula;
            }
            if (![string]::IsNullOrWhiteSpace($template.MnfFormula)) {
                $templateObj.U_MnfTempl = $template.MnfFormula;
            }
            if (![string]::IsNullOrWhiteSpace($template.LotFormula)) {
                $templateObj.U_LotTempl = $template.LotFormula;
            }


            $result = 0
            #Adding or updating Items Properties depends on exists in the database
            if ($exists -eq $true) {
                $result = $templateObj.Update()
            }
            else {
                $result = $templateObj.Add()
            }
            
            if ($result -lt 0) {    
                $err = $pfCompany.GetLastErrorDescription()
                Throw [System.Exception] ($err)
            }
        }
        Catch {
            $err = $_.Exception.Message;
            if ($exists -eq $false) {
                $taskMsg = "adding";
            }
            else {
                $taskMsg = "updating"
            }
            $ms = [string]::Format("Error when {0} Serial Template with Code {1} Details: {2}", $taskMsg, [string]$template.TemplateCode, [string]$err);
            Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
            if ($pfCompany.InTransaction) {
                $pfCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
            } 
        }
    }
}
Catch {
    $err = $_.Exception.Message;
    $ms = [string]::Format("Exception occured: {0}", $err);
    Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
    if ($pfCompany.InTransaction) {
        $pfCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
    } 
}
Finally {
    #region Close connection
    if ($pfCompany.IsConnected) {
        $pfCompany.Disconnect()
        Write-Host '';
        write-host  –backgroundcolor green –foregroundcolor black "Disconnected from the company"
    }
    #endregion
}
