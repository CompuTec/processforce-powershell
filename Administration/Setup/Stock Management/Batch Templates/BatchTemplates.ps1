#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Batch Templates
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.170) PL: 07 R1 (32-bit)
# Description:
#      Import Batch Templates. Script add new or will update existing Batch Templates.
#      You need to have all requred files for import.
# Warning:
#   It's recommended run script when all users all disconnected.
#   Before running this script please do database backup.
# Troubleshooting:
#   https://connect.computec.pl/display/PF930EN/FAQ+PowerShell
#   https://connect.computec.pl/display/PF920EN/FAQ+PowerShell
# Script source:
#   https://connect.computec.pl/display/PF930EN/
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

$csvBatchTemplatesPath = -join ($csvImportCatalog, "BatchTemplates.csv")

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
    [array] $csvBatchTemplates = Import-Csv -Delimiter ';' $csvBatchTemplatesPath;
	
    write-Host 'Preparing data: '
    $totalRows = $csvBatchTemplates.Count;
    $batchTemplatesList = New-Object 'System.Collections.Generic.List[array]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
	
    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $csvBatchTemplates) {
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        $batchTemplatesList.Add([array]$row);
    }

    Write-Host '';
    Write-Host 'Adding/updating data: ' -NoNewline;

    $rs = $pfCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
    if ($batchTemplatesList.Count -gt 1) {
        $total = $batchTemplatesList.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;
    foreach ($template in $batchTemplatesList) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $rs.DoQuery([string]::Format("SELECT ""Code"" AS ""Code"" FROM ""@CT_PF_OBTM"" WHERE ""Code"" = N'{0}'", $template.TemplateCode));
	
            #Creating Item Property object
            $templateObj = $pfCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::BatchTemplate);
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
            if (![string]::IsNullOrWhiteSpace($template.DateFormat)) {
                $templateObj.U_DateFormat = $template.DateFormat;
            }
            if (![string]::IsNullOrWhiteSpace($template.Counter)) {
                $templateObj.U_Counter = $template.Counter;
            }
            if (![string]::IsNullOrWhiteSpace($template.SuppDateFormat)) {
                $templateObj.U_SuppDateFormat = $template.SuppDateFormat;
            }
            if (![string]::IsNullOrWhiteSpace($template.SuppCounter)) {
                $templateObj.U_SuppCounter = $template.SuppCounter;
            }
            if (![string]::IsNullOrWhiteSpace($template.Formula)) {
                $templateObj.U_Template = $template.Formula;
            }
            if (![string]::IsNullOrWhiteSpace($template.SuppFormula)) {
                $templateObj.U_SuppTemplate = $template.SuppFormula;
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
            $ms = [string]::Format("Error when {0} Batch Template with Code {1} Details: {2}", $taskMsg, [string]$template.TemplateCode, [string]$err);
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
