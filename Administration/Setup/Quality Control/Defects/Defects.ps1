﻿#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Defects
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Defects. Script add new or will update existing Defects.
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

$csvDefectsPath = -join ($csvImportCatalog, "Defects.csv")

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

    #Data loading from a csv file
    $mds = Import-Csv -Delimiter ';' -Path $csvDefectsPath;
    $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
    if ($mds.Count -gt 1) {
        $total = $mds.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;

    foreach ($md in $mds) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_ODEF"" WHERE ""U_DefCode"" = N'{0}'", $md.DefectCode));
	
            #Creating object
            $mdObj = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Defect)
            #Checking if object already exist
            if ($rs.RecordCount -gt 0) {
                $dummy = $mdObj.GetByKey($rs.Fields.Item(0).Value);
                $exists = $true
            }
            else {
                $mdObj.U_DefCode = $md.DefectCode;
                $exists = $false
            }
   
            $mdObj.U_DefName = $md.DefectName;
            $mdObj.U_DefGrpCode = $md.Group; 
            $mdObj.U_DefRemarks = $md.Remarks;
	
            $message = 0
            #Adding or updating depends if it already exists in the database
            if ($exists -eq $true) {
                $message = $mdObj.Update()
            }
            else {
                $message = $mdObj.Add()
            }
            
            if ($message -lt 0) {    
                $err = $pfcCompany.GetLastErrorDescription()
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
            $ms = [string]::Format("Error when {0} Defect with Code {1} Details: {2}", $taskMsg, $md.DefectCode, $err);
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