#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Item Classifications
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Item Classifications. Script add new or will update existing Classifications.
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

$csvItemClassificationsPath = -join ($csvImportCatalog, "ItemClassifications.csv")

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
    $itmClass = Import-Csv -Delimiter ';' -Path $csvItemClassificationsPath;
    $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
    if ($itmClass.Count -gt 1) {
        $total = $itmClass.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;
    foreach ($itmC in $itmClass) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OICL"" WHERE ""U_ClsCode"" = N'{0}'", $itmC.ClassificationCode));
	
            #Creating Item Property object
            $itmClassificaiton = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ItemClassification)
            #Checking that the property already exist
            if ($rs.RecordCount -gt 0) {
                $dummy = $itmClassificaiton.GetByKey($rs.Fields.Item(0).Value);
                $exists = $true
            }
            else {
                $itmClassificaiton.U_ClsCode = $itmC.ClassificationCode;
                $exists = $false
            }
   
            $itmClassificaiton.U_ClsName = $itmC.ClassificationName;
            $itmClassificaiton.U_GrpCode = $itmC.Group; 
   
            if ($itmC.ProductionOrders -eq 'Y') {
                $itmClassificaiton.U_ProdOrders = 'Y'
            }
            else {
                $itmClassificaiton.U_ProdOrders = 'N'
            }
	
            if ($itmC.ShipmentsDocumentation -eq 'Y') {
                $itmClassificaiton.U_ShipDoc = 'Y'
            }
            else {
                $itmClassificaiton.U_ShipDoc = 'N'
            }
	
            if ($itmC.PickLists -eq 'Y') {
                $itmClassificaiton.U_PickLists = 'Y'
            }
            else {
                $itmClassificaiton.U_PickLists = 'N'
            }
	
            if ($itmC.MSDS -eq 'Y') {
                $itmClassificaiton.U_MSDS = 'Y'
            }
            else {
                $itmClassificaiton.U_MSDS = 'N'
            }
	
            if ($itmC.PurchaseOrders -eq 'Y') {
                $itmClassificaiton.U_PurOrders = 'Y'
            }
            else {
                $itmClassificaiton.U_PurOrders = 'N'
            }
	
            if ($itmC.Returns -eq 'Y') {
                $itmClassificaiton.U_Returns = 'Y'
            }
            else {
                $itmClassificaiton.U_Returns = 'N'
            }
	
            if ($itmC.Other -eq 'Y') {
                $itmClassificaiton.U_Other = 'Y'
            }
            else {
                $itmClassificaiton.U_Other = 'N'
            }
	
            $itmClassificaiton.U_Remarks = $itmC.Remarks;
	
            $message = 0
            #Adding or updating Items Properties depends on exists in the database
            if ($exists -eq 1) {
                $message = $itmClassificaiton.Update()
            }
            else {
                $message = $itmClassificaiton.Add()
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
            $ms = [string]::Format("Error when {0} Item Classification Group with Code {1} Details: {2}", $taskMsg, $itmC.ClassificationCode, $err);
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
