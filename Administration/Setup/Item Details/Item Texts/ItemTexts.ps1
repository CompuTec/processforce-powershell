#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Texts
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Texts. Script add new or will update existing Texts.
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

$csvItemTextsPath = -join ($csvImportCatalog, "ItemTexts.csv")

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
    $itemTexts = Import-Csv -Delimiter ';' -Path $csvItemTextsPath;
    $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
    if ($itemTexts.Count -gt 1) {
        $total = $itemTexts.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;
    foreach ($itemText in $itemTexts) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OITX"" WHERE ""U_TxtCode"" = N'{0}'", $itemText.TextCode));
	
            #Creating Item Property object
            $text = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ItemText)
            #Checking that the property already exist
            if ($rs.RecordCount -gt 0) {
                $dumy = $text.GetByKey($rs.Fields.Item(0).Value);
                $exists = $true
            }
            else {
                $text.U_TxtCode = $itemText.TextCode;
                $exists = $false
            }
   
            $text.U_TxtName = $itemText.TextName;
            $text.U_GrpCode = $itemText.Group; 
	
            if ($itemText.ProductionOrders -eq 'Y') {
                $text.U_ProdOrders = 'Y'
            }
            else {
                $text.U_ProdOrders = 'N'
            }
	
            if ($itemText.ShipmentsDocumentation -eq 'Y') {
                $text.U_ShipDoc = 'Y'
            }
            else {
                $text.U_ShipDoc = 'N'
            }
	
            if ($itemText.PickLists -eq 'Y') {
                $text.U_PickLists = 'Y'
            }
            else {
                $text.U_PickLists = 'N'
            }
	
            if ($itemText.MSDS -eq 'Y') {
                $text.U_MSDS = 'Y'
            }
            else {
                $text.U_MSDS = 'N'
            }
	
            if ($itemText.PurchaseOrders -eq 'Y') {
                $text.U_PurOrders = 'Y'
            }
            else {
                $text.U_PurOrders = 'N'
            }
	
            if ($itemText.Returns -eq 'Y') {
                $text.U_Returns = 'Y'
            }
            else {
                $text.U_Returns = 'N'
            }
	
            if ($itemText.Other -eq 'Y') {
                $text.U_Other = 'Y'
            }
            else {
                $text.U_Other = 'N'
            }
	
            $text.U_Remarks = $itemText.Remarks;
	
            $message = 0
            #Adding or updating Items Texts depends on exists in the database
            if ($exists -eq $true) {
                $message = $text.Update()
            }
            else {
                $message = $text.Add()
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
            $ms = [string]::Format("Error when {0} Item Texts with Code {1} Details: {2}", $taskMsg, $itemText.TextCode, $err);
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