﻿#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Batch Master Data
########################################################################
$SCRIPT_VERSION = "3.1"
# Last tested PF version: ProcessForce 9.3 (9.30.210) (64-bit)
# Description:
#      Import Batch Master Data. Script will update existing Batch Master Data.
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

$csvBatchDetailsPath = -join ($csvImportCatalog, "BatchMasterData.csv")

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
	#Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
	$csvItems = Import-Csv -Delimiter ';' -Path $csvBatchDetailsPath;
	if ($csvItems.Count -gt 1) {
		$total = $csvItems.Count
	}
	else {
		$total = 1
	}
	$progressItterator = 0;
	$progress = 0;
	$beforeProgress = 0; 
	#Checking that Item Details already exist
	foreach ($csvItem in $csvItems) {
		try {
			$progressItterator++;
			$progress = [math]::Round(($progressItterator * 100) / $total);
			if ($progress -gt $beforeProgress) {
				Write-Host $progress"% " -NoNewline
				$beforeProgress = $progress
			}
			$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
			$rs.DoQuery([string]::Format( "SELECT ""Code"" FROM ""@CT_PF_OABT""
    WHERE ""U_DistNumber"" = N'{0}' AND ""U_ItemCode"" =  N'{1}'", $csvItem.BatchCode, $csvItem.ItemCode))
			if ($rs.RecordCount -eq 0) {
				$err = [string]::Format("Batch Master Data for ItemCode: {0} and Batch: {1} don't exists. Run Restore Batch Master Data or check correctnes of csv fiels", $csvItem.ItemCode, $csvItem.BatchCode);
				Throw [System.Exception]($err);
			}
			#Creating Additional Batch Master Data
			$abd = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::AdditionalBatchDetails)
			$dummy = $abd.GetByKey($rs.Fields.Item(0).Value);
		
			if ($csvItem.BatchAttribute1 -gt '') {
				$abd.U_MnfSerial = $csvItem.BatchAttribute1;
			}
		
			if ($csvItem.BatchAttribute2 -gt '') {
				$abd.U_LotNumber = $csvItem.BatchAttribute2;
			}
		
			if ($csvItem.SupplierBatch -gt '') {
				$abd.U_SupNumber = $csvItem.SupplierBatch;
			}
		
			if ($csvItem.Status -eq 'R') {
				$abd.U_Status = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Released
			}
		
			if ($csvItem.Status -eq 'A') {
				$abd.U_Status = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::NotAccesible;
			}
		
			if ($csvItem.Status -eq 'L') {
				$abd.U_Status = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Locked
			}
		
			if ($csvItem.AddminsionDate -gt '') {
				$abd.U_AdmDate = $csvItem.AddminsionDate;
			}
		
			if ($csvItem.VendorManufacturingDate -gt '') {
				$abd.U_VndDate = $csvItem.VendorManufacturingDate;
			}
		
			if ($csvItem.ExpiryDate -gt '') {
				$abd.U_ExpiryDate = $csvItem.ExpiryDate;
			}
		
			if ($csvItem.ExpiryTime -gt '') {
				$abd.U_ExpiryTime = $csvItem.ExpiryTime;
			}
		
			if ($csvItem.ConsumeByDate -gt '') {
				$abd.U_ConsDate = $csvItem.ConsumeByDate;
			}
		
			if ($csvItem.LastInspectionDate -gt '') {
				$abd.U_LstInDate = $csvItem.LastInspectionDate;
			}
		
			if ($csvItem.InspectionDate -gt '') {
				$abd.U_InDate = $csvItem.InspectionDate;
			}
		
			if ($csvItem.NextInspectionDate -gt '') {
				$abd.U_NxtInDate = $csvItem.NextInspectionDate;
			}
		
			if ($csvItem.WarningDatePriorExpiry -gt '') {
				$abd.U_WExDate = $csvItem.WarningDatePriorExpiry;
			}
		
			if ($csvItem.WarningDatePriorConsume -gt '') {
				$abd.U_WCoDate = $csvItem.WarningDatePriorConsume;
			}
		
			if ($csvItem.Revision -gt '') {
				$abd.U_Revision = $csvItem.Revision;
			}

			if ($csvItem.RevisionDesc -gt '') {
				$abd.U_RevisionDesc = $csvItem.RevisionDesc;
			}
			
			if ($csvItem.Remarks -gt '') {
				$abd.U_Remarks = $csvItem.Remarks;
			}
		
			$message = 0

			$message = $abd.Update()
       
			if ($message -lt 0) {    
				$err = $pfcCompany.GetLastErrorDescription()
				Throw [System.Exception] ($err)
			}
		}
		Catch {
			$err = $_.Exception.Message;
			$taskMsg = "updating"

			$ms = [string]::Format("Error when {0} Batch Master Data with ItemCode: {1} and Batch: {2} Details: {3}", $taskMsg, $csvItem.ItemCode, $csvItem.BatchCode, $err);
			Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
			if ($pfcCompany.InTransaction) {
				$pfcCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
			} 
		}		

    
	}
}
Catch {
	$err = $_.Exception.Message;
	$ms = [string]::Format("Exception occured:{0}", $err);
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
