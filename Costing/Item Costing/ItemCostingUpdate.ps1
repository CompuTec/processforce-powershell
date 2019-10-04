
#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Item Costing
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.210) (64-bit)
# Description:
#      Import Item Costing. Script will update only existing Item Costing Data. Remember to run Restore Item Costing Details before running this script.
#      If csv file for given section is not presented or don't have related records then on ItemDetails this section will be ommited. 
# Warning:
#   It's recommended run script when all users all disconnected.
#   Before running this script please do database backup.
#   Before running this script please restore Item Costing Details. #
#   This script allows only to update Item Costing on categories different than 000 #
# Troubleshooting:
#   https://connect.computec.pl/display/PF930EN/FAQ+PowerShell
#   https://connect.computec.pl/display/PF920EN/FAQ+PowerShell
# Script source:
#   https://connect.computec.pl/display/PF930EN/Item+details+script
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
# $csvImportCatalog = "C:\PS\PF\Costing\ItemCosting";

$csvItemCostingsPath = -join ($csvImportCatalog, "ItemCosting.csv");
$csvItemCostingDetailsPath = -join ($csvImportCatalog, "ItemCostingDetails.csv");

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
#endregion


try {

	#Data loading from a csv file
	Write-Host 'Preparing data: '
	#region import csv files
	[array]$csvItemCostings = Import-Csv -Delimiter ';' -Path $csvItemCostingsPath;
	[array]$csvItemCostingDetails = Import-Csv -Delimiter ';' -Path $csvItemCostingDetailsPath;

	
	
	$totalRows = $csvItemCostings.Count + $csvItemCostingDetails.Count 
	
	$dictionaryItemCosting = New-Object 'System.Collections.Generic.Dictionary[string,psobject]';

	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
 else {
		$total = 1
	}

	
	foreach ($row in $csvItemCostings) {
		$key = $row.ItemCode + '__' + $row.Revision + '__' + $row.CostCategory;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if (-Not $dictionaryItemCosting.ContainsKey($key)) {
			$dictionaryItemCosting.Add($key, [psobject]@{
					ItemCode     = $row.ItemCode
					Revision     = $row.Revision
					CostCategory = $row.CostCategory
					Details      = New-Object 'System.Collections.Generic.List[array]'
				});
		}
	}
	
	foreach ($row in $csvItemCostingDetails) {
		$key = $row.ItemCode + '__' + $row.Revision + '__' + $row.Category;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryItemCosting.ContainsKey($key)) {
			$list = $dictionaryItemCosting[$key].Details;
			$list.Add([array]$row);
		}
	}
	#endregion

	Write-Host '';
	Write-Host 'Adding/updating data: ' -NoNewline;

	$totalRows = $dictionaryItemCosting.Count
	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
 else {
		$total = 1
	}
	
	foreach ($key in $dictionaryItemCosting.Keys) {
		try {
			$csvItemCosting = $dictionaryItemCosting[$key];
			$progressItterator++;
			$progres = [math]::Round(($progressItterator * 100) / $total);
			if ($progres -gt $beforeProgress) {
				Write-Host $progres"% " -NoNewline
				$beforeProgress = $progres
			}
			#Creating Item Costing Object
			$ic = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ItemCosting);
			if ($csvItemCosting.CostCategory -eq '000') {
				throw [System.Exception] ("Masive update for Cost Category 000 is turned off - please make updates on custom Cost Category and use Roll-Over functionality");
			}
			#Checking if ItemCosting exists
			$retValue = $ic.Get($csvItemCosting.ItemCode, $csvItemCosting.Revision, $csvItemCosting.CostCategory)
	
			if (-Not $retValue) {
				throw [System.Exception] ("Item Costing Details don't exists");
			}
   
			#Data loading from the csv file - Costing Details for positions from ItemCosting.csv file
			#[array]$csvCostingDetails = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Costing\ItemCostingDetails.csv" | Where-Object {$_.ItemCode -eq $csvItemCosting.ItemCode -and $_.Revision -eq $csvItemCosting.Revision -and $_.Category -eq $csvItemCosting.CostCategory}
			$csvCostingDetails = $csvItemCosting.Details
			if($csvCostingDetails.Count -eq 0){
				throw [System.Exception]("Item Costing Details are missing. Check your csv file.");
			}
			foreach ($csvCD in $csvCostingDetails) {
			
				$count = $ic.CostingDetails.Count;
				for ($i = 0; $i -lt $count ; $i++) {
					$ic.CostingDetails.SetCurrentLine($i);
					if ($ic.CostingDetails.U_WhsCode -eq $csvCD.WhsCode) {
						#ML - Manual, MN - Manual no Roll-up, PL - Price List, PN - Price List no Roll-up, AC - Automatic, AN - Automatic no Roll-up
						$ic.CostingDetails.U_Type = $csvCD.Type
						$ic.CostingDetails.U_PriceList = $csvCD.PriceListCode
						$ic.CostingDetails.U_WhenZero = $csvCD.WhenZero
						$ic.CostingDetails.U_ItemCost = $csvCD.ItemCost
						$ic.CostingDetails.U_FixOH = $csvCD.FixedOH
						$ic.CostingDetails.U_FixOHPrct = $csvCD.FixedOHPrct
						$ic.CostingDetails.U_FixOHOther = $csvCD.FixedOHOther
						$ic.CostingDetails.U_VarOH = $csvCD.VariableOH
						$ic.CostingDetails.U_VarOHPrct = $csvCD.VariableOHPrct
						$ic.CostingDetails.U_VarOHOther = $csvCD.VariableOHOther
						$ic.CostingDetails.U_Remarks = $csvCD.Remarks
						break;
					}
				}
			}
			$ic.RecalculateCostingDetails()
			$ic.RecalculateRolledCosts()

			$message = $ic.Update()
			if ($message -lt 0) {  
				$err = $pfcCompany.GetLastErrorDescription()
				Throw [System.Exception] ($err);
			}
			
		}
		Catch {
			$err = $_.Exception.Message;
			$ms = [string]::Format("Error when updating Item Costing for ItemCode {0}, Revision {1}, Cost Category {2}, Details: {3}", $csvItemCosting.ItemCode, $csvItemCosting.Revision, $csvItemCosting.CostCategory, $err);
			Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
		}
	}
}
Catch {
	$err = $_.Exception.Message;
	$ms = [string]::Format("Exception occured: {0}", $err);
}
Finally {
	#region Close connection
	if ($pfcCompany.IsConnected) {
		$pfcCompany.Disconnect()
		write-host  –backgroundcolor green –foregroundcolor black "Disconnected from the company"
	}
	#endregion
}
