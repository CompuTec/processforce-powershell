
#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Item Costing
########################################################################
$SCRIPT_VERSION = "3.1"
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
# $csvImportCatalog = "C:\PS\PF\Costing\ItemCosting";

$csvItemCostingsPath = -join ($csvImportCatalog, "ItemCosting.csv");
$csvItemCostingDetailsPath = -join ($csvImportCatalog, "ItemCostingDetails.csv");
$csvItemCostingCoproductsPath = -join ($csvImportCatalog, "ItemCostingCoproducts.csv");
$csvItemCostingScrapsPath = -join ($csvImportCatalog, "ItemCostingScraps.csv");
$csvItemCostingOverheadsPath = -join ($csvImportCatalog, "ItemCostingOverheads.csv");

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
	$code = $pfcCompany.Connect();
 
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

#region addional functions
function getOverheadType($Type) {
	$OverheadType = $null;
	switch ($Type) {
		"F" { $OverheadType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.OverheadType]::Fixed; break; }
		"V" { $OverheadType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.OverheadType]::Variable; break; }
		Default {
			$msg = [String]::Format("Incorrect Overhead Type: '{0}'. Allowed values: F - Fixed, V - Variable.", [string]$Type);
			throw [System.Exception] ($msg);
		}
	}
	return $OverheadType;
}
function getOverheadSubType($Type) {
	$OverheadType = $null;
	switch ($Type) {
		"F" { $OverheadType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.OverheadSubType]::Fixed; break; }
		"FP" { $OverheadType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.OverheadSubType]::FixedPercentage; break; }
		"FO" { $OverheadType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.OverheadSubType]::FixedOther; break; }
		"V" { $OverheadType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.OverheadSubType]::Variable; break; }
		"VP" { $OverheadType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.OverheadSubType]::VariablePercentage; break; }
		"VO" { $OverheadType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.OverheadSubType]::VariableOther; break; }
		Default {
			$msg = [String]::Format("Incorrect Overhead SubType: '{0}'. Allowed values: F - Fixed, FP - Fixed Percentage, FO - Fixed Other, V - Variable, VP - Variable Percentage, VO - Variable Other", [string]$Type);
			throw [System.Exception] ($msg);
		}
	}
	return $OverheadType;
}
#endregion


try {

	#Data loading from a csv file
	Write-Host 'Preparing data: ' -NoNewline;
	#region import csv files
	[array]$csvItemCostings = Import-Csv -Delimiter ';' -Path $csvItemCostingsPath;
	[array]$csvItemCostingDetails = Import-Csv -Delimiter ';' -Path $csvItemCostingDetailsPath;
	[array]$csvItemCoproductsDetails = Import-Csv -Delimiter ';' -Path $csvItemCostingCoproductsPath;

	if ((Test-Path -Path $csvItemCostingCoproductsPath -PathType leaf) -eq $true) {
		[array] $csvItemCoproductsDetails = Import-Csv -Delimiter ';' $csvItemCostingCoproductsPath; 
	}
	else {
		[array] $csvItemCoproductsDetails = $null; write-host "Item Costing CoProducts Details - csv not available."
	}

	[array]$csvItemScrapsDetails = Import-Csv -Delimiter ';' -Path $csvItemCostingScrapsPath;
	if ((Test-Path -Path $csvItemCostingScrapsPath -PathType leaf) -eq $true) {
		[array] $csvItemScrapsDetails = Import-Csv -Delimiter ';' $csvItemCostingScrapsPath; 
	}
	else {
		[array] $csvItemScrapsDetails = $null; write-host "Item Costing Scraps Details - csv not available."
	}
	
	if ((Test-Path -Path $csvItemCostingOverheadsPath -PathType leaf) -eq $true) {
		[array] $csvItemCostingOverheads = Import-Csv -Delimiter ';' $csvItemCostingOverheadsPath;
	}
	else {
		[array] $csvItemCostingOverheads = $null; write-host "Item Costing Overheads - csv not available."
	}
	
	$totalRows = $csvItemCostings.Count + $csvItemCostingDetails.Count + $csvItemCostingOverheads.Count + $csvItemCoproductsDetails.Count + $csvItemScrapsDetails.Count; 
	
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
					Coproducts   = New-Object 'System.Collections.Generic.List[array]'
					Scraps       = New-Object 'System.Collections.Generic.List[array]'
					Overheads    = New-Object 'System.Collections.Generic.List[array]'
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

	foreach ($row in $csvItemCoproductsDetails) {
		$key = $row.ItemCode + '__' + $row.Revision + '__' + $row.Category;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryItemCosting.ContainsKey($key)) {
			$list = $dictionaryItemCosting[$key].Coproducts;
			$list.Add([array]$row);
		}
	}

	foreach ($row in $csvItemScrapsDetails) {
		$key = $row.ItemCode + '__' + $row.Revision + '__' + $row.Category;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryItemCosting.ContainsKey($key)) {
			$list = $dictionaryItemCosting[$key].Scraps;
			$list.Add([array]$row);
		}
	}
	
	foreach ($row in $csvItemCostingOverheads) {
		$key = $row.ItemCode + '__' + $row.Revision + '__' + $row.Category;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryItemCosting.ContainsKey($key)) {
			$list = $dictionaryItemCosting[$key].Overheads;
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
			$csvCostingDetails = $csvItemCosting.Details
			if ($csvCostingDetails.Count -eq 0) {
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

			$csvCoproducts = $csvItemCosting.Coproducts;
			if ($csvCoproducts.Count -gt 0) {
				foreach ($csvCP in $csvCoproducts) {
					$matchingCPs = $ic.Coproducts.Where( { $_.U_SequenceNo -eq $csvCP.Sequence -and $_.U_WhsCode -eq $csvCP.WhsCode });
					if ($matchingCPs.Count -gt 0) {
						$cp = $matchingCPs[0];
					}
					else {
						$msg = [string]::Format("Item Costing Coproducts Line with sequence {0} and WhsCode {1} is missing.", [string]$csvCP.Sequence, [string]$csvCP.WhsCode);
						throw [System.Exception]($msg);
					}
					
					$cp.U_Type = $csvCP.Type;
					$cp.U_PriceList = $csvCP.PriceListCode;
					$cp.U_WhenZero = $csvCP.WhenZero;
					$cp.U_ItemCost = $csvCP.ItemCost;
					$cp.U_RscCost = $csvCP.ResourceCost;
					$cp.U_FixOH = $csvCP.FixedOH;
					$cp.U_VarOH = $csvCP.VariableOH;
					$cp.U_Remarks = $csvCP.Remarks;
				}
			}

			$csvScraps = $csvItemCosting.Scraps;
			if ($csvScraps.Count -gt 0) {
				foreach ($csvSC in $csvScraps) {
					$matchingSCs = $ic.Scrap.Where( { $_.U_SequenceNo -eq $csvSC.Sequence -and $_.U_WhsCode -eq $csvSC.WhsCode });
					if ($matchingSCs.Count -gt 0) {
						$sc = $matchingSCs[0];
					}
					else {
						$msg = [string]::Format("Item Costing Scraps Line with sequence {0} and WhsCode {1} is missing.", [string]$csvCP.Sequence, [string]$csvCP.WhsCode);
						throw [System.Exception]($msg);
					}
					
					$sc.U_Type = $csvSC.Type;
					$sc.U_PriceList = $csvSC.PriceListCode;
					$sc.U_WhenZero = $csvSC.WhenZero;
					$sc.U_ItemCost = $csvSC.ItemCost;
					$sc.U_RscCost = $csvSC.ResourceCost;
					$sc.U_FixOH = $csvSC.FixedOH;
					$sc.U_VarOH = $csvSC.VariableOH;
					$sc.U_Remarks = $csvSC.Remarks;
				}
			}

			#Multistructure Fixed and Variable Cost
			$csvOverheads = $csvItemCosting.Overheads;
			if ($csvOverheads) {
				$newPositions = New-Object 'System.Collections.Generic.List[array]';
				foreach ($csvCO in $csvOverheads) {
					$OverheadType = getOverheadType -Type $csvCO.OverheadType;
					$OverheadSubType = getOverheadSubType -Type $csvCO.OverheadSubType;
					$count = $ic.OverheadCosts.Count;
					$existingOverhead = $ic.OverheadCosts.Where( { $_.U_WhsCode -eq $csvCO.WhsCode -and $_.U_OverheadTypeCode -eq $csvCO.OverheadTypeCode -and $_.U_OverheadType -eq $OverheadType -and $_.U_OverheadSubtype -eq $OverheadSubType } );

					if ($existingOverhead.Count -gt 0) {
						$existingOverhead[0].U_Value = $csvCO.Value;
						$existingOverhead[0].U_OverheadTypeName = $csvCO.OverheadTypeName;
					}
					else {
						$newPositions.Add($csvCO);
					}
				}
				$ic.OverheadCosts.SetCurrentLine($ic.OverheadCosts.Count - 1);
				foreach ($csvCO in $newPositions) {
					if ($ic.OverheadCosts.IsRowFilled()) {
						$dummy = $ic.OverheadCosts.Add();
					}
					$OverheadType = getOverheadType -Type $csvCO.OverheadType;
					$OverheadSubType = getOverheadSubType -Type $csvCO.OverheadSubType;
					$ic.OverheadCosts.U_OverheadTypeCode = $csvCO.OverheadTypeCode;
					$ic.OverheadCosts.U_OverheadTypeName = $csvCO.OverheadTypeName;
					$ic.OverheadCosts.U_WhsCode = $csvCO.WhsCode
					$ic.OverheadCosts.U_OverheadType = $OverheadType;
					$ic.OverheadCosts.U_OverheadSubtype = $OverheadSubType;
					$ic.OverheadCosts.U_Value = $csvCO.Value;
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
		Finally {
			$progressItterator++;
			$progres = [math]::Round(($progressItterator * 100) / $total);
			if ($progres -gt $beforeProgress) {
				Write-Host $progres"% " -NoNewline
				$beforeProgress = $progres
			}
		}
	}
}
Catch {
	$err = $_.Exception.Message;
	$ms = [string]::Format("Exception occured: {0}", $err);
	Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
}
Finally {
	#region Close connection
	if ($pfcCompany.IsConnected) {
		$pfcCompany.Disconnect()
		write-host  –backgroundcolor green –foregroundcolor black "Disconnected from the company"
	}
	#endregion
}
