#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Resource Costing
########################################################################
$SCRIPT_VERSION = "3.1"
# Last tested PF version: ProcessForce 9.3 (9.30.210) (64-bit)
# Description:
#      Import Resource Costing. Script will update only existing Resource Costing Data. Remember to run Restore Resource Costing Details before running this script.
# Warning:
#   It's recommended run script when all users all disconnected.
#   Before running this script please do database backup.
#   Before running this script please restore Resource Costing Details. #
#   This script allows only to update Resource Costing on categories different than 000 #
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
# $csvImportCatalog = "C:\PS\PF\Costing\ResourceCosting";

$csvResourceCostingsPath = -join ($csvImportCatalog, "ResourceCosting.csv");
$csvResourceCostingDetailsPath = -join ($csvImportCatalog, "ResourceCostingDetails.csv");
$csvResourceCostingOverheadsPath = -join ($csvImportCatalog, "ResourceCostingOverheads.csv");

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


#region addional functions
function getResourceTimeType($CostType) {
	$ResourceTimeType = $null;
	switch ($CostType) {
		"QT" { $ResourceTimeType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.ResourceTimeType]::Queue; break; }
		"ST" { $ResourceTimeType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.ResourceTimeType]::Setup; break; }
		"RT" { $ResourceTimeType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.ResourceTimeType]::Run; break; }
		"TT" { $ResourceTimeType = [CompuTec.ProcessForce.API.Documents.Costing.ResourceCosting.ResourceTimeType]::Stock; break; }
		Default {
			$msg = [String]::Format("Incorrect Cost Type: '{0}'. Allowed values: QT - Queue Time, ST - Setup Time, RT - Run Time, TT - Stock Time.", [string]$CostType);
			throw [System.Exception] ($msg);
		}
	}
	return $ResourceTimeType;
}
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
	Write-Host 'Preparing data: '
	#region import csv files
	[array]$csvResourceCostings = Import-Csv -Delimiter ';' -Path $csvResourceCostingsPath;
	[array]$csvResourceCostingDetails = Import-Csv -Delimiter ';' -Path $csvResourceCostingDetailsPath;

	if ((Test-Path -Path $csvResourceCostingOverheadsPath -PathType leaf) -eq $true) {
		[array] $csvResourceCostingOverheads = Import-Csv -Delimiter ';' $csvResourceCostingOverheadsPath;
	}
	else {
		[array] $csvResourceCostingOverheads = $null; write-host "Resource Costing Overheads - csv not available."
	}
	
	$totalRows = $csvResourceCostings.Count + $csvResourceCostingDetails.Count + $csvResourceCostingOverheads.Count ;
	$dictionaryResourceCosting = New-Object 'System.Collections.Generic.Dictionary[string,psobject]';

	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}


	foreach ($row in $csvResourceCostings) {
		$key = $row.ResourceCode + '___' + $row.CostCategory;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if (-Not $dictionaryResourceCosting.ContainsKey($key)) {
			$dictionaryResourceCosting.Add($key, [psobject]@{
					ResourceCode = $row.ResourceCode
					CostCategory = $row.CostCategory
					Details      = New-Object 'System.Collections.Generic.List[array]'
					Overheads    = New-Object 'System.Collections.Generic.List[array]'
				});
		}
	}

	foreach ($row in $csvResourceCostingDetails) {
		$key = $row.ResourceCode + '___' + $row.CostCategory;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryResourceCosting.ContainsKey($key)) {
			$list = $dictionaryResourceCosting[$key].Details;
			$list.Add([array]$row);
		}
	}

	foreach ($row in $csvResourceCostingOverheads) {
		$key = $row.ResourceCode + '___' + $row.CostCategory;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryResourceCosting.ContainsKey($key)) {
			$list = $dictionaryResourceCosting[$key].Overheads;
			$list.Add([array]$row);
		}
	}
	#endregion

	Write-Host '';
	Write-Host 'Adding/updating data: ' -NoNewline;

	$totalRows = $dictionaryResourceCosting.Count
	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}

	foreach ($key in $dictionaryResourceCosting.Keys) {
		try {
			$surceResourceCosting = $dictionaryResourceCosting[$key];
			$progressItterator++;
			$progres = [math]::Round(($progressItterator * 100) / $total);
			if ($progres -gt $beforeProgress) {
				Write-Host $progres"% " -NoNewline
				$beforeProgress = $progres
			}

			#Creating Resource Costing Object
			$rc = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"ResourceCosting")
			if ($surceResourceCosting.CostCategory -eq '000') {
				Throw [System.Exception]("Masive update for Cost Category 000 is turned off - please make updates on custom Cost Category and use Roll-Over functionality");
			}
			#Checking if ResourceCosting exists
			$retValue = $rc.Get($surceResourceCosting.ResourceCode, $surceResourceCosting.CostCategory)
	
			if (-not $retValue) {
				Throw [System.Exception] ("Resource costing don't exists. Please make sure resource exists and/or Restore Resource Costing");
			}


			#Costing Details for positions from ResourceCostingDetails.csv file
			$csvCostingDetails = $surceResourceCosting.Details;
			if ($csvCostingDetails.count -gt 0) {
				foreach ($csvCD in $csvCostingDetails) {
					$count = $rc.Costs.Count
					for ($i = 0; $i -lt $count ; $i++) {
						$rc.Costs.SetCurrentLine($i);
						#QT - Queue Time, ST - Setup Time, RT - Run Time, TT - Stock Time
						if ( $rc.Costs.U_CostType -eq $csvCD.CostType) {
							$rc.Costs.U_HourRate = $csvCD.HourlyRate
							$rc.Costs.U_FixOH = $csvCD.FixedOH
							$rc.Costs.U_FixOHPrct = $csvCD.FixedOHPrct
							$rc.Costs.U_FixOHOther = $csvCD.FixedOHOther
							$rc.Costs.U_VarOH = $csvCD.VariableOH
							$rc.Costs.U_VarOHPrct = $csvCD.VariableOHPrct
							$rc.Costs.U_VarOHOther = $csvCD.VariableOHOther
							$rc.Costs.U_Remarks = $csvCD.Remarks
							break;
						}
					}
				}
			}

			#Multistructure Fixed and Variable Cost
			$csvOverheads = $surceResourceCosting.Overheads;
			if ($csvOverheads) {
				$newPositions = New-Object 'System.Collections.Generic.List[array]';
				foreach ($csvCO in $csvOverheads) {
					$ResourceTimeType = getResourceTimeType -CostType $csvCO.CostType;
					$OverheadType = getOverheadType -Type $csvCO.OverheadType;
					$OverheadSubType = getOverheadSubType -Type $csvCO.OverheadSubType;
					$count = $rc.OverheadCosts.Count;
					$existingOverhead = $rc.OverheadCosts.Where( { $_.U_ResourceTimeType -eq $ResourceTimeType -and $_.U_OverheadTypeCode -eq $csvCO.OverheadTypeCode -and $_.U_OverheadType -eq $OverheadType -and $_.U_OverheadSubtype -eq $OverheadSubType } );

					if ($existingOverhead.Count -gt 0) {
							$existingOverhead[0].U_Value = $csvCO.Value;
							$existingOverhead[0].U_OverheadTypeName = $csvCO.OverheadTypeName;
					}
					else {
						$newPositions.Add($csvCO);
					}
				}
				$rc.OverheadCosts.SetCurrentLine($rc.OverheadCosts.Count - 1);
				foreach ($csvCO in $newPositions) {
					if ($rc.OverheadCosts.IsRowFilled()) {
						$dummy = $rc.OverheadCosts.Add();
					}
					$ResourceTimeType = getResourceTimeType -CostType $csvCO.CostType;
					$OverheadType = getOverheadType -Type $csvCO.OverheadType;
					$OverheadSubType = getOverheadSubType -Type $csvCO.OverheadSubType;
					$rc.OverheadCosts.U_OverheadTypeCode = $csvCO.OverheadTypeCode;
					$rc.OverheadCosts.U_OverheadTypeName = $csvCO.OverheadTypeName;
					$rc.OverheadCosts.U_ResourceTimeType = $ResourceTimeType;
					$rc.OverheadCosts.U_OverheadType = $OverheadType;
					$rc.OverheadCosts.U_OverheadSubtype = $OverheadSubType;
					$rc.OverheadCosts.U_Value = $csvCO.Value;
				}
			}
			$message = 0
			$message = $rc.Update()
			if ($message -lt 0) {  
				$err = $pfcCompany.GetLastErrorDescription()
				Throw [System.Exception] ($err);
			}
		}
		Catch {
			$err = $_.Exception.Message;
			$ms = [string]::Format("Error when updating Resource Costing Details for Resource {0}, Cost Category {1}, Details: {2}", $surceResourceCosting.ResourceCode, $surceResourceCosting.CostCategory, $err);
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
