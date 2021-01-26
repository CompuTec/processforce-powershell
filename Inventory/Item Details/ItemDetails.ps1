#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Item Details
########################################################################
$SCRIPT_VERSION = "3.7"
# Last tested PF version: ProcessForce 9.3 (9.30.210) PL: MAIN (64-bit)
# Description:
#      Import Item Details. Script will update only existing ItemDetails. Remember to run Restore Item Details before running this script.
#      Only required csv for this file is ItemDetails.csv.
#      If csv file for given section is not presented or don't have related records then on ItemDetails this section will be ommited. 
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
# $csvImportCatalog = "C:\PS\PF\TestProtocols\";

$csvItemDetailsPath = -join ($csvImportCatalog, "ItemDetails.csv")
$csvItemDetailsBatchDetailsPath = -join ($csvImportCatalog, "ItemDetailsBatchDetails.csv")
$csvItemDetailsClassificationPath = -join ($csvImportCatalog, "ItemDetailsClassification.csv")
$csvItemDetailsGroupsPath = -join ($csvImportCatalog, "ItemDetailsGroups.csv")
$csvItemDetailsOriginsPath = -join ($csvImportCatalog, "ItemDetailsOrigins.csv")
$csvItemDetailsPhrasesPath = -join ($csvImportCatalog, "ItemDetailsPhrases.csv")
$csvItemDetailsPropertiesPath = -join ($csvImportCatalog, "ItemDetailsProperties.csv")
$csvItemDetailsPropertiesCertifiacteOfAnalysisPath = -join ($csvImportCatalog, "ItemDetailsPropertiesCertifiacteOfAnalysis.csv")
$csvItemDetailsRevisionsPath = -join ($csvImportCatalog, "ItemDetailsRevisions.csv")
$csvItemDetailsTextsPath = -join ($csvImportCatalog, "ItemDetailsTexts.csv")
$csvItemDetailsPlanningDataPath = -join ($csvImportCatalog, "ItemDetailsPlanningData.csv")

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
	$atLeastOneSuccess = $false;
	#Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
	Write-Host 'Preparing data: '
	#region import csv files
	[array]$csvItemDetails = Import-Csv -Delimiter ';' -Path $csvItemDetailsPath
    
	if ((Test-Path -Path $csvItemDetailsBatchDetailsPath -PathType leaf) -eq $true) {
		[array] $csvItemDetailsBatchDetails = Import-Csv -Delimiter ';' $csvItemDetailsBatchDetailsPath; 
	}
	else {
		[array] $csvItemDetailsBatchDetails = $null; write-host "Item Details Batch Details - csv not available."
	}
	if ((Test-Path -Path $csvItemDetailsClassificationPath -PathType leaf) -eq $true) {
		[array] $csvItemDetailsClassification = Import-Csv -Delimiter ';' $csvItemDetailsClassificationPath; 
	}
	else {
		[array] $csvItemDetailsClassification = $null; write-host "Item Details Classifications - csv not available."
	}
	if ((Test-Path -Path $csvItemDetailsGroupsPath -PathType leaf) -eq $true) {
		[array] $csvItemDetailsGroups = Import-Csv -Delimiter ';' $csvItemDetailsGroupsPath; 
	}
	else {
		[array] $csvItemDetailsGroups = $null; write-host "Item Details Groups - csv not available."
	}
	if ((Test-Path -Path $csvItemDetailsOriginsPath -PathType leaf) -eq $true) {
		[array] $csvItemDetailsOrigins = Import-Csv -Delimiter ';' $csvItemDetailsOriginsPath; 
	}
	else {
		[array] $csvItemDetailsOrigins = $null; write-host "Item Details Origins - csv not available."
	}
	if ((Test-Path -Path $csvItemDetailsPhrasesPath -PathType leaf) -eq $true) {
		[array] $csvItemDetailsPhrases = Import-Csv -Delimiter ';' $csvItemDetailsPhrasesPath;
	}
	else {
		[array] $csvItemDetailsPhrases = $null; write-host "Item Details Phrases - csv not available."
	}
	if ((Test-Path -Path $csvItemDetailsPropertiesPath -PathType leaf) -eq $true) {
		[array] $csvItemDetailsProperties = Import-Csv -Delimiter ';' $csvItemDetailsPropertiesPath;
	}
	else {
		[array] $csvItemDetailsProperties = $null; write-host "Item Details Properties - csv not available."
	}
	if ((Test-Path -Path $csvItemDetailsPropertiesCertifiacteOfAnalysisPath -PathType leaf) -eq $true) {
		[array] $csvItemDetailsPropertiesCerts = Import-Csv -Delimiter ';' $csvItemDetailsPropertiesCertifiacteOfAnalysisPath;
	}
	else {
		[array] $csvItemDetailsPropertiesCerts = $null; write-host "Item Details Properties Certificate of Analysis - csv not available."
	}
	if ((Test-Path -Path $csvItemDetailsRevisionsPath -PathType leaf) -eq $true) {
		[array] $csvItemDetailsRevisions = Import-Csv -Delimiter ';' $csvItemDetailsRevisionsPath; 
	}
	else {
		[array] $csvItemDetailsRevisions = $null; write-host "Item Details Revisions - csv not available."
	}
	if ((Test-Path -Path $csvItemDetailsTextsPath -PathType leaf) -eq $true) {
		[array] $csvItemDetailsTexts = Import-Csv -Delimiter ';' $csvItemDetailsTextsPath;
	}
	else {
		[array] $csvItemDetailsTexts = $null; write-host "Item Details Texts - csv not available."
	}
	if ((Test-Path -Path $csvItemDetailsPlanningDataPath -PathType leaf) -eq $true) {
		[array] $csvItemDetailsPlanningData = Import-Csv -Delimiter ';' $csvItemDetailsPlanningDataPath;
	}
	else {
		[array] $csvItemDetailsPlanningData = $null; write-host "Item Details Planning Data - csv not available."
	}
	#endregion

	$totalRows = $csvItemDetails.Count + $csvItemDetailsBatchDetails.Count + $csvItemDetailsClassification.Count + $csvItemDetailsGroups.Count + $csvItemDetailsOrigins.Count;
	$totalRows += $csvItemDetailsPhrases.Count + $csvItemDetailsProperties.Count + $csvItemDetailsRevisions.Count + $csvItemDetailsTexts.Count + $csvItemDetailsPropertiesCerts.Count + $csvItemDetailsPlanningData.Count;

	$itemDetailsList = New-Object 'System.Collections.Generic.List[array]';
	$dictItemDetailsBatchDetails = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]';
	$dictItemDetailsClassification = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]';
	$dictItemDetailsGroups = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]';
	$dictItemDetailsOrigins = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]';
	$dictItemDetailsPhrases = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]';
	$dictItemDetailsProperties = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]';
	$dictItemDetailsPropertiesCerts = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]';
	$dictItemDetailsRevisions = New-Object 'System.Collections.Generic.Dictionary[string,psobject]';
	$dictItemDetailsTexts = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]';
	$dictItemDetailsPlanningData = New-Object 'System.Collections.Generic.Dictionary[string,psobject]';

	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}
	#region parseInputData
	foreach ($row in $csvItemDetails) {
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		$itemDetailsList.Add([array]$row);
	}

	foreach ($row in $csvItemDetailsBatchDetails) {
		$key = $row.ItemCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($dictItemDetailsBatchDetails.ContainsKey($key) -eq $false) {
			$dictItemDetailsBatchDetails[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $dictItemDetailsBatchDetails[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvItemDetailsClassification) {
		$key = $row.ItemCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($dictItemDetailsClassification.ContainsKey($key) -eq $false) {
			$dictItemDetailsClassification[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $dictItemDetailsClassification[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvItemDetailsGroups) {
		$key = $row.ItemCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($dictItemDetailsGroups.ContainsKey($key) -eq $false) {
			$dictItemDetailsGroups[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $dictItemDetailsGroups[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvItemDetailsOrigins) {
		$key = $row.ItemCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($dictItemDetailsOrigins.ContainsKey($key) -eq $false) {
			$dictItemDetailsOrigins[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $dictItemDetailsOrigins[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvItemDetailsPhrases) {
		$key = $row.ItemCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($dictItemDetailsPhrases.ContainsKey($key) -eq $false) {
			$dictItemDetailsPhrases[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $dictItemDetailsPhrases[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvItemDetailsProperties) {
		$key = $row.ItemCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($dictItemDetailsProperties.ContainsKey($key) -eq $false) {
			$dictItemDetailsProperties[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $dictItemDetailsProperties[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvItemDetailsPropertiesCerts) {
		$key = $row.ItemCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($dictItemDetailsPropertiesCerts.ContainsKey($key) -eq $false) {
			$dictItemDetailsPropertiesCerts[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $dictItemDetailsPropertiesCerts[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvItemDetailsRevisions) {
		$key = $row.ItemCode;
		$revCode = $row.RevisionCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictItemDetailsRevisions.ContainsKey($key) -eq $false) {
			$dictItemDetailsRevisions.Add($key, (New-Object 'System.Collections.Generic.Dictionary[string,array]'));
		}

		if ($dictItemDetailsRevisions[$key].ContainsKey($revCode) -eq $false) {
			$dictItemDetailsRevisions[$key].Add($revCode, [array]$row);
		}
	}
	foreach ($row in $csvItemDetailsTexts) {
		$key = $row.ItemCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($dictItemDetailsTexts.ContainsKey($key) -eq $false) {
			$dictItemDetailsTexts[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $dictItemDetailsTexts[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvItemDetailsPlanningData) {
		$key = $row.ItemCode;
		$revCode = $row.RevisionCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictItemDetailsPlanningData.ContainsKey($key) -eq $false) {
			$dictItemDetailsPlanningData.Add($key, (New-Object 'System.Collections.Generic.Dictionary[string,array]'));
		}

		if ($dictItemDetailsPlanningData[$key].ContainsKey($revCode) -eq $false) {
			$dictItemDetailsPlanningData[$key].Add($revCode, [array]$row);
		}
	}
	#endregion
    
	Write-Host '';
	Write-Host 'Adding/updating data: ' -NoNewline;
	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	if ($itemDetailsList.Count -gt 1) {
		$total = $itemDetailsList.Count
	}
	else {
		$total = 1
	}

	#checking state of DirectAccess mode
	$DirectAccessState = [CompuTec.Core.DI.Database.DataLayer]::GetLayerState($pfcCompany.Token);
	if ($DirectAccessState -eq [CompuTec.Core.DI.Database.DataLayerState]::"Connected") {
		$DIRECT_ACCESS_MODE = $true;
		[CompuTec.ProcessForce.API.DynamicCosting.Restoration.ItemCostingRestorationManager]::SuspendItemCostingRestoration = $true
	}
	else {
		$DIRECT_ACCESS_MODE = $false;
		write-host -backgroundcolor yellow -foregroundcolor blue "Please enable Direct Access Mode to speed up import process";
	}
    
	foreach ($csvItem in $itemDetailsList) {
		try {
			$key = $csvItem.ItemCode
			$progressItterator++;
			$progres = [math]::Round(($progressItterator * 100) / $total);
			if ($progres -gt $beforeProgress) {
				Write-Host $progres"% " -NoNewline
				$beforeProgress = $progres
			}
			#Checking that Item Details already exist
			$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]::"BoRecordset")
			[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
			$dummy = $rs.DoQuery([string]::Format( "SELECT T0.""ItemCode"" AS ""ItemCode"" FROM OITM T0
                INNER JOIN ""@CT_PF_OIDT"" T1 ON T0.""ItemCode"" = T1.""U_ItemCode"" WHERE T0.""ItemCode"" = N'{0}'", $key))
        
			if ($rs.RecordCount -eq 0) {
				$err = [string]::Format("Item Master Data with ItemCode {0} don't exists. Please restore Item Details", $key);
				Throw [System.Exception] ($err)
			}
   
			#Creating Item Details Object
			$idt = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::"ItemDetails")
      
			#Restoring Item Costs and setting Inherit Batch/Serial to 'Yes'
			$dummy = $idt.GetByItemCode($key);
			if([string]::IsNullOrWhiteSpace($csvItem.Yield) -eq $false){
				$idt.U_Yield = $csvItem.Yield
			} else {
				$idt.U_Yield = 100;
			}

			$idt.U_IgnoreYield = $csvItem.IgnoreYield
			$idt.U_DftOrigin = $csvItem.DefaultOrigin
			$idt.U_AcptLowerQty = $csvItem.AllowResidualQty

			#region Revisions
			try {
				$revisions = $dictItemDetailsRevisions[$key];
				if ($revisions.count -gt 0) {
    
					$linesToBeRemoved = New-Object 'System.Collections.Generic.List[int]';
					$currentPosDict = New-Object 'System.Collections.Generic.Dictionary[string,int]';

					$revIndex = 0;
					foreach ($revision in $idt.Revisions) {
						if ($revision.U_Code -gt "") {
							if ($revisions.ContainsKey($revision.U_Code)) {
								$currentPosDict.Add($revision.U_Code, $revIndex);
							}
							else {
								$linesToBeRemoved.Add($revIndex);
							}
						}
						$revIndex++;
					}
          
					#updating existing revision
					foreach ($revCode in $revisions.Keys) {
						$rev = $revisions[$revCode];
						if ($currentPosDict.ContainsKey($revCode)) {
							$idt.Revisions.SetCurrentLine($currentPosDict[$revCode]);
						}
						else {
							$idt.Revisions.SetCurrentLine($idt.Revisions.Count - 1);
							if($idt.Revisions.IsRowFilled()){
								$dummy = $idt.Revisions.Add();
							}
							$idt.Revisions.U_Code = $rev.RevisionCode
						}

						$idt.Revisions.U_Description = $rev.RevisionName
						$idt.Revisions.U_Status = $rev.Status #enum type; Revision Status, Active ACT = 1, BeingPhasedOut BPO = 2, Engineering ENG = 3, Obsolete OBS = 4
						if ($rev.ValidFrom -gt '') {
							$idt.Revisions.U_ValidFrom = $rev.ValidFrom
						}
						if ($rev.ValidTo -gt '') {
							$idt.Revisions.U_ValidTo = $rev.ValidTo
						}
						$idt.Revisions.U_Remarks = $rev.Remarks
						if($rev.IsDefault -eq 1 -or $rev.IsDefault -eq "Y"){
							$idt.Revisions.U_Default = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes
						} else {
							$idt.Revisions.U_Default = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
						}
						if($rev.IsMRPDefault -eq 1 -or $rev.IsMRPDefault -eq "Y"){
							$idt.Revisions.U_IsMRPDefault = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes
						} else {
							$idt.Revisions.U_IsMRPDefault = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
						}
						if($rev.DefaultForCosting -eq 1 -or $rev.DefaultForCosting -eq "Y"){
							$idt.Revisions.U_IsCostingDefault = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes
						} else {
							$idt.Revisions.U_IsCostingDefault = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
						}
						

						#region Planning Data
						$planningInfo = $dictItemDetailsPlanningData[$key];
						if($planningInfo.ContainsKey($revCode) -eq $true){
							$revPlanInf = $planningInfo[$revCode];
							$idt.Revisions.U_PlanningMethod = $revPlanInf.PlanningMethod;
							$idt.Revisions.U_ProcurementMethod = $revPlanInf.ProcurementMethod;
							$idt.Revisions.U_OrderInterval = $revPlanInf.OrderInterval;
							$idt.Revisions.U_OrderMultiple = $revPlanInf.OrderMultiple;
							$idt.Revisions.U_LeadTime = $revPlanInf.LeadTime;
							$idt.Revisions.U_ToleranceDays = $revPlanInf.ToleranceDays;
							$idt.Revisions.U_ILeadTime = $revPlanInf.ILeadTime;
							$idt.Revisions.U_MinOrderQty = $revPlanInf.MinOrderQty;
							$idt.Revisions.U_MaxOrderQty = $revPlanInf.MaxOrderQty;
							if($revPlanInf.ForcePrimaryDemand -eq "Y") {
								$idt.Revisions.U_ForcePrimaryDemand = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes
							} else {
								$idt.Revisions.U_ForcePrimaryDemand = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
							}
							if($revPlanInf.UseItmPerPurchUnit -eq "Y") {
								$idt.Revisions.U_UseItmPerPurchUnit = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes
							} else {
								$idt.Revisions.U_UseItmPerPurchUnit = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
							}
							$idt.Revisions.U_InternalLeadTimeTransfer = $revPlanInf.InternalLeadTimeTransfer;
							$idt.Revisions.U_PlanerI = $revPlanInf.PlanerI;
							$idt.Revisions.U_PlanerII = $revPlanInf.PlanerII;
						}
						#endregion
						if ($currentPosDict.ContainsKey($revCode) -eq $false) {
							$dummy = $idt.Revisions.Add();
						}
					}
					#Deleting revision
					for ($idxD = $linesToBeRemoved.Count - 1; $idxD -ge 0; $idxD--) {
						[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
						$dummy = $idt.Revisions.DelRowAtPos($linesToBeRemoved[$idxD]);
					}
          
				} 
			}
			Catch {
				$err = $_.Exception.Message;
				$ms = [string]::Format("Error when updating Revisions for Item Details with ItemCode {0} Details: {1}", $key, $err);
				Throw [System.Exception]($ms);
			}
			#endregion

			#region BatchDetails
			try {
				[array] $batchDetails = $dictItemDetailsBatchDetails[$key];
				if ($batchDetails.count -gt 0) {
					if ( $batchDetails.BatchInherit -eq 0) {
						$idt.U_InheritBatch = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
						$idt.U_BtchTmpl = $batchDetails.BatchTemplate;
					}
					else {
						$idt.U_InheritBatch = 1
					}
                    
					if ( $batchDetails.SerialIncherit -eq 0) {
						$idt.U_InheritSerial = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
						$idt.U_SrlTmpl = $batchDetails.SerialTemplate;
					}
					else {
						$idt.U_InheritSerial = 1;
					}
                    
					if ( $batchDetails.ExpiryInherit -eq 0) {
						$idt.U_Inherit = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
						if ( $batchDetails.Expiry -eq 1 ) {
							$idt.U_Expiry	= [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
							$idt.U_ExpWarn = $batchDetails.ExpiryWarning;
						}
                        
						if ( $batchDetails.Consume -eq 1) {
							$idt.U_Consume = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
							$idt.U_ConsWarn = $batchDetails.ConsWarn;
						}
						if ( $batchDetails.ShelfLife -ne "") {
							$idt.U_ShelfTime = $batchDetails.ShelfLife;
						}
						if ( $batchDetails.InspectionInterval -ne "") {
							$idt.U_InspDays = $batchDetails.InspectionInterval;
						}
                        
					}
					else {
						$idt.U_Inherit = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
					}
                    
					if ( $batchDetails.InheritBatchQueue -eq 0) {
						$idt.U_InheritQueue = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
						$idt.U_BtchTmpl = $batchDetails.BatchTemplate;
                        
                        
						if ( $batchDetails.BatchQueue -eq "F") {
							$idt.U_BatchQueue = [CompuTec.ProcessForce.API.Enumerators.BatchQueueType]::FIFO;
						}
                        
						if ( $batchDetails.BatchQueue -eq "E") {
							$idt.U_BatchQueue = [CompuTec.ProcessForce.API.Enumerators.BatchQueueType]::FEFO;
						}
                            
						if ( $batchDetails.BatchQueue -eq "M") {
							$idt.U_BatchQueue = [CompuTec.ProcessForce.API.Enumerators.BatchQueueType]::FMFO;
						}
					}
					else {
						$idt.U_InheritQueue = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
					}
                    
					$idt.U_ExpTmpl = "";
                
					if ( $batchDetails.ExpTyp -eq "C")
					{ $idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::CreateDate }
                
					if ( $batchDetails.ExpTyp -eq "N" -or $batchDetails.ExpTyp -eq "" ) {
						$idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::CurrentDate 
					}
                
					if ( $batchDetails.ExpTyp -eq "E")
					{ $idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::EndDate }
                
					if ( $batchDetails.ExpTyp -eq "Q") { 
						$idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::Query 
						$idt.U_ExpTmpl = $batchDetails.ExpTempl;
					}
                
					if ( $batchDetails.ExpTyp -eq "R")
					{ $idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::RequiredDate }
                
					if ( $batchDetails.ExpTyp -eq "S")
					{ $idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::StartDate }
                    
					if ( $batchDetails.InheritStatus -eq 0) {
						$idt.U_InheritStatus = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
						if ( $batchDetails.U_SapDfBS -eq 'R') {
							$idt.U_SapDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Released;
						}
						if ( $batchDetails.U_SapDfBS -eq 'L') {
							$idt.U_SapDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Locked;
						}
						if ( $batchDetails.U_SapDfBS -eq 'A') {
							$idt.U_SapDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::NotAccesible;
						}
                
						if ( $batchDetails.U_SapDfQCS -eq 'F') {
							$idt.U_SapDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Failed
						}
						if ( $batchDetails.U_SapDfQCS -eq 'H') {
							$idt.U_SapDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::OnHold
						}
						if ( $batchDetails.U_SapDfQCS -eq 'I') {
							$idt.U_SapDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Inspection
						}
						if ( $batchDetails.U_SapDfQCS -eq 'P') {
							$idt.U_SapDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Passed
						}
						if ( $batchDetails.U_SapDfQCS -eq 'T') {
							$idt.U_SapDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::QCTesting
						}
                
						if ( $batchDetails.U_PFDfBS -eq 'R') {
							$idt.U_PFDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Released;
						}
						if ( $batchDetails.U_PFDfBS -eq 'L') {
							$idt.U_PFDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Locked;
						}
						if ( $batchDetails.U_PFDfBS -eq 'A') {
							$idt.U_PFDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::NotAccesible;
						}
                
						if ( $batchDetails.U_PFDfQCS -eq 'F') {
							$idt.U_PFDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Failed
						}
						if ( $batchDetails.U_PFDfQCS -eq 'H') {
							$idt.U_PFDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::OnHold
						}
						if ( $batchDetails.U_PFDfQCS -eq 'I') {
							$idt.U_PFDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Inspection
						}
						if ( $batchDetails.U_PFDfQCS -eq 'P') {
							$idt.U_PFDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Passed
						}
						if ( $batchDetails.U_PFDfQCS -eq 'T') {
							$idt.U_PFDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::QCTesting
						}
					}
					else {
						$idt.U_InheritStatus = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
					}
				} 
			}
			Catch {
				$err = $_.Exception.Message;
				$ms = [string]::Format("Error when updating Batch Details for Item Details with ItemCode {0} Details: {1}", $key, $err);
				Throw [System.Exception]($ms);
			}
			#endregion

			#region Classifications
			try {
				$classifications = $dictItemDetailsClassification[$key];
				if ($classifications.count -gt 0) {
					#Deleting all exisitng Classification
					$count = $idt.Classifications.Count
					for ($i = 0; $i -lt $count; $i++) {
						$dummy = $idt.Classifications.DelRowAtPos(0);
					}
					$idt.Classifications.SetCurrentLine($idt.Classifications.Count - 1);
 
					#Adding classifications
					foreach ($classification in $classifications) {
						$idt.Classifications.U_ClsCode = $classification.ClassificationCode;
						if ($classification.ProductionOrders -eq 1) {
							$idt.Classifications.U_ProdOrders = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
						}
						if ($classification.ShipmentDocuments -eq 1) {
							$idt.Classifications.U_ShipDoc = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
						}	
						if ( $classification.PickLists -eq 1) {
							$idt.Classifications.U_PickLists = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
						}
						if ($classification.MSDS -eq 1) {
							$idt.Classifications.U_MSDS = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
						}	
						if ( $classification.PurchaseOrders -eq 1) {
							$idt.Classifications.U_PurOrders = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
						}	
						if ($classification.Returns -eq 1) {
							$idt.Classifications.U_Returns = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
						}	
						if ($classification.Other -eq 1) {
							$idt.Classifications.U_Other = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
						}
						$dummy = $idt.Classifications.Add()
					}
        
				} 
			}
			Catch {
				$err = $_.Exception.Message;
				$ms = [string]::Format("Error when updating Classifications for Item Details with ItemCode {0} Details: {1}", $key, $err);
				Throw [System.Exception]($ms);
			}
			#endregion
            
			#region Groups
			try {
				[array] $groups = $dictItemDetailsGroups[$key];
				if ($groups.count -gt 0) {
					#Deleting all exisitng Groups
					$count = $idt.Groups.Count
					for ($i = 0; $i -lt $count; $i++) {
						$dummy = $idt.Groups.DelRowAtPos(0);
					}
					$idt.Groups.SetCurrentLine($idt.Groups.Count - 1);
     
					#Adding Groups
					foreach ($group in $groups) {
						$idt.Groups.U_GrpCode = $group.GroupCode;
						if ($group.ProductionOrders -eq 1) {
							$idt.Groups.U_ProdOrders = "Y"
						}
						if ($group.ShipmentDocuments -eq 1) {
							$idt.Groups.U_ShipDoc = "Y"
						}	
						if ( $group.PickLists -eq 1) {
							$idt.Groups.U_PickLists = "Y"
						}
						if ($group.MSDS -eq 1) {
							$idt.Groups.U_MSDS = "Y"
						}	
						if ( $group.PurchaseOrders -eq 1) {
							$idt.Groups.U_PurOrders = "Y"
						}	
						if ($group.Returns -eq 1) {
							$idt.Groups.U_Returns = "Y"
						}	
						if ($group.Other -eq 1) {
							$idt.Groups.U_Other = "Y"
						}
						$dummy = $idt.Groups.Add()
					}
				} 
			}
			Catch {
				$err = $_.Exception.Message;
				$ms = [string]::Format("Error when updating Groups for Item Details with ItemCode {0} Details: {1}", $key, $err);
				Throw [System.Exception]($ms);
			}
			#endregion

			#region Origins
			try {
				[array] $origins = $dictItemDetailsOrigins[$key];
				if ($origins.count -gt 0) {
					#Deleting all exisitng Origins
					$count = $idt.Origins.Count
					for ($i = 0; $i -lt $count; $i++) {
						$dummy = $idt.Origins.DelRowAtPos(0);
					}
					$idt.Origins.SetCurrentLine($idt.Origins.Count - 1);

					#Adding Origins
					foreach ($origin in $origins) {
						$idt.Origins.U_CountryCode = $origin.CountryCode;
						$dummy = $idt.Origins.Add()
					}
				} 
			}
			Catch {
				$err = $_.Exception.Message;
				$ms = [string]::Format("Error when updating Origins for Item Details with ItemCode {0} Details: {1}", $key, $err);
				Throw [System.Exception]($ms);
			}
			#endregion

			#region Phrases
			try {
				[array] $phrases = $dictItemDetailsPhrases[$key];
				if ($phrases.count -gt 0) {
					#Deleting all exisitng Phrases
					$count = $idt.Phrases.Count
					for ($i = 0; $i -lt $count; $i++) {
						$dummy = $idt.Phrases.DelRowAtPos(0);
					}
					$idt.Phrases.SetCurrentLine($idt.Phrases.Count - 1);
         
					#Adding Phrases
					foreach ($phrase in $phrases) {
						$idt.Phrases.U_PhCode = $phrase.PhraseCode;
						if ($phrase.ProductionOrders -eq 1) {
							$idt.Phrases.U_ProdOrders = "Y"
						}
						if ($phrase.ShipmentDocuments -eq 1) {
							$idt.Phrases.U_ShipDoc = "Y"
						}	
						if ( $phrase.PickLists -eq 1) {
							$idt.Phrases.U_PickLists = "Y"
						}
						if ($phrase.MSDS -eq 1) {
							$idt.Phrases.U_MSDS = "Y"
						}	
						if ( $phrase.PurchaseOrders -eq 1) {
							$idt.Phrases.U_PurOrders = "Y"
						}	
						if ($phrase.Returns -eq 1) {
							$idt.Phrases.U_Returns = "Y"
						}	
						if ($phrase.Other -eq 1) {
							$idt.Phrases.U_Other = "Y"
						}
						$dummy = $idt.Phrases.Add()
					}
				} 
			}
			Catch {
				$err = $_.Exception.Message;
				$ms = [string]::Format("Error when updating Phrases for Item Details with ItemCode {0} Details: {1}", $key, $err);
				Throw [System.Exception]($ms);
			}
			#endregion

			#region Properties
			try {
				$PropertiesLineNumDict = New-Object 'System.Collections.Generic.Dictionary[string,int]';
				[array] $properties = $dictItemDetailsProperties[$key];
				if ($properties.count -gt 0) {
					#Deleting all exisitng Properties
					$count = $idt.Properties.Count
					for ($i = 0; $i -lt $count; $i++) {
						$dummy = $idt.Properties.DelRowAtPos(0);
					}
					$idt.Properties.SetCurrentLine($idt.Properties.Count - 1);
         
					#Adding Properies
					foreach ($prop in $properties) {
						$idt.Properties.U_PrpCode = $prop.PropertyCode;
                        
						switch ($prop.Expression) {
							'BT' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::Between; break; }
							'EQ' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::Equal; break; }
							'NE' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::NotEqual; break; }
							'GT' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::GratherThan; break; }
							'GE' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::GratherThanOrEqual; break; }
							'LE' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::LessThanOrEqual; break; }
							'LT' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::LessThan; break; }
							Default { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::Equal; }
						} 
						$idt.Properties.U_Expression = $EnumExpressionValue;

						if ($prop.RangeFrom -ne "") {
							$idt.Properties.U_RangeValueFrom = $prop.RangeFrom;
						}
						else {
							$idt.Properties.U_RangeValueFrom = 0;
						}
						$idt.Properties.U_RangeValueTo = $prop.RangeTo;
						if ($prop.ReferenceCode -ne "") {
							$idt.Properties.U_WordCode = $prop.ReferenceCode;
						}
						$idt.Properties.U_Remarks = $prop.Remarks
						$PropertiesLineNumDict.Add($prop.PropertyCode, $idt.Properties.U_LineNum);
						$dummy = $idt.Properties.Add()
					}
				} 
			}
			Catch {
				$err = $_.Exception.Message;
				$ms = [string]::Format("Error when updating Properties for Item Details with ItemCode {0} Details: {1}", $key, $err);
				Throw [System.Exception]($ms);
			}
			#endregion
			
			#region Properties Certificates of Analysis
			try {
				[array] $propertiesCerts = $dictItemDetailsPropertiesCerts[$key];
				if ($propertiesCerts.count -gt 0) {
					$BusinessPartnerRelations = $idt.BusinessPartnerRelations;
					#Deleting all exisitng lines
					$count = $BusinessPartnerRelations.Count
					for ($i = 0; $i -lt $count; $i++) {
						$BusinessPartnerRelations.SetCurrentLine(0);
						if ($BusinessPartnerRelations.IsRowFilled()) {
							$dummy = $BusinessPartnerRelations.DelRowAtPos(0);
						}
					}
					
					$BusinessPartnerRelations.SetCurrentLine(0);
			
					#Adding Certificates
					foreach ($crt in $propertiesCerts) {
						if ($PropertiesLineNumDict.ContainsKey($crt.PropertyCode) -eq $false) {
							$err = [string]::Format("Property with Code:{0} don't exists.", $crt.PropertyCode);
							throw [System.Exception]($err)
						}
						$BusinessPartnerRelations.U_BaseLineNum = $PropertiesLineNumDict[$crt.PropertyCode]; 
						$BusinessPartnerRelations.U_CardCode = $crt.CardCode;
						switch ($crt.Expression) {
							'BT' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::Between; break; }
							'EQ' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::Equal; break; }
							'NE' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::NotEqual; break; }
							'GT' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::GratherThan; break; }
							'GE' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::GratherThanOrEqual; break; }
							'LE' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::LessThanOrEqual; break; }
							'LT' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::LessThan; break; }
							Default { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::Equal; }
						} 
						$BusinessPartnerRelations.U_Expression = $EnumExpressionValue;

						if ([string]::IsNullOrWhiteSpace($crt.ValueFrom) -eq $false ) {
							$BusinessPartnerRelations.U_ValueFrom = $crt.ValueFrom;
						}
						else {
							$BusinessPartnerRelations.U_ValueFrom = 0;
						}
						$BusinessPartnerRelations.U_ValueTo = $crt.ValueTo;

						if ([string]::IsNullOrWhiteSpace($crt.ValidFrom) -eq $false) {
							$BusinessPartnerRelations.U_FromDate = $crt.ValidFrom
						}
						if ([string]::IsNullOrWhiteSpace($crt.ValidTo) -eq $false) {
							$BusinessPartnerRelations.U_ToDate = $crt.ValidTo
						}
						$dummy = $BusinessPartnerRelations.Add();
					}
				} 
			}
			Catch {
				$err = $_.Exception.Message;
				$ms = [string]::Format("Error when updating Properties Certificates of Analysis for Item Details with ItemCode {0} Details: {1}", $key, $err);
				Throw [System.Exception]($ms);
			}
			#endregion

			#region Texts
			try {
				[array] $texts = $dictItemDetailsTexts[$key];
				if ($texts.count -gt 0) {
					#Deleting all exisitng Texts
					$count = $idt.Texts.Count
					for ($i = 0; $i -lt $count; $i++) {
						$dummy = $idt.Texts.DelRowAtPos(0);
					}
					$idt.Texts.SetCurrentLine($idt.Texts.Count - 1);
         
					#Adding Texts
					foreach ($Text in $texts) {
						$idt.Texts.U_TxtCode = $Text.TextCode;
						if ($Text.ProductionOrders -eq 1) {
							$idt.Texts.U_ProdOrders = "Y"
						}
						if ($Text.ShipmentDocuments -eq 1) {
							$idt.Texts.U_ShipDoc = "Y"
						}	
						if ( $Text.PickLists -eq 1) {
							$idt.Texts.U_PickLists = "Y"
						}
						if ($Text.MSDS -eq 1) {
							$idt.Texts.U_MSDS = "Y"
						}	
						if ( $Text.PurchaseOrders -eq 1) {
							$idt.Texts.U_PurOrders = "Y"
						}	
						if ($Text.Returns -eq 1) {
							$idt.Texts.U_Returns = "Y"
						}	
						if ($Text.Other -eq 1) {
							$idt.Texts.U_Other = "Y"
						}
						$dummy = $idt.Texts.Add()
					}
				} 
			}
			Catch {
				$err = $_.Exception.Message;
				$ms = [string]::Format("Error when updating Texts for Item Details with ItemCode {0} Details: {1}", $key, $err);
				Throw [System.Exception]($ms);
			}
			#endregion

			$message = $idt.Update()
        
			if ($message -lt 0) {  
				$err = $pfcCompany.GetLastErrorDescription()
				Throw [System.Exception] ($err);
			}
			else {
				$atLeastOneSuccess = $true;
			}
        
        
		}
		Catch {
			$err = $_.Exception.Message;
			$ms = [string]::Format("Error when updating Item Details for ItemCode {0} Details: {1}", $key, $err);
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
	if ($DIRECT_ACCESS_MODE -eq $true) {
		[CompuTec.ProcessForce.API.DynamicCosting.Restoration.ItemCostingRestorationManager]::SuspendItemCostingRestoration = $false
		if ($atLeastOneSuccess -eq $true) { 
			Write-Host '';
			Write-Host 'Restoring cosing data'
			$manager = New-Object  "CompuTec.ProcessForce.API.DynamicCosting.Restoration.ItemCostingRestorationManager" $pfcCompany.Token
			$manager.Initialize();
			$manager.Restore();
		}
	}

	#region Close connection
    
	if ($pfcCompany.IsConnected) {
		$pfcCompany.Disconnect()
		write-host  –backgroundcolor green –foregroundcolor black "Disconnected from the company"
	}
    
	#endregion

}


