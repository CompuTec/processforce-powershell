#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Quality Control Tests
########################################################################
$SCRIPT_VERSION = "3.2"
# Last tested PF version: ProcessForce 9.3 (9.30.210) PL: MAIN (64-bit)
# Description:y
#      Import Quality Control Tests. Script add new Quality Control Tests.
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

$csvQualityControlTestsPath = -join ($csvImportCatalog, "QualityControlTests.csv")
$csvQualityControlTestPropertiesPath = -join ($csvImportCatalog, "QualityControlTestProperties.csv")
$csvQualityControlTestPropertiesCertsPath = -join ($csvImportCatalog, "QualityControlTestPropertiesCertifiacteOfAnalysis.csv")
$csvQualityControlTestPropertiesItemPath = -join ($csvImportCatalog, "QualityControlTestPropertiesItem.csv")
$csvQualityControlTestPropertiesItemCertsPath = -join ($csvImportCatalog, "QualityControlTestPropertiesItemCertifiacteOfAnalysis.csv")
$csvQualityControlTestItemsPath = -join ($csvImportCatalog, "QualityControlTestItems.csv")
$csvQualityControlTestDefectsPath = -join ($csvImportCatalog, "QualityControlTestDefects.csv")
$csvQualityControlTestResourcesPath = -join ($csvImportCatalog, "QualityControlTestResources.csv")
$csvQualityControlTestTransBatchesPath = -join ($csvImportCatalog, "QualityControlTestTransBatches.csv")
$csvQualityControlTestTransSerialNumbersPath = -join ($csvImportCatalog, "QualityControlTestTransSerialNumbers.csv")

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

	[array] $csvQualityControlTests = Import-Csv -Delimiter ';' -Path $csvQualityControlTestsPath 			
	
	[array] $csvQualityControlTestProperties = $null;
	if ((Test-Path -Path $csvQualityControlTestPropertiesPath -PathType leaf) -eq $true) {
		[array] $csvQualityControlTestProperties = Import-Csv -Delimiter ';' -Path $csvQualityControlTestPropertiesPath
	}
	else { 
		write-host "Properties - csv not available." 
	}
	[array] $csvQualityControlTestPropertiesCerts = $null;
	if ((Test-Path -Path $csvQualityControlTestPropertiesCertsPath -PathType leaf) -eq $true) {
		[array] $csvQualityControlTestPropertiesCerts = Import-Csv -Delimiter ';' -Path $csvQualityControlTestPropertiesCertsPath
	}
	else { 
		write-host "Properties Certificate Of Analiysis - csv not available." 
	}
	[array] $csvQualityControlTestPropertiesItem = $null;
	if ((Test-Path -Path $csvQualityControlTestPropertiesItemPath -PathType leaf) -eq $true) { 
		[array] $csvQualityControlTestPropertiesItem = Import-Csv -Delimiter ';' -Path $csvQualityControlTestPropertiesItemPath 
	}
	else { 
		write-host "Item Properties - csv not available." 
	}
	[array] $csvQualityControlTestPropertiesItemCerts = $null;
	if ((Test-Path -Path $csvQualityControlTestPropertiesItemCertsPath -PathType leaf) -eq $true) { 
		[array] $csvQualityControlTestPropertiesItemCerts = Import-Csv -Delimiter ';' -Path $csvQualityControlTestPropertiesItemCertsPath 
	}
	else { 
		write-host "Item Properties Certificate Of Analysis- csv not available." 
	}
	[array] $csvQualityControlTestItems = $null;
	if ((Test-Path -Path $csvQualityControlTestItemsPath -PathType leaf) -eq $true) {
		[array] $csvQualityControlTestItems = Import-Csv -Delimiter ';' -Path $csvQualityControlTestItemsPath	
	}
	else { 
		write-host "Items - csv not available." 
	}
	[array] $csvQualityControlTestDefects = $null;
	if ((Test-Path -Path $csvQualityControlTestDefectsPath -PathType leaf) -eq $true) { 
		[array] $csvQualityControlTestDefects = Import-Csv -Delimiter ';' -Path $csvQualityControlTestDefectsPath 
	} 
	else { write-host "Defects - csv not available." }
	[array] $csvQualityControlTestResources = $null;
	if ((Test-Path -Path $csvQualityControlTestResourcesPath -PathType leaf) -eq $true) { 
		[array] $csvQualityControlTestResources = Import-Csv -Delimiter ';' -Path $csvQualityControlTestResourcesPath 
	} 
	else { 
		write-host "Resources - csv not available." 
	}
	[array] $csvQualityControlTestTransBatches = $null;
	if ((Test-Path -Path $csvQualityControlTestTransBatchesPath -PathType leaf) -eq $true) {
		[array] $csvQualityControlTestTransBatches = Import-Csv -Delimiter ';' -Path $csvQualityControlTestTransBatchesPath 
	}
	else { 
		write-host "Batches - csv not available." 
	}
	[array] $csvQualityControlTestTransSerialNumbers = $null
	if ((Test-Path -Path $csvQualityControlTestTransSerialNumbersPath -PathType leaf) -eq $true) {
		[array] $csvQualityControlTestTransSerialNumbers = Import-Csv -Delimiter ';' -Path $csvQualityControlTestTransSerialNumbersPath 
	}
	else { 
		write-host "Serial Numbers - csv not available." 
	}

	write-Host 'Preparing data: '
	$totalRows = $csvQualityControlTests.Count + $csvQualityControlTestProperties.Count + $csvQualityControlTestPropertiesItem.Count +
	$csvQualityControlTestItems.Count + $csvQualityControlTestDefects.Count + $csvQualityControlTestResources.Count + $csvQualityControlTestTransBatches.Count +
	$csvQualityControlTestTransSerialNumbers.Count + $csvQualityControlTestPropertiesCerts.Count + $csvQualityControlTestPropertiesItemCerts.Count;
	
	$qcTestsList = New-Object 'System.Collections.Generic.List[array]'
	$qcTestPropertiesDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$qcTestPropertiesCertsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$qcTestPropertiesItemDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$qcTestPropertiesItemCertsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$qcTestItemsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$qcTestDefectsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$qcTestResourcesDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$qcTestTransBatchesDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$qcTestTransSerialNumbersDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}

	foreach ($row in $csvQualityControlTests) {
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		$qcTestsList.Add([array]$row);
	}

	foreach ($row in $csvQualityControlTestProperties) {
		$key = $row.Key;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($qcTestPropertiesDict.ContainsKey($key) -eq $false) {
			$qcTestPropertiesDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $qcTestPropertiesDict[$key];
		
		$list.Add([array]$row);
	}

	foreach ($row in $csvQualityControlTestPropertiesCerts) {
		$key = $row.Key;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($qcTestPropertiesCertsDict.ContainsKey($key) -eq $false) {
			$qcTestPropertiesCertsDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $qcTestPropertiesCertsDict[$key];
		
		$list.Add([array]$row);
	}

	foreach ($row in $csvQualityControlTestPropertiesItem) {
		$key = $row.Key;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($qcTestPropertiesItemDict.ContainsKey($key) -eq $false) {
			$qcTestPropertiesItemDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $qcTestPropertiesItemDict[$key];
		
		$list.Add([array]$row);
	}

	foreach ($row in $csvQualityControlTestPropertiesItemCerts) {
		$key = $row.Key;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($qcTestPropertiesItemCertsDict.ContainsKey($key) -eq $false) {
			$qcTestPropertiesItemCertsDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $qcTestPropertiesItemCertsDict[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvQualityControlTestItems) {
		$key = $row.Key;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($qcTestItemsDict.ContainsKey($key) -eq $false) {
			$qcTestItemsDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $qcTestItemsDict[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvQualityControlTestDefects) {
		$key = $row.Key;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($qcTestDefectsDict.ContainsKey($key) -eq $false) {
			$qcTestDefectsDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $qcTestDefectsDict[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvQualityControlTestResources) {
		$key = $row.Key;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($qcTestResourcesDict.ContainsKey($key) -eq $false) {
			$qcTestResourcesDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $qcTestResourcesDict[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvQualityControlTestTransBatches) {
		$key = $row.Key;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($qcTestTransBatchesDict.ContainsKey($key) -eq $false) {
			$qcTestTransBatchesDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $qcTestTransBatchesDict[$key];
		
		$list.Add([array]$row);
	}
	foreach ($row in $csvQualityControlTestTransSerialNumbers) {
		$key = $row.Key;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($qcTestTransSerialNumbersDict.ContainsKey($key) -eq $false) {
			$qcTestTransSerialNumbersDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $qcTestTransSerialNumbersDict[$key];
		
		$list.Add([array]$row);
	}
	
	Write-Host '';
	Write-Host 'Adding/updating data: ' -NoNewLine;

	if ($qcTestsList.Count -gt 1) {
		$total = $qcTestsList.Count
	}
	else {
		$total = 1
	}
	$progressItterator = 0;
	$progress = 0;
	$beforeProgress = 0;
    
	foreach ($csvTest in $qcTestsList) {
		try {
			$progressItterator++;
			$progress = [math]::Round(($progressItterator * 100) / $total);
			if ($progress -gt $beforeProgress) {
				Write-Host $progress"% " -NoNewline
				$beforeProgress = $progress
			}
			#Creating ControlTest
			$test = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"QualityControlTest")
            
			$test.U_TestProtocolNo = $csvTest.TestProtocolCode;
			$test.U_ItemCode = $csvTest.ItemCode;
			$test.U_RevCode = $csvTest.RevisionCode;
			$test.U_WhsCode = $csvTest.Warehouse;
			$test.U_ComplaintNo = $csvTest.ComplaintNo;
			$test.U_PrjCode = $csvTest.Project;
			$test.U_InsCode = $csvTest.InspectorCode;
			$test.U_ElectronicSign = $csvTest.ElectronicSign;
			$test.U_Status = [string] $csvTest.Status;
			if ($csvTest.CreatedDate -ne "") {
				$test.U_Created = $csvTest.CreatedDate;
			}
			else {
				$test.U_Created = [datetime]::MinValue
			}
			if ($csvTest.StartDate -ne "") {
				$test.U_Start = $csvTest.StartDate;
			}
			else {
				$test.U_Start = [datetime]::MinValue
			}
			if ( $csvTest.OnHoldDate -ne "") {
				$test.U_OnHold = $csvTest.OnHoldDate;
			}
			else {
				$test.U_OnHold = [datetime]::MinValue
			}
			if ( $csvTest.WaitingNcmrDate -ne "") {
				$test.U_WaitingNcmr = $csvTest.WaitingNcmrDate;
			}
			else {
				$test.U_WaitingNcmr = [datetime]::MinValue;
			}
			if ( $csvTest.ClosedDate -ne "") {
				$test.U_Closed = $csvTest.ClosedDate;
			}
			else {
				$test.U_Closed = [datetime]::MinValue
			}
	
			$test.U_TestStatus = $csvTest.TestStatus;
			if ($csvTest.Pass_FailDate -ne "") {
				$test.U_PassFailDate = $csvTest.Pass_FailDate;
			}
			else {
				$test.U_PassFailDate = [datetime]::MinValue
			}
			#Defects
			$test.U_SampleSize = $csvTest.DefSampleSize;
			$test.U_UoM = $csvTest.DefUoM;
			$test.U_PassedQty = $csvTest.DefPassedQty;
			$test.U_DefectQty = $csvTest.DefectQty;
			$test.U_InvMove = $csvTest.InventoryMovements;
			$test.U_Ncmr = $csvTest.NCMR;
			$test.U_NcmrInsCode = $csvTest.NcmrInspectorCode;
			$test.U_Remarks = $csvTest.DefRemarks;
			#Transactions
			$test.U_TransType = $csvTest.TransactionType;
			$test.U_BpCode = $csvTest.BPCode;
			$test.U_MnfOprCode = $csvTest.OperationCode;
	
			#Properties
			$TestPropertiesLineNumDict = New-Object 'System.Collections.Generic.Dictionary[string,int]';
			[array]$Properties = $qcTestPropertiesDict[$csvTest.Key] 
			if ($Properties.count -gt 0) {
				#Deleting all exisitng Properties
				$count = $test.TestResults.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $test.TestResults.DelRowAtPos(0);
				}
				$test.TestResults.SetCurrentLine(0);
				#Adding Properties
				foreach ($prop in $Properties) {
					$test.TestResults.U_PrpCode = $prop.PropertyCode;

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
					$test.TestResults.U_Expression = $EnumExpressionValue;
                        
					if ($prop.RangeFrom -ne "") {
						$test.TestResults.U_RangeValueFrom = $prop.RangeFrom;
					}
					else {
						$test.TestResults.U_RangeValueFrom = 0;
					}
					$test.TestResults.U_RangeValueTo = $prop.RangeTo;
				
					if ($prop.UoM -ne "") {
						$test.TestResults.U_UnitOfMeasure = $prop.UoM;
					}
				
					$test.TestResults.U_TestedValue = $prop.TestedValue;
				
					if ($prop.ReferenceCode -ne "") {
						$test.TestResults.U_RefCode = $prop.ReferenceCode;
					}
					$test.TestResults.U_TestedRefCode = $prop.TestedRefCode;
					$test.TestResults.U_PassFail = $prop.Pass_Fail;
					$test.TestResults.U_RsnCode = $prop.ReasonCode;
					$test.TestResults.U_Remarks = $prop.Remarks;
					$TestPropertiesLineNumDict.Add($prop.PropertyCode, $test.TestResults.U_LineNum);
					$dummy = $test.TestResults.Add()
				}
			}
	
			#Properties Certificates of Analysis
			[array]$PropertiesCerts = $qcTestPropertiesCertsDict[$csvTest.Key] 
			if ($PropertiesCerts.count -gt 0) {
				$BusinessPartnerRelations = $test.TestResultBusinessPartnerRelations;
				#Deleting all exisitng lines
				$count = $BusinessPartnerRelations.Count;
				for ($i = 0; $i -lt $count; $i++) {
					$BusinessPartnerRelations.SetCurrentLine(0);
					if ($BusinessPartnerRelations.IsRowFilled()) {
						$dummy = $BusinessPartnerRelations.DelRowAtPos(0);
					}
				}
				$BusinessPartnerRelations.SetCurrentLine(0);
         
				#Adding Certificates
				foreach ($crt in $PropertiesCerts) {
					if ($TestPropertiesLineNumDict.ContainsKey($crt.PropertyCode) -eq $false) {
						$err = [string]::Format("Test Property with Code:{0} don't exists.", $crt.PropertyCode);
						throw [System.Exception]($err)
					}
					$BusinessPartnerRelations.U_BaseLineNum = $TestPropertiesLineNumDict[$crt.PropertyCode]; 
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
			#ItemProperties
			$ItemPropertiesLineNumDict = New-Object 'System.Collections.Generic.Dictionary[string,int]';
			[array]$ItemProperties = $qcTestPropertiesItemDict[$csvTest.Key];
			if ($ItemProperties.count -gt 0) {
				#Deleting all exisitng ItemProperties
				$count = $test.ItemProperties.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $test.ItemProperties.DelRowAtPos(0);
				}
				$test.ItemProperties.SetCurrentLine($test.ItemProperties.Count - 1);
	         
				#Adding Item Properies
				foreach ($itprop in $properties) {
					$test.ItemProperties.U_PrpCode = $itprop.PropertyCode;

					switch ($itprop.Expression) {
						'BT' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::Between; break; }
						'EQ' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::Equal; break; }
						'NE' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::NotEqual; break; }
						'GT' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::GratherThan; break; }
						'GE' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::GratherThanOrEqual; break; }
						'LE' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::LessThanOrEqual; break; }
						'LT' { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::LessThan; break; }
						Default { $EnumExpressionValue = [CompuTec.ProcessForce.API.Enumerators.ConditionType]::Equal; }
					} 
					$test.ItemProperties.U_Expression = $EnumExpressionValue;

					if ($itprop.RangeFrom -ne "") {
						$test.ItemProperties.U_RangeValueFrom = $itprop.RangeFrom;
					}
					else {
						$test.ItemProperties.U_RangeValueFrom = 0;
					}
					$test.ItemProperties.U_RangeValueTo = $itprop.RangeTo;
				
					$test.ItemProperties.U_TestedValue = $itprop.TestedValue;
				
					if ($itprop.ReferenceCode -ne "") {
						$test.ItemProperties.U_RefCode = $itprop.ReferenceCode;
					}
					$test.ItemProperties.U_TestedRefCode = $itprop.TestedRefCode;
					$test.ItemProperties.U_PassFail = $itprop.Pass_Fail;
					$test.ItemProperties.U_RsnCode = $itprop.ReasonCode;
					$test.ItemProperties.U_Remarks = $itprop.Remarks;
					$ItemPropertiesLineNumDict.Add($itprop.PropertyCode, $test.ItemProperties.U_LineNum); 
					$dummy = $test.ItemProperties.Add()
				}
			}
			#ItemProperties Certificates of Analysis
			[array]$ItemPropertiesCerts = $qcTestPropertiesItemCertsDict[$csvTest.Key];
			if ($ItemPropertiesCerts.count -gt 0) {
				$BusinessPartnerRelations = $test.ItemPropertyBusinessPartnerRelations;
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
				foreach ($crt in $ItemPropertiesCerts) {
					if ($ItemPropertiesLineNumDict.ContainsKey($crt.PropertyCode) -eq $false) {
						$err = [string]::Format("Test Property with Code:{0} don't exists.", $crt.PropertyCode);
						throw [System.Exception]($err)
					}
					$BusinessPartnerRelations.U_BaseLineNum = $ItemPropertiesLineNumDict[$crt.PropertyCode]; 
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
			#Resources
			[array]$Resources = $qcTestResourcesDict[$csvTest.Key];
			if ($Resources.count -gt 0) {
				#Deleting all exisitng Resources
				$count = $test.Resources.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $test.Resources.DelRowAtPos(0);
				}
				$test.Resources.SetCurrentLine($test.Resources.Count - 1);
	         
				#Adding Resources
				foreach ($resource in $Resources) {
					$test.Resources.U_RscCode = $resource.ResourceCode;
					$test.Resources.U_WhsCode = $resource.Warehouse;
					$test.Resources.U_PlannedQty = $resource.PlanedQuantity;
					$test.Resources.U_ActualQty = $resource.ActualQuantity;
					$test.Resources.U_Remarks = $resource.Remarks;
					$dummy = $test.Resources.Add()
				}
			}
  	
	
			#Items
			[array]$Items = $qcTestItemsDict[$csvTest.Key];
			if ($Items.count -gt 0) {
				#Deleting all exisitng Items
				$count = $test.Items.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $test.Items.DelRowAtPos(0);
				}
				$test.Items.SetCurrentLine($test.Items.Count - 1);
	         
				#Adding Items
				foreach ($Item in $Items) {
					$test.Items.U_ItemCode = $Item.ItemCode;
					$test.Items.U_WhsCode = $Item.Warehouse;
					$test.Items.U_PlannedQty = $Item.PlanedQuantity;
					$test.Items.U_ActualQty = $Item.ActualQuantity;
					$test.Items.U_Remarks = $Item.Remarks;
					$dummy = $test.Items.Add()
				}
			}
	
			#Defects
			[array]$defects = $qcTestDefectsDict[$csvTest.Key];
			if ($defects.count -gt 0) {
				#Deleting all exisitng Items
				$count = $test.Defects.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $test.Defects.DelRowAtPos(0);
				}
				$test.Defects.SetCurrentLine($test.Defects.Count - 1);
	         
				#Adding Defects
				foreach ($defect in $defects) {
					$test.Defects.U_DefCode = $defect.DefectCode;
					$dummy = $test.Defects.Add()
				}
			}
	
			#Batches
			[array]$batches = $qcTestTransBatchesDict[$csvTest.Key];
			if ($batches.count -gt 0) {
				#Deleting all exisitng Items
				$count = $test.Batches.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $test.Batches.DelRowAtPos(0);
				}
				$test.Batches.SetCurrentLine($test.Batches.Count - 1);
	         
				#Adding Defects
				foreach ($batch in $batches) {
					$test.Batches.U_Batch = $batch.Batch;
					$dummy = $test.Batches.Add()
				}
			}
  	
	
			#SerialNumbers
			[array]$SerilaNumbers = $qcTestTransSerialNumbersDict[$csvTest.Key];
			if ($SerilaNumbers.count -gt 0) {
				#Deleting all exisitng Items
				$count = $test.SerialNumbers.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $test.SerialNumbers.DelRowAtPos(0);
				}
				$test.SerialNumbers.SetCurrentLine($test.SerialNumbers.Count - 1);
	         
				#Adding Defects
				foreach ($sn in $SerilaNumbers) {
					$test.SerialNumbers.U_SerialNo = $sn.SerialNumber;
					$dummy = $test.SerialNumbers.Add()
				}
			}
	
			$message = 0
    
			#Adding or updating Test
			$message = $test.Add()
         
			if ($message -lt 0) {    
				$err = $pfcCompany.GetLastErrorDescription()
				Throw [System.Exception] ($err)
			}
		}
		Catch {
			$err = $_.Exception.Message;
			$taskMsg = "adding";

			$ms = [string]::Format("Error when {0} Quality Test Control with Key {1} Details: {2}", $taskMsg, $csvTest.Key, $err);
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