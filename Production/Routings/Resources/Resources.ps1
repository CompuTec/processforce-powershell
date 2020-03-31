#region Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Resources
########################################################################
$SCRIPT_VERSION = "3.1"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Resources. Script add new or will update existing Resources.
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

$csvResourcesPath = -join ($csvImportCatalog, "Resources.csv")
$csvResourcesPropertiesPath = -join ($csvImportCatalog, "ResourcesProperties.csv")
$csvResourcesAtachmentsPath = -join ($csvImportCatalog, "ResourcesAttachments.csv")
$csvResourcesPlanningInfoPath = -join ($csvImportCatalog, "ResourcesPlanningInfo.csv")

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
	[array] $csvResources = Import-Csv -Delimiter ';' -Path  $csvResourcesPath

	if ((Test-Path -Path $csvResourcesPropertiesPath -PathType leaf) -eq $true) {
		[array] $csvResourcesProperties = Import-Csv -Delimiter ';' -Path  $csvResourcesPropertiesPath
	}
	else {
		[array] $csvResourcesProperties = $null;
		write-host "Resources Properties - csv not available."
	}

	if ((Test-Path -Path $csvResourcesAtachmentsPath -PathType leaf) -eq $true) {
		[array] $csvResourcesAtachments = Import-Csv -Delimiter ';' -Path  $csvResourcesAtachmentsPath
	}
	else {
		[array] $csvResourcesAtachments = $null;
		write-host "Resources Attachment - csv not available."
	}
	
	if ((Test-Path -Path $csvResourcesPlanningInfoPath -PathType leaf) -eq $true) {
		[array] $csvResourcesPlanningInfo = Import-Csv -Delimiter ';' -Path  $csvResourcesPlanningInfoPath
	}
	else {
		[array] $csvResourcesPlanningInfo = $null;
		write-host "Resources Planning Info - csv not available."
	}
 
	#region preparing data
	write-Host 'Preparing data: ' -NoNewline
	$totalRows = $csvResources.Count + $csvResourcesProperties.Count + $csvResourcesAtachments.Count + $csvResourcesPlanningInfoPath.Count;

	$resourcesList = New-Object 'System.Collections.Generic.List[array]'
	$resourcesPropertiesDict = New-Object 'System.Collections.Generic.Dictionary[String,System.Collections.Generic.List[array]]'
	$resourcesAtachementsDict = New-Object 'System.Collections.Generic.Dictionary[String,System.Collections.Generic.List[array]]'
	$resourcesPlanningInfoDict = New-Object 'System.Collections.Generic.Dictionary[String,System.Collections.Generic.List[array]]'
    
	$progressItterator = 0;
	$progress = 0;
	$beforeProgress = 0;

	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}

	foreach ($row in $csvResources) {
		$progressItterator++;
		$progress = [math]::Round(($progressItterator * 100) / $total);
		if ($progress -gt $beforeProgress) {
			Write-Host $progress"% " -NoNewline
			$beforeProgress = $progress
		}

		$resourcesList.Add([array]$row);
	}

	foreach ($row in $csvResourcesProperties) {
		$key = $row.ResourceCode;
		$progressItterator++;
		$progress = [math]::Round(($progressItterator * 100) / $total);
		if ($progress -gt $beforeProgress) {
			Write-Host $progress"% " -NoNewline
			$beforeProgress = $progress
		}

		if ($resourcesPropertiesDict.ContainsKey($key)) {
			$list = $resourcesPropertiesDict[$key];
		}
		else {
			$list = New-Object System.Collections.Generic.List[array];
			$resourcesPropertiesDict[$key] = $list;
		}

		$list.Add([array]$row);
	}

	foreach ($row in $csvResourcesAtachments) {
		$key = $row.ResourceCode;
		$progressItterator++;
		$progress = [math]::Round(($progressItterator * 100) / $total);
		if ($progress -gt $beforeProgress) {
			Write-Host $progress"% " -NoNewline
			$beforeProgress = $progress
		}

		if ($resourcesAtachementsDict.ContainsKey($key)) {
			$list = $resourcesAtachementsDict[$key];
		}
		else {
			$list = New-Object System.Collections.Generic.List[array];
			$resourcesAtachementsDict[$key] = $list;
		}

		$list.Add([array]$row);
	}

	foreach ($row in $csvResourcesPlanningInfo) {
		$key = $row.ResourceCode;
		$progressItterator++;
		$progress = [math]::Round(($progressItterator * 100) / $total);
		if ($progress -gt $beforeProgress) {
			Write-Host $progress"% " -NoNewline
			$beforeProgress = $progress
		}

		if ($resourcesPlanningInfoDict.ContainsKey($key)) {
			$list = $resourcesPlanningInfoDict[$key];
		}
		else {
			$list = New-Object System.Collections.Generic.List[array];
			$resourcesPlanningInfoDict[$key] = $list;
		}

		$list.Add([array]$row);
	}
	Write-Host '';
	Write-Host 'Adding/updating data: ' -NoNewline;
	#endregion

	$progressItterator = 0;
	$progress = 0;
	$beforeProgress = 0;
	$totalRows = $resourcesList.Count;
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}
	foreach ($csvItem in $resourcesList) {
		try {
			$progressItterator++;
			$progress = [math]::Round(($progressItterator * 100) / $total);
			if ($progress -gt $beforeProgress) {
				Write-Host $progress"% " -NoNewline
				$beforeProgress = $progress
			}
			#Creating Resource object
			$res = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"Resource")
			#Checking that the resource already exist
			$retVal = $res.GetByRscCode($csvItem.ResourceCode)
			if ($retVal -ne 0) {
				#Adding the new data
				$res.U_RscType = $csvItem.ResourceType #enum type; Machine = 1 or M, Labour = 2 or L, Tool = 3 or T, Subcontractor = 4 or S 
				$res.U_RscCode = $csvItem.ResourceCode
				$exists = $false;
			}
			else {
				$exists = $true;
			}
			$res.U_RscName = $csvItem.ResourceName
			$res.U_RscGrpCode = $csvItem.ResourceGroup
			$res.U_QueueTime = $csvItem.QueTime
			if ($csvItem.QueTimeUoM -ne '') {
				$res.U_QueueRate = $csvItem.QueTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9 
			}
			$res.U_SetupTime = $csvItem.SetupTime
			if ($csvItem.SetupTimeUoM -ne '') {
				$res.U_SetupRate = $csvItem.SetupTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9 
			}
			$res.U_RunTime = $csvItem.RunTime
			if ($csvItem.RunTimeUoM -ne '') {
				$res.U_RunRate = $csvItem.RunTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9 
			}
			$res.U_StockTime = $csvItem.StockTime
			if ($csvItem.StockTimeUoM -ne '') {
				$res.U_StockRate = $csvItem.StockTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9     
			}
			if ($csvItem.ResourceNumber -ne '') {
				$res.U_ResourceCount = $csvItem.ResourceNumber
			}
			if ($csvItem.HasCycle -eq 1) {
				$res.U_HasCycles = $csvItem.HasCycle #enum type; 1 = Yes, 2 = No
				$res.U_CycleCap = $csvItem.CycleCapacity
			}
        
			$res.U_ResActCode = $csvItem.ResourceAccountingCode
			$res.U_Project = $csvItem.Project
			$res.U_OcrCode = $csvItem.Dimension1
			$res.U_OcrCode2 = $csvItem.Dimension2
			$res.U_OcrCode3 = $csvItem.Dimension3
			$res.U_OcrCode4 = $csvItem.Dimension4
			$res.U_OcrCode5 = $csvItem.Dimension5
			$res.U_WhsCode = $csvItem.IssueWhsCode
			$res.U_BinAbs = $csvItem.IssueBinAbs
			$res.U_RWhsCode = $csvItem.ReceiptWhsCode
			$res.U_RBinAbs = $csvItem.ReceiptBinAbs

    
			#$res.UDFItems.Item("U_UDF1").Value = $csvItem.UDF1 ## how to import UDFs
	
			if ($res.U_RscType -eq 'Subcontractor') {
				$res.U_VendorCode = $csvItem.VendorCode
				$res.U_ItemCode = $csvItem.ItemCode
			}
   			switch ($csvItem.PlanningWarningTable) {
				"CT_PF_OMOR" {
					$res.U_WarningTable = [CompuTec.ProcessForce.API.Documents.Resource.ResourcePlanningInformationTableType]::ManufacturingOrder; 
					break;
				}
				"CT_PF_MOR12" {
					$res.U_WarningTable = [CompuTec.ProcessForce.API.Documents.Resource.ResourcePlanningInformationTableType]::Operation; 
					break;
				}
				"CT_PF_MOR16" {
					$res.U_WarningTable = [CompuTec.ProcessForce.API.Documents.Resource.ResourcePlanningInformationTableType]::Resource; 
					break;
				}
				"DYNAMIC" {
					$res.U_WarningTable = [CompuTec.ProcessForce.API.Documents.Resource.ResourcePlanningInformationTableType]::Dynamic; 
					break;
				}
				Default {
					$res.U_WarningTable = [CompuTec.ProcessForce.API.Documents.Resource.ResourcePlanningInformationTableType]::Undefined; 
				 }
			}

			   
			$res.U_WarningField = $csvItem.PlanningWarningField;

			if ([string]::Equals($csvItem.PlanningWarningSQLEnabled, "Y", [System.StringComparison]::InvariantCultureIgnoreCase)) {
				$res.U_IsWarnSqlEnabled = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
				$res.U_WarningSqlQuery = $csv.PlanningWarningSqlQuery;
			}
			else {
				$res.U_IsWarnSqlEnabled = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
			}
			#Data loading from a csv file - Resource Properties
			[array]$resProps = $resourcesPropertiesDict[$csvItem.ResourceCode]
			if ($resProps.count -gt 0) {
				#Deleting all existing properties
				$count = $res.Properties.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $res.Properties.DelRowAtPos(0);
				}
        
				#Adding the new data       
				foreach ($prop in $resProps) {
					$res.Properties.U_PrpCode = $prop.PropertiesCode
					$res.Properties.U_PrpConType = $prop.Condition #enum ConditionType; Equal EQ = 1, NotEqual NE = 2, GratherThan GT = 3, GratherThanOrEqual GE = 4, LessThan LT = 5, LessThanOrEqual LE = 6, Between BT = 7
					$res.Properties.U_PrpConValue = $prop.Value
					$res.Properties.U_PrpConValueTo = $prop.ToValue
					$res.Properties.U_UnitOfMeasure = $prop.UoM
					$dummy = $res.Properties.Add()
				}
        
        
        
			}
    
			#Adding attachments to Resources
			[array]$resAttachments = $resourcesAtachementsDict[$csvItem.ResourceCode];
			if ($resAttachments.count -gt 0) {
				#Deleting all existing attachments
				$count = $res.Attachments.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $res.Attachments.DelRowAtPos(0);
				}
        
				#Adding the new data
				foreach ($att in $resAttachments) {
					$fileName = [System.IO.Path]::GetFileName($att.AttachmentPath)
					$res.Attachments.U_FileName = $fileName
					$res.Attachments.U_AttDate = [System.DateTime]::Today
					$res.Attachments.U_Path = $att.AttachmentPath
					$dummy = $res.Attachments.Add()
				}
			}

			#Adding planning info to Resources
			[array]$resPlanningInfos = $resourcesPlanningInfoDict[$csvItem.ResourceCode];
			if ($resPlanningInfos.count -gt 0) {
				#Deleting all existing attachments
				$count = $res.PlanningInformation.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $res.PlanningInformation.DelRowAtPos(0);
				}
        
				#Adding the new data
				foreach ($pi in $resPlanningInfos) {
					$res.PlanningInformation.U_PlanningInfo = $pi.PlanningInfo;
					switch ($pi.Table) {
						"CT_PF_OMOR" {
							$res.PlanningInformation.U_Table = [CompuTec.ProcessForce.API.Documents.Resource.ResourcePlanningInformationTableType]::ManufacturingOrder; 
							break;
						}
						"CT_PF_MOR12" {
							$res.PlanningInformation.U_Table = [CompuTec.ProcessForce.API.Documents.Resource.ResourcePlanningInformationTableType]::Operation; 
							break;
						}
						"CT_PF_MOR16" {
							$res.PlanningInformation.U_Table = [CompuTec.ProcessForce.API.Documents.Resource.ResourcePlanningInformationTableType]::Resource; 
							break;
						}
						"DYNAMIC" {
							$res.PlanningInformation.U_Table = [CompuTec.ProcessForce.API.Documents.Resource.ResourcePlanningInformationTableType]::Dynamic; 
							break;
						}
						Default {
							$res.PlanningInformation.U_Table = [CompuTec.ProcessForce.API.Documents.Resource.ResourcePlanningInformationTableType]::Undefined; 
						 }
					}
					$res.PlanningInformation.U_Field = $pi.Field;

					if ([string]::Equals($pi.SQLEnabled, "Y", [System.StringComparison]::InvariantCultureIgnoreCase)) {
						$res.PlanningInformation.U_IsSqlQueryEnabled = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
						$res.PlanningInformation.U_SqlQuery = $pi.SqlQuery;
					}
					else {
						$res.PlanningInformation.U_IsSqlQueryEnabled = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
					}
					$dummy = $res.PlanningInformation.Add()
				}
			}
 
			$message = 0
    
			#Adding or updating Resources depends on exists in the database
			if ($exists -eq $true) {
				$message = $res.Update()
			}
			else {
				$message = $res.Add()
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
			$ms = [string]::Format("Error when {0} Resource with Code {1} Details: {2}", $taskMsg, $csvItem.ResourceCode, $err);
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


