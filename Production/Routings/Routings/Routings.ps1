﻿#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Routings
########################################################################
$SCRIPT_VERSION = "3.2"
# Last tested PF version: ProcessForce 9.3 (9.30.210) (64-bit)
# Description:
#      Import Routings. Script add new or will update existing Routings.
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

$csvRoutingsFilePath = -join ($csvImportCatalog, "Routings.csv")
$csvRoutingOperationsFilePath = -join ($csvImportCatalog, "RoutingOperations.csv")
$csvRoutingOperationPropertiesFilePath = -join ($csvImportCatalog, "RoutingOperationsProperties.csv")
$csvRoutingOperationResourcesFilePath = -join ($csvImportCatalog, "RoutingOperationsResources.csv")
$csvRoutingOperationRelationsFilePath = -join ($csvImportCatalog, "RoutingOperationsRelations.csv")
$csvRoutingOperationResourcesPropertiesFilePath = -join ($csvImportCatalog, "RoutingOperationsResourcesProperties.csv")

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
	write-host ""
	[array]$csvRoutings = Import-Csv -Delimiter ';' -Path $csvRoutingsFilePath;
	[array]$csvRoutingOperations = Import-Csv -Delimiter ';' -Path $csvRoutingOperationsFilePath
	if ((Test-Path -Path $csvRoutingOperationPropertiesFilePath -PathType leaf) -eq $true) {
		[array]$csvRoutingOperationProperties = Import-Csv -Delimiter ';' -Path $csvRoutingOperationPropertiesFilePath
	}
	else {
		[array] $csvRoutingOperationProperties = $null;
		write-host "Item Properties References - csv not available."
	}

	[array]$csvRoutingOperationResources = Import-Csv -Delimiter ';' -Path $csvRoutingOperationResourcesFilePath
	if ((Test-Path -Path $csvRoutingOperationRelationsFilePath -PathType leaf) -eq $true) {
		[array]$csvRoutingOperationRelations = Import-Csv -Delimiter ';' -Path $csvRoutingOperationRelationsFilePath
	}
	else {
		[array] $csvRoutingOperationRelations = $null;
		write-host "Advanced Operation Relations - csv not available."
	}
	
	if ((Test-Path -Path $csvRoutingOperationResourcesPropertiesFilePath -PathType leaf) -eq $true) {
		[array]$csvRoutingOperationResourcesProperties = Import-Csv -Delimiter ';' -Path $csvRoutingOperationResourcesPropertiesFilePath
	}
	else {
		[array] $csvRoutingOperationResourcesProperties = $null;
		write-host "Resource Properties - csv not available."
	}


	write-Host 'Preparing data: ' -NoNewline
	$totalRows = $csvRoutings.Count + $csvRoutingOperations.Count + $csvRoutingOperationProperties.Count + $csvRoutingOperationResources.Count + $csvRoutingOperationRelations.Count + $csvRoutingOperationResourcesProperties.Count;
    
	$routingsList = New-Object 'System.Collections.Generic.List[array]';
	$dictionaryRoutingsOperations = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$dictionaryRoutingsOperationsProperties = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$dictionaryRoutingsOperationsResources = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$dictionaryRoutingsOperationsRelations = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$dictionaryResourceProperties = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
    
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}

	foreach ($row in $csvRoutings) {
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		$routingsList.Add([array]$row);
	}

	foreach ($row in $csvRoutingOperations) {
		$key = $row.RoutingCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryRoutingsOperations.ContainsKey($key)) {
			$list = $dictionaryRoutingsOperations[$key];
		}
		else {
			$list = New-Object System.Collections.Generic.List[array];
			$dictionaryRoutingsOperations[$key] = $list;
		}
    
		$list.Add([array]$row);
	}

	foreach ($row in $csvRoutingOperationProperties) {
		$key = $row.RoutingCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryRoutingsOperationsProperties.ContainsKey($key)) {
			$list = $dictionaryRoutingsOperationsProperties[$key];
		}
		else {
			$list = New-Object System.Collections.Generic.List[array];
			$dictionaryRoutingsOperationsProperties[$key] = $list;
		}
    
		$list.Add([array]$row);
	}

	foreach ($row in $csvRoutingOperationResources) {
		$key = $row.RoutingCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryRoutingsOperationsResources.ContainsKey($key)) {
			$list = $dictionaryRoutingsOperationsResources[$key];
		}
		else {
			$list = New-Object System.Collections.Generic.List[array];
			$dictionaryRoutingsOperationsResources[$key] = $list;
		}
    
		$list.Add([array]$row);
	}

	foreach ($row in $csvRoutingOperationRelations) {
		$key = $row.RoutingCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryRoutingsOperationsRelations.ContainsKey($key)) {
			$list = $dictionaryRoutingsOperationsRelations[$key];
		}
		else {
			$list = New-Object System.Collections.Generic.List[array];
			$dictionaryRoutingsOperationsRelations[$key] = $list;
		}
    
		$list.Add([array]$row);
	}
    
	foreach ($row in $csvRoutingOperationResourcesProperties) {
		$key = $row.RoutingCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryResourceProperties.ContainsKey($key)) {
			$list = $dictionaryResourceProperties[$key];
		}
		else {
			$list = New-Object System.Collections.Generic.List[array];
			$dictionaryResourceProperties[$key] = $list;
		}
    
		$list.Add([array]$row);
	}

	Write-Host '';
	Write-Host 'Adding/updating data: ' -NoNewline;
	if ($routingsList.Count -gt 1) {
		$total = $routingsList.Count
	}
	else {
		$total = 1
	}
	$progressItterator = 0;
	$progress = 0;
	$beforeProgress = 0;
	foreach ($csvItem in $routingsList) {
		try {
			$progressItterator++;
			$progress = [math]::Round(($progressItterator * 100) / $total);
			if ($progress -gt $beforeProgress) {
				Write-Host $progress"% " -NoNewline
				$beforeProgress = $progress
			}
			$key = $csvItem.RoutingCode;
			#Creating Operation object
			$routing = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Routing)
			#Checking that the operation already exist    
			$retValue = $routing.GetByRtgCode($csvItem.RoutingCode)
			if ($retValue -ne 0) { 
				#Adding the new data
				$routing.U_RtgCode = $csvItem.RoutingCode
				$exists = $false;
			}
			else {
				$exists = $true;
			}
			$routing.U_RtgName = $csvItem.RoutingName
			$routing.U_Active = $csvItem.Active #enum type; 1 = Yes, 2 = No
			$routing.U_Remarks = $csvItem.Remarks

			$routingOperations = $dictionaryRoutingsOperations[$key];
			if ($routingOperations.count -gt 0) {
				#Deleting all existing operations
				$count = $routing.Operations.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $routing.Operations.DelRowAtPos(0);
				}
				$count = $routing.OperationResourceProperties.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $routing.OperationResourceProperties.DelRowAtPos(0);
				}
				$drivers = New-Object 'System.Collections.Generic.Dictionary[String,int]'
				$operLineNumDict = New-Object 'System.Collections.Generic.Dictionary[int,int]';
				$operCodeDict = New-Object 'System.Collections.Generic.Dictionary[int,string]';
				#Adding the new data       
				foreach ($rtOper in $routingOperations) {
					$routing.Operations.U_OprCode = $rtOper.OperationCode
					if($routing.Operations.U_RtgOprCode -eq 0)
                    {
						$msg = [string]::Format("Operation {0} does not exists", [string]$rtOper.OperationCode);
						throw [System.Exception]($msg);
					}
					$overlaySequence = $rtOper.OperationOverlaySequence;
					if([string]::IsNullOrWhiteSpace($overlaySequence) -eq $false) {
						$overlayId = -1;
						if ($operLineNumDict.TryGetValue($overlaySequence, [ref] $overlayId) -eq $false) {
							Throw [System.Exception] (([string]::Format("Can't configure overlay for operation: {0}. Operation Sequence: {1} don't exists", $rtOper.OperationCode, $overlaySequence)));
						}
 						$routing.Operations.U_OprOverlayId = $overlayId 
						$routing.Operations.U_OprOverlayCode = $operCodeDict[$overlaySequence]
						$routing.Operations.U_OprOverlayQty = $rtOper.OverlayQty
					}

					$routing.Operations.U_OprSequence = $rtOper.Sequence
					$routing.Operations.U_Remarks = $rtOper.Remarks
					# Relation Types:
					# "NO" - NOne, 
					# "LF" - Last Finish,
					# "FF" - First Finish,
					# "SO" - Starting Operation,
					# "OV" - Overloading
					switch ($rtOper.RelationType) {
						"LF" {
							$routing.Operations.U_RelationType = [CompuTec.ProcessForce.API.Enumerators.OperationRelationType]::LastFinish; 
							$routing.Operations.U_HasRelations = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
							break;
						}
						"FF" { 
							$routing.Operations.U_RelationType = [CompuTec.ProcessForce.API.Enumerators.OperationRelationType]::FirstFinish; 
							$routing.Operations.U_HasRelations = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
							break; 
						}
						"SO" { 
							$routing.Operations.U_RelationType = [CompuTec.ProcessForce.API.Enumerators.OperationRelationType]::StartingOperation; 
							$routing.Operations.U_HasRelations = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
							break;
						}
						"OV" {
							$routing.Operations.U_RelationType = [CompuTec.ProcessForce.API.Enumerators.OperationRelationType]::Overloading; 
							$routing.Operations.U_HasRelations = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
							break; 
						}
						Default {
							$routing.Operations.U_RelationType = [CompuTec.ProcessForce.API.Enumerators.OperationRelationType]::None;
							$routing.Operations.U_HasRelations = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
						}
					}	


					$drivers.Add($routing.Operations.U_OprSequence, $routing.Operations.U_RtgOprCode);
					$operLineNumDict.Add($routing.Operations.U_OprSequence, $routing.Operations.U_LineNum);
					$operCodeDict.Add($routing.Operations.U_OprSequence, $routing.Operations.U_OprCode);
					$dummy = $routing.Operations.Add()
				}
				
				#region operation advanced relations
				$routingOperationsRelations = $dictionaryRoutingsOperationsRelations[$key];
				if ($routingOperationsRelations.count -gt 0) {
					#Deleting all existing properties
				
					$count = $routing.Relations.Count
					for ($i = $count - 1; $i -ge 0; $i--) {
						$routing.Relations.SetCurrentLine($i);
							$dummy = $routing.Relations.DelRowAtPos($i);
					}

					$routing.Relations.SetCurrentLine($routing.Relations.Count - 1);
					#Adding the new data
					foreach ($rtgOprRel in $routingOperationsRelations) {
						if ($routing.Relations.IsRowFilled()) {
							$dummy = $routing.Relations.Add();
						}
						$routing.Relations.U_BaseLine = $operLineNumDict[$rtgOprRel.Sequence];
						$operLineNum = $routing.Operations.Where( { $_.U_OprSequence -eq $rtgOprRel.RelOprSequence })[0].U_LineNum;
						if (-not $operLineNum -gt 0) {
							throw [System.Exception]([string]::Format("Couldn't find operation with sequence: '{0}' in Routing: '{1}' for Relation", [string]$rtgOprRel.RelOprSequence, [string]$csvItem.RoutingCode));
						}
						$routing.Relations.U_POprLine = $operLineNum;
					}
				}
				#endregion
				
				#operation properties
				$routingsOperationsProperties = $dictionaryRoutingsOperationsProperties[$key];
				if ($routingsOperationsProperties.count -gt 0) {
					#Deleting all existing properties
					$count = $routing.OperationProperties.Count
					for ($i = 0; $i -lt $count; $i++) {
						$dummy = $routing.OperationProperties.DelRowAtPos(0);
					}
		        
					#Adding the new data       
					foreach ($prop in $routingsOperationsProperties) {
						$routing.OperationProperties.U_RtgOprCode = $drivers[$prop.Sequence]
						$routing.OperationProperties.U_PrpCode = $prop.PropertiesCode
						$routing.OperationProperties.U_PrpConType = $prop.Condition #enum ConditionType; Equal EQ = 1, NotEqual NE = 2, GratherThan GT = 3, GratherThanOrEqual GE = 4, LessThan LT = 5, LessThanOrEqual LE = 6, Between BT = 7
						$routing.OperationProperties.U_PrpConValue = $prop.Value
						$routing.OperationProperties.U_PrpConValueTo = $prop.ToValue
						$routing.OperationProperties.U_UnitOfMeasure = $prop.UoM
						$dummy = $routing.OperationProperties.Add()
					}
				}
		
				#Deleting default resources copied from operations   
				$count = $routing.OperationResources.Count
				for ($i = $count - 1; $i -ge 0; $i--) {
					$dummy = $routing.OperationResources.DelRowAtPos($i);
				}    
				$count = $routing.OperationResourceProperties.Count - 1
				for ($i = $count - 1; $i -ge 0; $i--) {
					$dummy = $routing.OperationResourceProperties.DelRowAtPos($i);      
				}
				$driversOprRsc = New-Object 'System.Collections.Generic.Dictionary[String,int]'
				#Adding resources for operations   
				$routingsOperationsResources = $dictionaryRoutingsOperationsResources[$key];
				if ($routingsOperationsResources.count -gt 0) {
					foreach ($rtgOperResc in $routingsOperationsResources) {
						$routing.OperationResources.U_RtgOprCode = $drivers[$rtgOperResc.Sequence];
						$routing.OperationResources.U_RscCode = $rtgOperResc.ResourceCode

						if ($rtgOperResc.MachineCode -ne '') {
							if ($routing.OperationResources.U_RscType -eq [CompuTec.ProcessForce.API.Enumerators.ResourceType]::Tool) {
								$routing.OperationResources.U_MachineCode = $rtgOperResc.MachineCode;
							}
						}

						$routing.OperationResources.U_OcrCode = $rtgOperResc.OcrCode
						$routing.OperationResources.U_OcrCode2 = $rtgOperResc.OcrCode2
						$routing.OperationResources.U_OcrCode3 = $rtgOperResc.OcrCode3
						$routing.OperationResources.U_OcrCode4 = $rtgOperResc.OcrCode4
						$routing.OperationResources.U_OcrCode5 = $rtgOperResc.OcrCode5
						$routing.OperationResources.U_IsDefault = $rtgOperResc.Default
						$routing.OperationResources.U_IssueType = $rtgOperResc.IssueType
						$routing.OperationResources.U_QueueTime = $rtgOperResc.QueTime
						$routing.OperationResources.U_QueueRate = $rtgOperResc.QueTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
						$routing.OperationResources.U_SetupTime = $rtgOperResc.SetupTime
						$routing.OperationResources.U_SetupRate = $rtgOperResc.SetupTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
						$routing.OperationResources.U_RunTime = $rtgOperResc.RunTime
						$routing.OperationResources.U_RunRate = $rtgOperResc.RunTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
						$routing.OperationResources.U_StockTime = $rtgOperResc.StockTime
						$routing.OperationResources.U_StockRate = $rtgOperResc.StockTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
					
					
					
						if ($rtgOperResc.HasCycles -ne '') {
							if ($rtgOperResc.HasCycles -eq 'Y') {
								$routing.OperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes
								if ($rtgOperResc.CycleCapacity -ne '') {
									$routing.OperationResources.U_CycleCap = $rtgOperResc.CycleCapacity
								}
							}
							else {
								$routing.OperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
							}
						}
					
						$routing.OperationResources.U_Remarks = $rtgOperResc.Remarks
						if ($rtgOperResc.Project -ne '') {
							$routing.OperationResources.U_Project = $rtgOperResc.Project
						}
					
						$key = $rtgOperResc.Sequence + '@#@' + $routing.OperationResources.U_RscCode
						$driversOprRsc.Add($key, $routing.OperationResources.U_RtgOprRscCode);
						$dummy = $routing.OperationResources.Add()
					}
				
					#Adding resources properties to Operations
					$opResourceProperties = $dictionaryResourceProperties[$key];
					if ($opResourceProperties.count -gt 0) {
						#Deleting all existing resources
						$count = $routing.OperationResourceProperties.Count - 1
						for ($i = $count - 1; $i -ge 0; $i--) {
							[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
							$dummy = $routing.OperationResourceProperties.DelRowAtPos($i); 
						}

						#Adding the new data
						foreach ($opResProp in $opResourceProperties) {
							$key = $opResProp.Sequence + '@#@' + $opResProp.RoutingCode
							$routing.OperationResourceProperties.U_RtgOprCode = $drivers[$opResProp.Sequence]
							$routing.OperationResourceProperties.U_RtgOprRscCode = $driversOprRsc[$key]
							$routing.OperationResourceProperties.U_PrpCode = $opResProp.PropertiesCode
							$routing.OperationResourceProperties.U_PrpConType = $opResProp.Condition
							$routing.OperationResourceProperties.U_PrpConValue = $opResProp.Value
							$routing.OperationResourceProperties.U_PrpConValueTo = $opResProp.ToValue
							$routing.OperationResourceProperties.U_UnitOfMeasure = $opResProp.UoM
						
							$dummy = $routing.OperationResourceProperties.Add()
			            
						}
					}
				}
			}
      
			$message = 0
    
			#Adding or updating Routings depends on exists in the database
			if ($exists -eq $true) {
				$message = $routing.Update()
			}
			else {
				$message = $routing.Add()
			}
			if ($message -lt 0) {    
				$err = $pfcCompany.GetLastErrorDescription()
				Throw [System.Exception] ($err)
			}
		}
		Catch {
			$err = $_.Exception.Message;
			$errInner = if([string]::IsNullOrEmpty($_.Exception.InnerException) -eq $false) { $_.Exception.InnerException.ToString() } else { [string]::Empty };
			if ($exists -eq $true) {
				$taskMsg = "updating";
			}
			else {
				$taskMsg = "adding"
			}
			$ms = [string]::Format("Error when {0} Routing with Code {1} Details: {2}. {3}", $taskMsg, $csvItem.RoutingCode, [string]$err, [string]$errInner);
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

