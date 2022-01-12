#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Substitutes
########################################################################
$SCRIPT_VERSION = "3.1"
# Last tested PF version: CompuTec ProcessForce 10.0 (), Release 14 (64-bit)
# Description:
#      Import Substitutes. Script add new or will update existing Substitutes.
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

$csvSubstitutesPath = -join ($csvImportCatalog, "Substitutes.csv")
$csvSubstitutesRevisionsPath = -join ($csvImportCatalog, "SubstitutesRevisions.csv")
$csvSubstitutesBomsPath = -join ($csvImportCatalog, "SubstitutesBOMs.csv")

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

function getKey($importKey1, $importKey2) {
	$importKey = [string]::Format("IK_{0}", $importKey1);

	if ([string]::IsNullOrWhiteSpace($importKey2) -eq $false) {
		$importKey += [string]::Format("_%IK2%_{0}", $importKey2);
	}

	return $importKey;
}

try {   
	[array] $csvSubstitutes = Import-Csv -Delimiter ';' $csvSubstitutesPath;
	[array] $csvSubstitutesRevisions = Import-Csv -Delimiter ';' $csvSubstitutesRevisionsPath;
	[array] $csvSubstitutesBoms = $null;
	if ((Test-Path -Path $csvSubstitutesBomsPath -PathType leaf) -eq $true) {
		[array]$csvSubstitutesBoms = Import-Csv -Delimiter ';' -Path $csvSubstitutesBomsPath
	}
	else {
		write-host "Substitues BOM Configurations - csv not available."
	}

	write-Host 'Preparing data: '
	$totalRows = $csvSubstitutes.Count + $csvSubstitutesRevisions.Count + $csvSubstitutesBoms.Count;
	$substitutesList = New-Object 'System.Collections.Generic.List[array]'
	$substitutesRevisionsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$substitutesBomsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}

	foreach ($row in $csvSubstitutes) {
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		$substitutesList.Add([array]$row);
	}

	foreach ($row in $csvSubstitutesRevisions) {
		$key = getKey -importKey1 $row.ImportKey1;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($substitutesRevisionsDict.ContainsKey($key) -eq $false) {
			$substitutesRevisionsDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $substitutesRevisionsDict[$key];
		$list.Add([array]$row);
	}

	foreach ($row in $csvSubstitutesBoms) {
		$key = getKey -importKey1 $row.ImportKey1 -importKey2 $row.ImportKey2;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		if ($substitutesBomsDict.ContainsKey($key) -eq $false) {
			$substitutesBomsDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $substitutesBomsDict[$key];
		$list.Add([array]$row);
	}

	Write-Host '';
	Write-Host 'Adding/updating data: ' -NoNewline;

	$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
	if ($substitutesList.Count -gt 1) {
		$total = $substitutesList.Count
	}
	else {
		$total = 1
	}
	$progressItterator = 0;
	$progress = 0;
	$beforeProgress = 0;
	foreach ($csvHeader in $substitutesList) {
		try {
			$progressItterator++;
			$progress = [math]::Round(($progressItterator * 100) / $total);
			if ($progress -gt $beforeProgress) {
				Write-Host $progress"% " -NoNewline
				$beforeProgress = $progress
			}
			$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OSIT"" WHERE ""Code"" = N'{0}'", $csvHeader.Code));
	
			#Creating object
			$md = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Substitute)
			#Checking if data already exists
			if ($rs.RecordCount -gt 0) {
				$dummy = $md.GetByKey($rs.Fields.Item(0).Value);
				$exists = $true
			}
			else {
				$md.Code = $csvHeader.Code;
				$exists = $false;
			}
			$headerKey = getKey -importKey1 $csvHeader.ImportKey1
			$md.U_Remarks = $csvHeader.Remarks;
			[array]$revisions = $substitutesRevisionsDict[$headerKey]
			if ($revisions.count -gt 0) {
				#region Deleting all existing revisions and boms
				$count = $md.Revisions.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $md.Revisions.DelRowAtPos(0);
				}

				$count = $md.BomConfigurations.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $md.BomConfigurations.DelRowAtPos(0);
				}
				#endregion

				#Adding the new data       
				foreach ($revision in $revisions) {
					$md.Revisions.U_Revision = $revision.Revision;
					$md.Revisions.U_SItemCode = $revision.SItemCode;
					$md.Revisions.U_SRevision = $revision.SRevision;
					
					if ($revision.Default -eq "Y") {
						$md.Revisions.U_IsDefault = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
					}
					else {
						$md.Revisions.U_IsDefault = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
					}

					if ([string]::IsNullOrWhiteSpace($revision.ValidFrom) -eq $false) {
						$md.Revisions.U_ValidFrom = $revision.ValidFrom;
					}

					if ([string]::IsNullOrWhiteSpace($revision.ValidTo) -eq $false) {
						$md.Revisions.U_ValidTo = $revision.ValidTo;
					}

					$md.Revisions.U_Ratio = [double]$revision.Ratio;

					if ($revision.RplItm -eq "Y") {
						$md.Revisions.U_RplItm = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
					}
					else {
						$md.Revisions.U_RplItm = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
					}

					if ($revision.RplCp -eq "Y") {
						$md.Revisions.U_RplCp = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
					}
					else {
						$md.Revisions.U_RplCp = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
					}

					if ($revision.RplSc -eq "Y") {
						$md.Revisions.U_RplCp = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
					}
					else {
						$md.Revisions.U_RplCp = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
					}

					$md.Revisions.U_Remarks = $revision.Remarks


					$revisionKey = getKey -importKey1 $revision.ImportKey1 -importKey2 $revision.ImportKey2;

					#region BOMs configurations
					[array]$boms = $substitutesBomsDict[$revisionKey]
					if ($boms.Count -gt 0) {
						foreach ($bc in $boms) {
							if ($md.BomConfigurations.IsRowFilled()) {
								$md.BomConfigurations.Add();
							}
							$md.BomConfigurations.U_ParentLineNo = $md.Revisions.U_LineNum;
							$md.BomConfigurations.U_ParentRevCode = $md.Revisions.U_Revision;
							$md.BomConfigurations.U_BomCode = $bc.BomCode;
							$md.BomConfigurations.U_BomRevCode = $bc.BomRevCode;

							if ($bc.DisableSubs -eq "Y") {
								$md.BomConfigurations.U_DisableSubs = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
							}
							else {
								$md.BomConfigurations.U_DisableSubs = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
							}

							if ([string]::IsNullOrWhiteSpace($bc.ValidFrom) -eq $false) {
								$md.BomConfigurations.U_ValidFrom = $bc.ValidFrom;
							}
		
							if ([string]::IsNullOrWhiteSpace($bc.ValidTo) -eq $false) {
								$md.BomConfigurations.U_ValidTo = $bc.ValidTo;
							}

							$md.BomConfigurations.U_BomRemarks = $bc.BomRemarks;
							$md.BomConfigurations.U_Remarks = $bc.Remarks;
						}
					}
					#endregion

					$dummy = $md.Revisions.Add();
				}
			}

			$message = 0
			#Adding or updating depends on exists in the database
			if ($exists -eq $true) {
				$message = $md.Update()
			}
			else {
				$message = $md.Add()
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
			$ms = [string]::Format("Error when {0} Substitutes with Code {1} Details: {2}", $taskMsg, $csvHeader.Code, $err);
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

