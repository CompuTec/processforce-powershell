#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Test Properties
########################################################################
$SCRIPT_VERSION = "3.1"
# Last tested PF version: ProcessForce 9.3 (9.30.210) (64-bit)
# Description:
#      Import Test Properties. Script add new or will update existing Properties.
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

$csvTestPropertiesPath = -join ($csvImportCatalog, "TestProperties.csv")
$csvTestPropertiesReferencesPath = -join ($csvImportCatalog, "TestPropertiesReferences.csv")
$csvTestPropertiesCertifiacteOfAnalysisPath = -join ($csvImportCatalog, "TestPropertiesCertifiacteOfAnalysis.csv")

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
	[array] $csvTestProps = Import-Csv -Delimiter ';' $csvTestPropertiesPath;
	
	if ((Test-Path -Path $csvTestPropertiesReferencesPath -PathType leaf) -eq $true) {
		[array] $csvTestPropsReferences = Import-Csv -Delimiter ';' $csvTestPropertiesReferencesPath;
	}
	else {
		write-host "Test Properties References - csv not available."
	}
	if ((Test-Path -Path $csvTestPropertiesCertifiacteOfAnalysisPath -PathType leaf) -eq $true) {
		[array] $csvTestPropsCertifiacteOfAnalysis = Import-Csv -Delimiter ';' $csvTestPropertiesCertifiacteOfAnalysisPath;
	}
	else {
		write-host "Test Properties Certificate Of Analysis - csv not available."
	}

	write-Host 'Preparing data: '
	$totalRows = $csvTestProps.Count + $csvTestPropsReferences.Count + $csvTestPropsCertifiacteOfAnalysis.Count;
	$testProplist = New-Object 'System.Collections.Generic.List[array]'
	$propRefDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$propCertDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}

	foreach ($row in $csvTestProps) {
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		$testProplist.Add([array]$row);
	}

	foreach ($row in $csvTestPropsReferences) {
		$key = $row.PropertyCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($propRefDict.ContainsKey($key) -eq $false) {
			$propRefDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $propRefDict[$key];
		
		$list.Add([array]$row);
	}

	foreach ($row in $csvTestPropsCertifiacteOfAnalysis) {
		$key = $row.PropertyCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($propCertDict.ContainsKey($key) -eq $false) {
			$propCertDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $propCertDict[$key];
		
		$list.Add([array]$row);
	}

	Write-Host '';
	Write-Host 'Adding/updating data: ' -NoNewLine;

	if ($testProplist.Count -gt 1) {
		$total = $testProplist.Count
	}
	else {
		$total = 1
	}
	$progressItterator = 0;
	$progress = 0;
	$beforeProgress = 0;
	
	$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

	foreach ($prop in $testProplist) {
		try {
			$progressItterator++;
			$progress = [math]::Round(($progressItterator * 100) / $total);
			if ($progress -gt $beforeProgress) {
				Write-Host $progress"% " -NoNewline
				$beforeProgress = $progress
			}
			$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OTPR"" WHERE ""U_TestPrpCode"" = N'{0}'", $prop.PropertyCode));
	
			#Creating Item Property object
			$testProperty = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::"TestProperty")
			#Checking that the property already exist
			if ($rs.RecordCount -gt 0) {
				$dummy = $testProperty.GetByKey($rs.Fields.Item(0).Value);
				$exists = $true
			}
			else {
				$testProperty.U_TestPrpCode = $prop.PropertyCode;
				$exists = $false
			}
   
			$testProperty.U_TestPrpName = $prop.PropertyName;
			$testProperty.U_TestPrpGrpCode = $prop.Group;
			$testProperty.U_TestPrpRemarks = $prop.Remarks;
	
			if ($prop.Group -ne '') {
				$testProperty.U_TestPrpGrpCode = $prop.Group 
			}
	
			#Data loading from the csv file - References for test properties
			[array]$references = $propRefDict[$prop.PropertyCode];
			if ($references.count -gt 0) {
				#Deleting all exisitng Revisions
				$count = $testProperty.Words.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $testProperty.Words.DelRowAtPos(0);
				}
				$testProperty.Words.SetCurrentLine(0);
         
				#Adding Revisions
				foreach ($ref in $references) {
					$testProperty.Words.U_WordCode = $ref.ReferenceCode;
					$dummy = $testProperty.Words.Add();
				}
			}
			
			#Certifiacate of Analysis
			[array]$certs = $propCertDict[$prop.PropertyCode];
			if ($certs.count -gt 0) {
				#Deleting all exisitng Revisions
				$count = $testProperty.BusinessPartnerRelations.Count
				for ($i = 0; $i -lt $count; $i++) {
					$testProperty.BusinessPartnerRelations.SetCurrentLine(0);
					if ($testProperty.BusinessPartnerRelations.IsRowFilled()) {
						$dummy = $testProperty.BusinessPartnerRelations.DelRowAtPos(0);
					}
				}
				$dummy = $testProperty.BusinessPartnerRelations.SetCurrentLine(0);

				#Adding Certificates
				foreach ($crt in $certs) {
					$testProperty.BusinessPartnerRelations.U_CardCode = $crt.CardCode;
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
					$testProperty.BusinessPartnerRelations.U_Expression = $EnumExpressionValue;

					if ([string]::IsNullOrWhiteSpace($crt.ValueFrom) -eq $false ) {
						$testProperty.BusinessPartnerRelations.U_ValueFrom = $crt.ValueFrom;
					}
					else {
						$testProperty.BusinessPartnerRelations.U_ValueFrom = 0;
					}
					$testProperty.BusinessPartnerRelations.U_ValueTo = $crt.ValueTo;

					if ([string]::IsNullOrWhiteSpace($crt.ValidFrom) -eq $false) {
						$testProperty.BusinessPartnerRelations.U_FromDate = $crt.ValidFrom
					}
					if ([string]::IsNullOrWhiteSpace($crt.ValidTo) -eq $false) {
						$testProperty.BusinessPartnerRelations.U_ToDate = $crt.ValidTo
					}
					$dummy = $testProperty.BusinessPartnerRelations.Add();
				}
			}
			$message = 0
			#Adding or updating Test Properties depends on exists in the database
			if ($exists -eq $true) {
				$message = $testProperty.Update()
			}
			else {
				$message = $testProperty.Add()
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
			$ms = [string]::Format("Error when {0} Test Property with Code {1} Details: {2}", $taskMsg, $prop.PropertyCode, $err);
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
