#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Item Properties
########################################################################
$SCRIPT_VERSION = "3.1"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Item Properties. Script add new or will update existing Properties.
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

$csvItemPropertiesPath = -join ($csvImportCatalog, "ItemProperties.csv")
$csvItemPropertiesReferencesPath = -join ($csvImportCatalog, "ItemPropertiesReferences.csv")
$csvItemPropertiesCertifiacteOfAnalysisPath = -join ($csvImportCatalog, "ItemPropertiesCertifiacteOfAnalysis.csv")

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
	[array] $csvItmProps = Import-Csv -Delimiter ';' $csvItemPropertiesPath;
	
	if ((Test-Path -Path $csvItemPropertiesReferencesPath -PathType leaf) -eq $true) {
		[array] $csvItmPropsReferences = Import-Csv -Delimiter ';' $csvItemPropertiesReferencesPath;
	}
	else {
		write-host "Item Properties References - csv not available."
	}
	if ((Test-Path -Path $csvItemPropertiesCertifiacteOfAnalysisPath -PathType leaf) -eq $true) {
		[array] $csvItemPropertiesCertifiacteOfAnalysis = Import-Csv -Delimiter ';' $csvItemPropertiesCertifiacteOfAnalysisPath;
	}
	else {
		write-host "Item Properties Certificate Of Analysis - csv not available."
	}

	write-Host 'Preparing data: '
	$totalRows = $csvItmProps.Count + $csvItmPropsReferences.Count + $csvItemPropertiesCertifiacteOfAnalysis.Count;
	$itemProplist = New-Object 'System.Collections.Generic.List[array]'
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

	foreach ($row in $csvItmProps) {
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		$itemProplist.Add([array]$row);
	}

	foreach ($row in $csvItmPropsReferences) {
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

	foreach ($row in $csvItemPropertiesCertifiacteOfAnalysis) {
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
	Write-Host 'Adding/updating data: ' -NoNewline;

	$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
	if ($itemProplist.Count -gt 1) {
		$total = $itemProplist.Count
	}
	else {
		$total = 1
	}
	$progressItterator = 0;
	$progress = 0;
	$beforeProgress = 0;
	foreach ($prop in $itemProplist) {
		try {
			$progressItterator++;
			$progress = [math]::Round(($progressItterator * 100) / $total);
			if ($progress -gt $beforeProgress) {
				Write-Host $progress"% " -NoNewline
				$beforeProgress = $progress
			}
			$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OIPR"" WHERE ""U_PrpCode"" = N'{0}'", $prop.PropertyCode));
	
			#Creating Item Property object
			$itmProperty = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::"ItemProperty")
			#Checking that the property already exist
			if ($rs.RecordCount -gt 0) {
				$dummy = $itmProperty.GetByKey($rs.Fields.Item(0).Value);
				$exists = $true
			}
			else {
				$itmProperty.U_PrpCode = $prop.PropertyCode;
				$exists = $false
			}
   
			$itmProperty.U_PrpName = $prop.PropertyName;
			$itmProperty.U_UoM = $prop.UoM;
	
	
	
			if ($prop.Group -ne '') {
				$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OIPG"" WHERE ""U_GrpCode"" = N'{0}'", $prop.Group));
				$itmProperty.U_GrpCode = $rs.Fields.Item(0).Value
	
	
				if ($prop.Subgroup -ne '') {
					$itmProperty.U_SubGrpLineNo = $prop.Subgroup
				}
			}
	
			if ($prop.QualityControlTesting -eq 'Y') {
				$itmProperty.U_IsQcTesting = 'Y'
			}
			else {
				$itmProperty.U_IsQcTesting = 'N'
			}
   
			if ($prop.ProductionOrders -eq 'Y') {
				$itmProperty.U_ProdOrders = 'Y'
			}
			else {
				$itmProperty.U_ProdOrders = 'N'
			}
	
			if ($prop.ShipmentsDocumentation -eq 'Y') {
				$itmProperty.U_ShipDoc = 'Y'
			}
			else {
				$itmProperty.U_ShipDoc = 'N'
			}
	
			if ($prop.PickLists -eq 'Y') {
				$itmProperty.U_PickLists = 'Y'
			}
			else {
				$itmProperty.U_PickLists = 'N'
			}
	
			if ($prop.MSDS -eq 'Y') {
				$itmProperty.U_MSDS = 'Y'
			}
			else {
				$itmProperty.U_MSDS = 'N'
			}
	
			if ($prop.PurchaseOrders -eq 'Y') {
				$itmProperty.U_PurOrders = 'Y'
			}
			else {
				$itmProperty.U_PurOrders = 'N'
			}
	
			if ($prop.Returns -eq 'Y') {
				$itmProperty.U_Returns = 'Y'
			}
			else {
				$itmProperty.U_Returns = 'N'
			}
	
			if ($prop.Other -eq 'Y') {
				$itmProperty.U_Other = 'Y'
			}
			else {
				$itmProperty.U_Other = 'N'
			}
	
			#Data loading from the csv file - References for itmes properties
			[array]$references = $propRefDict[$prop.PropertyCode];
			if ($references.count -gt 0) {
				#Deleting all exisitng Revisions
				$count = $itmProperty.Words.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $itmProperty.Words.DelRowAtPos(0);
				}
				$dummy = $itmProperty.Words.SetCurrentLine(0);
         
				#Adding References
				foreach ($ref in $references) {
					$itmProperty.Words.U_WordCode = $ref.ReferenceCode;
					$dummy = $itmProperty.Words.Add();
				}
			}


			#Certifiacate of Analysis
			[array]$certs = $propCertDict[$prop.PropertyCode];
			if ($certs.count -gt 0) {
				#Deleting all exisitng Revisions
				$count = $itmProperty.BusinessPartnerRelations.Count
				for ($i = 0; $i -lt $count; $i++) {
					$itmProperty.BusinessPartnerRelations.SetCurrentLine(0);
					if ($itmProperty.BusinessPartnerRelations.IsRowFilled()) {
						$dummy = $itmProperty.BusinessPartnerRelations.DelRowAtPos(0);
					}
				}
				$dummy = $itmProperty.BusinessPartnerRelations.SetCurrentLine(0);

				#Adding Certificates
				foreach ($crt in $certs) {
					$itmProperty.BusinessPartnerRelations.U_CardCode = $crt.CardCode;
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
					$itmProperty.BusinessPartnerRelations.U_Expression = $EnumExpressionValue;

					if ([string]::IsNullOrWhiteSpace($crt.ValueFrom) -eq $false ) {
						$itmProperty.BusinessPartnerRelations.U_ValueFrom = $crt.ValueFrom;
					}
					else {
						$itmProperty.BusinessPartnerRelations.U_ValueFrom = 0;
					}
					$itmProperty.BusinessPartnerRelations.U_ValueTo = $crt.ValueTo;

					if ([string]::IsNullOrWhiteSpace($crt.ValidFrom) -eq $false) {
						$itmProperty.BusinessPartnerRelations.U_FromDate = $crt.ValidFrom
					}
					if ([string]::IsNullOrWhiteSpace($crt.ValidTo) -eq $false) {
						$itmProperty.BusinessPartnerRelations.U_ToDate = $crt.ValidTo
					}
					$dummy = $itmProperty.BusinessPartnerRelations.Add();
				}
			}

			$message = 0
			#Adding or updating Items Properties depends on exists in the database
			if ($exists -eq $true) {
				$message = $itmProperty.Update()
			}
			else {
				$message = $itmProperty.Add()
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
			$ms = [string]::Format("Error when {0} Item Property with Code {1} Details: {2}", $taskMsg, $prop.PropertyCode, $err);
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
