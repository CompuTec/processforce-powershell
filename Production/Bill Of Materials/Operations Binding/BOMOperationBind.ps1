#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import BOM Bindings
########################################################################
$SCRIPT_VERSION = "3.2"
# Last tested PF version: ProcessForce 9.3 (9.30.150) PL: 05 R1 HF1 (64-bit)
# Description:
#      Import BOM Bindings. Script will update existing BOM Bindings.
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

$csvBOMHeaderPath = -join ($csvImportCatalog, "BOMHeader.csv")
$csvBOMRoutingsOperationsBindPath = -join ($csvImportCatalog, "BOMRoutingsOperationsBind.csv")

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
	[array] $csvBOMs = Import-Csv -Delimiter ';' $csvBOMHeaderPath;
	[array] $csvBOMsBindings = Import-Csv -Delimiter ';' $csvBOMRoutingsOperationsBindPath;
	

	write-Host 'Preparing data: '
	$totalRows = $csvBOMs.Count + $csvBOMsBindings.Count;
	$BOMsList = New-Object 'System.Collections.Generic.List[array]'
	$BOMsBindingsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}

	foreach ($row in $csvBOMs) {
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		$BOMsList.Add([array]$row);
	}

	foreach ($row in $csvBOMsBindings) {
		$key = [string] $row.BOM_Header + "___" + [string] $row.Revision;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($BOMsBindingsDict.ContainsKey($key) -eq $false) {
			$BOMsBindingsDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $BOMsBindingsDict[$key];
		
		$list.Add([array]$row);
	}

	Write-Host '';


	$SQL_QUERY_RTGCODE = "SELECT RO.""U_RtgOprCode"" FROM ""@CT_PF_BOM12"" RO 
                                WHERE RO.""U_RtgCode"" =  N'{0}' AND RO.""U_OprCode"" =  N'{1}' AND RO.""U_OprSequence"" =  N'{2}' 
                                AND RO.""U_BomCode"" = N'{3}' AND RO.""U_RevCode"" = N'{4}'";
	$SQL_QUERY_BASELINE = "SELECT ISNULL(BS4.""U_LineNum"",ISNULL(BS3.""U_LineNum"",BS.""U_LineNum"")) 
                                FROM ""@CT_PF_OBOM"" B LEFT OUTER JOIN ""@CT_PF_BOM1"" BS ON B.""Code"" = BS.""Code"" AND 'IT' = N'{3}' AND BS.""U_ItemCode"" =  N'{2}' AND BS.""U_Sequence"" =  N'{4}'
                                LEFT OUTER JOIN ""@CT_PF_BOM3"" BS3 ON B.""Code"" = BS3.""Code"" AND 'CP' = N'{3}' AND BS3.""U_ItemCode"" =  N'{2}' AND BS3.""U_Sequence"" =  N'{4}'
                                LEFT OUTER JOIN ""@CT_PF_BOM4"" BS4 ON B.""Code"" = BS4.""Code"" AND 'SC' = N'{3}' AND BS4.""U_ItemCode"" =  N'{2}' AND BS4.""U_Sequence"" =  N'{4}'
                                WHERE B.""U_ItemCode"" =  N'{0}' AND B.""U_Revision"" =  N'{1}' AND ISNULL(BS4.""U_LineNum"",ISNULL(BS3.""U_LineNum"",ISNULL(BS.""U_LineNum"",-1))) != -1";
	if ($pfcCompany.DbServerType -eq [SAPbobsCOM.BoDataServerTypes]::dst_HANADB) {
		$SQL_QUERY_BASELINE = "SELECT IFNULL(BS4.""U_LineNum"",IFNULL(BS3.""U_LineNum"",BS.""U_LineNum"")) 
                                FROM ""@CT_PF_OBOM"" B LEFT OUTER JOIN ""@CT_PF_BOM1"" BS ON B.""Code"" = BS.""Code"" AND 'IT' = N'{3}' AND BS.""U_ItemCode"" =  N'{2}' AND BS.""U_Sequence"" =  N'{4}'
                                LEFT OUTER JOIN ""@CT_PF_BOM3"" BS3 ON B.""Code"" = BS3.""Code"" AND 'CP' = N'{3}' AND BS3.""U_ItemCode"" =  N'{2}' AND BS3.""U_Sequence"" =  N'{4}'
                                LEFT OUTER JOIN ""@CT_PF_BOM4"" BS4 ON B.""Code"" = BS4.""Code"" AND 'SC' = N'{3}' AND BS4.""U_ItemCode"" =  N'{2}' AND BS4.""U_Sequence"" =  N'{4}'
                                WHERE B.""U_ItemCode"" =  N'{0}' AND B.""U_Revision"" =  N'{1}' AND IFNULL(BS4.""U_LineNum"",IFNULL(BS3.""U_LineNum"",IFNULL(BS.""U_LineNum"",-1))) != -1";
	};
    



	Write-Host 'Adding/updating data: ' -NoNewLine;
    
	$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
	#Data loading from a csv file
	if ($BOMsList.Count -gt 1) {
		$total = $BOMsList.Count
	}
	else {
		$total = 1
	}
	$progressItterator = 0;
	$progress = 0;
	$beforeProgress = 0;
	foreach ($csvItem in $BOMsList) {
		try {
			
			#Creating BOM object
			$bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"BillOfMaterial")
			#Checking that the BOM already exist
			$retValue = $bom.GetByItemCodeAndRevision($csvItem.BOM_Header, $csvItem.Revision)
			if ($retValue -ne 0) { 
				$err = [string]::Format("BOM with ItemCode: {0} and Revision: {1} don't exists.", $csvItem.BOM_Header, $csvItem.Revision);
				Throw [System.Exception]($err);
			}
       
			$key = [string] $csvItem.BOM_Header + "___" + [string] $csvItem.Revision;
			[array]$bomBindings = $BOMsBindingsDict[$key];
			if ($bomBindings.count -eq 0) {
				$err = [string]::Format("Missing bindings positions for Item Code: {0} and Revision: {1}", $csvItem.BOM_Header, $csvItem.Revision);
				Throw [System.Exception]($err);
			}

			#Deleting all existing routings, operations, resources
			$count = $bom.RoutingsOperationInputOutput.Count;
			for ($i = 0; $i -lt $count; $i++) {
				$dummy = $bom.RoutingsOperationInputOutput.DelRowAtPos(0);
			}            
			$dummy = $bom.RoutingsOperationInputOutput.SetCurrentLine($bom.RoutingsOperationInputOutput.Count - 1);
         
			#Adding a new data - Bind
			foreach ($bb in $bomBindings) {
				$rs.DoQuery([string]::Format($SQL_QUERY_RTGCODE, $bb.RoutingCode, $bb.OperationCode, $bb.OperationSequence, $bb.BOM_Header, $bb.Revision));

				if ($rs.RecordCount -gt 0) {
                
					$bom.RoutingsOperationInputOutput.U_RtgOprCode = $rs.Fields.Item(0).Value
				}
				else {
					$err = [System.String]::Format("Error adding binding Routing: {0}, Operation: {1}, OperationSequence: {2}
                 - RECORD NOT FOUND", $bb.RoutingCode, $bb.OperationCode, $bb.OperationSequence);
					Throw [System.Exception]($err);
				}

				$dummy = $rs.DoQuery([string]::Format($SQL_QUERY_BASELINE, $bb.BOM_Header, $bb.Revision, $bb.ItemCode, $bb.ItemType , $bb.ItemSequence));

				$bom.RoutingsOperationInputOutput.U_RtgCode = $bb.RoutingCode
				$bom.RoutingsOperationInputOutput.U_OprCode = $bb.OperationCode
				$bom.RoutingsOperationInputOutput.U_ItemCode = $bb.ItemCode
				$bom.RoutingsOperationInputOutput.U_Direction = $bb.Direction

				if ($bb.ItemType -eq 'IT') {
					$bom.RoutingsOperationInputOutput.U_ItemType = [CompuTec.ProcessForce.API.Enumerators.ManufacturingComponentType]::Item
				}
				if ($bb.ItemType -eq 'CP') {
					$bom.RoutingsOperationInputOutput.U_ItemType = [CompuTec.ProcessForce.API.Enumerators.ManufacturingComponentType]::Cooproduct
				}
				if ($bb.ItemType -eq 'SC') {
					$bom.RoutingsOperationInputOutput.U_ItemType = [CompuTec.ProcessForce.API.Enumerators.ManufacturingComponentType]::Scrap
				}

				if ($bb.TimeCalc -eq 'Y') {
					$bom.RoutingsOperationInputOutput.U_InTimeCalc = 'Y'
				}
				else {
					$bom.RoutingsOperationInputOutput.U_InTimeCalc = 'N'
				}

				if ($rs.RecordCount -gt 0) {
                
					$bom.RoutingsOperationInputOutput.U_BaseLine = $rs.Fields.Item(0).Value
				}
				else {
					$err = [System.String]::Format("Error adding binding BOM: {0}, Revision: {1}, ItemCode: {2},
                IemType: {3}, Sequence: {4} - RECORD NOT FOUND", $bb.BOM_Header, $bb.Revision, $bb.ItemCode, $bb.ItemType , $bb.ItemSequence);
					Throw [System.Exception]($err);
				}

				$dummy = $bom.RoutingsOperationInputOutput.Add();
			}

			$message = 0

			$message = $bom.Update()
			if ($message -lt 0) {    
				$err = $pfcCompany.GetLastErrorDescription()
				Throw [System.Exception] ($err)
			}
			
		}
		Catch {
			$err = $_.Exception.Message;
			$taskMsg = "updating"
			$ms = [string]::Format("Error when {0} BOM Bindings with ItemCode {1} and Revision: {2}. Details: {3}", $taskMsg, $csvItem.BOM_Header, $csvItem.Revision, $err);
			Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
			if ($pfcCompany.InTransaction) {
				$pfcCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
			} 
		}
		Finally {
			#region progress
			$progressItterator++;
			$progress = [math]::Round(($progressItterator * 100) / $total);
			if ($progress -gt $beforeProgress) {
				Write-Host $progress"% " -NoNewline
				$beforeProgress = $progress
			}
			#endregion
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

