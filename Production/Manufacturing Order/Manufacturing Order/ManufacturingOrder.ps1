#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Manufacturing Order Add script - tutorial 
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 16 (64-bit)
# Description:
#      Base on csv files this script will add Manufacturing Orders.
# Troubleshooting:
#   https://connect.computec.pl/display/PF100EN/PowerShell+FAQ
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

#region Script parametersy

$csvImportCatalog = $PSScriptRoot + "\"
#If you are using lower version of PowerShell than PowerShell 4.0 you can use static path
# $csvImportCatalog = "C:\PS\PF\TestProtocols\";

$csvManufacturingOrdersPath = -join ($csvImportCatalog, "ManufacturingOrder.csv")
$csvManufacturingOrderOperationsPath = -join ($csvImportCatalog, "ManufacturingOrderOperations.csv")

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

#Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
Write-Host 'Preparing data: '
try {
	#region import csv files
	[array]$csvManufacturingOrders = Import-Csv -Delimiter ';' -Path $csvManufacturingOrdersPath
	[array]$csvManufacturingOrdersOperations = Import-Csv -Delimiter ';' -Path $csvManufacturingOrderOperationsPath

	#endregion

	$manufacturingOrdersList = New-Object 'System.Collections.Generic.List[array]';
	$manufacturingOrdersOperationsDictionary = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]';

	$totalRows = $csvManufacturingOrders.Count + $csvManufacturingOrdersOperations.Count;

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
	foreach ($row in $csvManufacturingOrders) {
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		$manufacturingOrdersList.Add([array]$row);
	}

	foreach ($row in $csvManufacturingOrdersOperations) {
		$key = $row.Key;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($manufacturingOrdersOperationsDictionary.ContainsKey($key) -eq $false) {
			$manufacturingOrdersOperationsDictionary[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $manufacturingOrdersOperationsDictionary[$key];
	
		$list.Add([array]$row);
	}
	#endregion


	Write-Host '';
	Write-Host 'Adding/updating data: ' -NoNewline;
	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	if ($manufacturingOrdersList.Count -gt 1) {
		$total = $manufacturingOrdersList.Count
	}
	else {
		$total = 1
	}


	foreach ($csvItem in $manufacturingOrdersList) {
		try {
			$key = $csvItem.Key;
			$progressItterator++;
			$progres = [math]::Round(($progressItterator * 100) / $total);
			if ($progres -gt $beforeProgress) {
				Write-Host $progres"% " -NoNewline
				$beforeProgress = $progres
			}
			#Creating BOM object
			$mo = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ManufacturingOrder)
			$bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::BillOfMaterial)
			$dummy = $bom.GetByItemCodeAndRevision($csvItem.ItemCode, $csvItem.Revision);
			$mo.U_BOMCode = $bom.Code;
			$mo.U_RtgCode = $csvItem.Routing
			$mo.U_Warehouse = $csvItem.Warehouse
			$mo.U_Quantity = $csvItem.Quantity
			$mo.U_Factor = $csvItem.Factor
			$mo.U_RequiredDate = $csvItem.RequiredDate
			switch ($csvItem.Status) {
				"RL" {
					$status = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Released
					break
				}
				"ST" {
					$status = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Started
					break
				}
				"FI" {
					$status = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Finished
					break
				}
				default {
					$status = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Scheduled
					break
				}
			}
			$mo.U_Status = $status
			$mo.CalculateManufacturingTimes($false);
			$count = $mo.RoutingOperations.Count;


			[array] $csvOperations = $manufacturingOrdersOperationsDictionary[$key];	

			foreach ($csvOper in $csvOperations) {	
			
				for ($i = 0; $i -lt $count; $i++) {
					$mo.RoutingOperations.SetCurrentLine($i);
					if ($mo.RoutingOperations.U_OprSequence -eq $csvOper.Sequence) {
						$mo.RoutingOperations.U_Status = $csvOper.Status;
						break;
					}
				}
			}
	
			#Adding Maufacturing Order depends on exists in a database
			$message = $mo.Add();
			if ($message -lt 0) {  
				$err = $pfcCompany.GetLastErrorDescription()
				Throw [System.Exception] ($err);
			}
		}
		Catch {
			$err = $_.Exception.Message;
			$ms = [string]::Format("Error when adding Manufacfturing Order with key: {0}. {1}", $key, $err);
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
		write-host  –backgroundcolor green –foregroundcolor black "Disconnected from the company"
	}
	#endregion
}

