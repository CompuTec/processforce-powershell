[CmdletBinding()]
Param(
	[Parameter(Mandatory = $True)]
	[string]$categories,
	[Parameter(Mandatory = $True)]
	[string]$databaseName,
	[Parameter(Mandatory = $True)]	
	[string]$itemFrom,
	[Parameter(Mandatory = $True)]	
	[string]$itemTo,
	[Parameter(Mandatory = $True)]	
	[string]$warehouses,
	[Parameter(Mandatory = $True)]
	[string]$USERNAME,
	[Parameter(Mandatory = $True)]
	[string]$PASSWORD,
	[Parameter(Mandatory = $True)]
	[string]$SQL_SERVER,
	[Parameter(Mandatory = $True)]
	[string]$LICENSE_SERVER,
	[Parameter(Mandatory = $True)]
	[string]$SERVER_TYPE,
	[string]$debugFlag
)
Clear-Host
. $PSScriptRoot\Logger.ps1
$log = [Logger]::new("CostingRollUp");
#### DI API path ####
$assemblyLoadResult = [System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

function WriteLog($msg) {
	$log.WriteLog($msg);
}


if ($debugFlag -eq 'enabled') {
	$msg = [string]::Format("RollUp procedure started with parameters: 
Categories: {0},Database: {1}, itemFrom: {2}, itemTo: {3}, warehouses: {4}, USERNAME: {5}, SQL_SERVER: {6}, LICENSE_SERVER: {7}, SERVER_TYPE: {8} | API Location: {9}",
		$categories, $databaseName, $itemFrom, $itemTo, $warehouses, $USERNAME, $SQL_SERVER, $LICENSE_SERVER, $SERVER_TYPE, $assemblyLoadResult.Location);

	WriteLog $msg;
}

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = $USERNAME
$pfcCompany.Password = $PASSWORD
$pfcCompany.SQLServer = $SQL_SERVER
$pfcCompany.LicenseServer = $LICENSE_SERVER
$pfcCompany.Databasename = $databaseName
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::$SERVER_TYPE

try {
	[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'code')]
	$code = $pfcCompany.Connect()
	if ($debugFlag -eq 'enabled') {
		WriteLog 'Connection succesfull';
	}
}
catch {
	#Show error messages & stop the script
	WriteLog ( "Connection Failure: " + $_.Exception.Message )
	WriteLog ( "LicenseServer:" + $pfcCompany.LicenseServer )
	WriteLog ( "SQLServer:" + $pfcCompany.SQLServer )
	WriteLog ( "DbServerType:" + $pfcCompany.DbServerType )
	WriteLog ( "Databasename" + $pfcCompany.Databasename )
	WriteLog ( "UserName:" + $pfcCompany.UserName )
}

#If company is not connected - stops the script
if (-not $pfcCompany.IsConnected) {
	write-host -backgroundcolor yellow -foregroundcolor black "Company is not connected"
	return ;
}

$costCategories = $categories.split(',');

#$itemFrom = $itemFrom.TrimEnd();
#$itemTo = $itemTo.TrimEnd();


foreach ($costCategory in $costCategories) {
	
	$listBom = New-Object 'System.Collections.Generic.List`1[CompuTec.ProcessForce.API.Costing.Data.RollingUpBillOfMaterials]'
	$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

	$rs.DoQuery(([string]::Format( "SELECT B.""U_ItemCode"", B.""U_Revision"" FROM ""@CT_PF_OBOM"" B WHERE (B.""U_ItemCode"" >= '{0}' OR '{0}' = '*') AND (B.""U_ItemCode"" <= '{1}' OR '{1}' = '*')", $itemFrom, $itemTo)));
	
	$i = 1;
	
	while (!$rs.EoF) {
		$bom = New-Object CompuTec.ProcessForce.API.Costing.Data.RollingUpBillOfMaterials( $rs.Fields.Item(0).Value, $rs.Fields.Item(1).Value);
		$listBom.Add($bom);
		$i = $i + 1;
		$rs.MoveNext();
	};
	
	
	$listRaw = New-Object 'System.Collections.Generic.List`1[String]';

	$listRawQuery = [string]::Format( "SELECT M.""ItemCode"" FROM OITM M 
									LEFT OUTER JOIN ""@CT_PF_OBOM"" B ON M.""ItemCode"" = B.""U_ItemCode"" WHERE ISNULL(B.""U_ItemCode"",'') = ''");
	
	#query for HANA
	if ($SERVER_TYPE -eq 'dst_HANADB') {
		$listRawQuery = [string]::Format( "SELECT M.""ItemCode"" FROM OITM M 
									LEFT OUTER JOIN ""@CT_PF_OBOM"" B ON M.""ItemCode"" = B.""U_ItemCode"" WHERE IFNULL(B.""U_ItemCode"",'') = ''");
	}

	$rs.DoQuery($listRawQuery);
	
	while (!$rs.EoF) {
		
		$str = $rs.Fields.Item(0).Value;
		$listRaw.Add($str);
		$rs.MoveNext();
	}
	
	

	$listWhs = New-Object 'System.Collections.Generic.List`1[String]';
	if ($warehouses -ne '*') {
		$costWarehouses = $warehouses.split(',');
		foreach ($whsCode in $costWarehouses) {
			$listWhs.Add($whsCode)
		}
	}
	else {
		$rs.DoQuery([string]::Format("SELECT ""WhsCode"" FROM OWHS"));
		while (!$rs.EoF) {

			$whsCode = $rs.Fields.Item(0).Value
			$listWhs.Add($whsCode)
			$rs.MoveNext();
		}
	}
	if ($debugFlag -eq 'enabled') {
		WriteLog 'API RollUp procedure started';
	}

	$Result = $pfcCompany.PerformRollUp($listBom, $listRaw, $costCategory, $listWhs);

	if ($debugFlag -eq 'enabled') {
		WriteLog 'API RollUp procedutre completed';
	}
	if ($Result.Success) {
		WriteLog 'Roll Up Completed Successfull';
	}
	else {
		$errorMsg = '';
		foreach ($err in $Result.Errors) {
			
			$errorMsg = $errorMsg + $err.Message;
		}
		WriteLog $errorMsg;
	}
}

$pfcCompany.Disconnect()
