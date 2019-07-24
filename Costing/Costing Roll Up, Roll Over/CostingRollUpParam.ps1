﻿[CmdletBinding()]
Param(
 	[Parameter(Mandatory=$True)]
   	[string]$categories,
	[Parameter(Mandatory=$True)]
	[string]$databaseName,
    [Parameter(Mandatory=$True)]	
    [string]$itemFrom,
	[Parameter(Mandatory=$True)]	
    [string]$itemTo,
    [Parameter(Mandatory=$True)]	
    [string]$warehouses,
    [Parameter(Mandatory=$True)]
	[string]$USERNAME,
    [Parameter(Mandatory=$True)]
	[string]$PASSWORD,
    [Parameter(Mandatory=$True)]
	[string]$SQL_SERVER,
    [Parameter(Mandatory=$True)]
	[string]$SQL_USERNAME,
    [Parameter(Mandatory=$True)]
	[string]$SQL_PASSWORD,
    [Parameter(Mandatory=$True)]
	[string]$SERVER_TYPE,
	[string]$debugFlag
)
Clear-Host
#### DI API path ####
$assemblyLoadResult = [System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

if($debugFlag -eq 'enabled')
{
$msg = [string]::Format("RollUP procedure started with parameters: 
Categories: {0},Database: {1}, itemFrom: {2}, itemTo: {3}, warehouses: {4}, USERNAME: {5}, SQL_SERVER: {6}, SQL_USERNAME: {7}, SERVER_TYPE: {8} | API Location: {9}",
$categories,$databaseName,$itemFrom,$itemTo,$warehouses,$USERNAME,$SQL_SERVER,$SQL_USERNAME,$SERVER_TYPE,$assemblyLoadResult.Location);

Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType Information -EventId 1 -Message $msg;
}

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = $USERNAME
$pfcCompany.Password = $PASSWORD
$pfcCompany.SQLPassword = $SQL_PASSWORD
$pfcCompany.SQLServer = $SQL_SERVER
$pfcCompany.SQLUserName = $SQL_USERNAME
$pfcCompany.Databasename = $databaseName
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::$SERVER_TYPE


$code = $pfcCompany.Connect()
if($code -eq 1)
#if($pfcCompany.IsConnected -eq 1)
{
if($debugFlag -eq 'enabled')
{
    Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType SuccessAudit -EventId 1 -Message 'Connection succesfull';
}
$costCategories = $categories.split(',');

#$itemFrom = $itemFrom.TrimEnd();
#$itemTo = $itemTo.TrimEnd();


foreach ($costCategory in $costCategories)
{
	
	$listBom = New-Object 'System.Collections.Generic.List`1[CompuTec.ProcessForce.API.Costing.Data.RollingUpBillOfMaterials]'
	$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

	$rs.DoQuery(([string]::Format( "SELECT B.""U_ItemCode"", B.""U_Revision"" FROM ""@CT_PF_OBOM"" B WHERE (B.""U_ItemCode"" >= '{0}' OR '{0}' = '*') AND (B.""U_ItemCode"" <= '{1}' OR '{1}' = '*')",$itemFrom,$itemTo)));
	
	$i = 1;
	
	while(!$rs.EoF)
	{
		$bom = New-Object CompuTec.ProcessForce.API.Costing.Data.RollingUpBillOfMaterials( $rs.Fields.Item(0).Value, $rs.Fields.Item(1).Value);
		$listBom.Add($bom);
		$i = $i + 1;
		$rs.MoveNext();
	};
	
	
	$listRaw = New-Object 'System.Collections.Generic.List`1[String]';

    $listRawQuery = [string]::Format( "SELECT M.""ItemCode"" FROM OITM M 
									LEFT OUTER JOIN ""@CT_PF_OBOM"" B ON M.""ItemCode"" = B.""U_ItemCode"" WHERE ISNULL(B.""U_ItemCode"",'') = ''");
	
    #query for HANA
    if($SERVER_TYPE -eq 'dst_HANADB'){
        $listRawQuery = [string]::Format( "SELECT M.""ItemCode"" FROM OITM M 
									LEFT OUTER JOIN ""@CT_PF_OBOM"" B ON M.""ItemCode"" = B.""U_ItemCode"" WHERE IFNULL(B.""U_ItemCode"",'') = ''");
    }

    $rs.DoQuery($listRawQuery);
	
	while(!$rs.EoF)
	{
		
    	$str = $rs.Fields.Item(0).Value;
		$listRaw.Add($str);
		$rs.MoveNext();
	}
	
	

	$listWhs = New-Object 'System.Collections.Generic.List`1[String]';
	if($warehouses -ne '*')
	{
		$costWarehouses = $warehouses.split(',');
		foreach($whsCode in $costWarehouses)
		{
			$listWhs.Add($whsCode)
		}
	}
	else
	{
		$rs.DoQuery([string]::Format("SELECT ""WhsCode"" FROM OWHS"));
		while(!$rs.EoF)
		{

			$whsCode = $rs.Fields.Item(0).Value
			$listWhs.Add($whsCode)
			$rs.MoveNext();
		}
	}
    if($debugFlag -eq 'enabled')
    {
	    Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType Information -EventId 1 -Message 'API RollUp procedure started';
    }

	$Result = $pfcCompany.PerformRollUp($listBom, $listRaw, $costCategory, $listWhs);

    if($debugFlag -eq 'enabled')
    {
	    Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType Information -EventId 1 -Message 'API RollUp procedutre completed';
    }
	if($Result.Success)
	{
		Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType SuccessAudit -EventId 1 -Message 'Roll Up Completed Successfull';
	}
	else 
	{
		$errorMsg = '';
		foreach($err in $Result.Errors)
		{
			
			$errorMsg = $errorMsg + $err.Message;
		}
		Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType Error -EventId 1 -Message $errorMsg;
	}
}

$pfcCompany.Disconnect()
}
else
{
	Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType Error -EventId 1 -Message 'Connection Failure';
}