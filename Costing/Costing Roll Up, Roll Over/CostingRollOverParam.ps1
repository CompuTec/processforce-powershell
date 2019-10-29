[CmdletBinding()]
Param(
 	[Parameter(Mandatory=$True)]
   	[string]$costCategoryFrom,
	[Parameter(Mandatory=$True)]
   	[string]$costCategoryTo,
	[Parameter(Mandatory=$True)]
	[string]$databaseName,
	[string]$itemFrom,
	[string]$itemTo,
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
clear
#### DI API path ####
$assemblyLoadResult = [System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")


if($debugFlag -eq 'enabled')
{
$msg = [string]::Format("RollOver procedure started with parameters: 
Category From: {0}, Category To: {1}, Database: {2}, itemFrom: {3}, itemTo: {4}, warehouses: {5}, USERNAME: {6}, SQL_SERVER: {7}, SQL_USERNAME: {8}, SERVER_TYPE: {9} | API Location: {10}",
$costCategoryFrom,$costCategoryTo,$databaseName,$itemFrom,$itemTo,$warehouses,$USERNAME,$SQL_SERVER,$SQL_USERNAME,$SERVER_TYPE,$assemblyLoadResult.Location);

Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType Information -EventId 1 -Message $msg;
}


#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = $USERNAME
$pfcCompany.Password = $PASSWORD
$pfcCompany.SQLServer = $SQL_SERVER
$pfcCompany.SQLUserName = $SQL_USERNAME
$pfcCompany.SQLPassword = $SQL_PASSWORD
$pfcCompany.Databasename = $databaseName
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::$SERVER_TYPE

$code = $pfcCompany.Connect()
if($code -eq 1)
{
    if($debugFlag -eq 'enabled')
    {
        Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType SuccessAudit -EventId 1 -Message 'Connection succesfull';
    }
	$listItem = New-Object 'System.Collections.Generic.List`1[CompuTec.ProcessForce.API.Costing.Data.RollOverItem]'
	$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
	$rs1 = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
	$rs.DoQuery([string]::Format( "SELECT M.""ItemCode"" FROM OITM M WHERE (M.""ItemCode"" >= '{0}' OR '{0}' = '*') AND (M.""ItemCode"" <= '{1}' OR '{1}' = '*')",$itemFrom,$itemTo));
	
	$i = 1;
	
	while(!$rs.EoF)
	{
		
		$item = [CompuTec.ProcessForce.API.Costing.Data.RollOverItem]::GetRollOverItem($pfcCompany.Token,$rs.Fields.Item(0).Value);
		
		$listItem.Add($item);
		$i = $i + 1;
		$rs.MoveNext();
	};
	
	$listWhs = New-Object 'System.Collections.Generic.List`1[String]';
	$rs.DoQuery([string]::Format("SELECT ""WhsCode"" FROM OWHS"));
	
	
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
	    Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType Information -EventId 1 -Message 'API RollOver procedure started';
    }

	$Result = $pfcCompany.PerformRollOver($listItem, $listWhs, $costCategoryFrom, $costCategoryTo);
	
    if($debugFlag -eq 'enabled')
    {
	    Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType Information -EventId 1 -Message 'API RollOver procedure completed';
    }


    if($Result.Success)
	{
		Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType SuccessAudit -EventId 1 -Message 'Roll Over Completed Successfull';
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
else
{
	Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType Error -EventId 1 -Message 'Connection Failure';
}