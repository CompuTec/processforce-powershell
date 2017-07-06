clear
#### DI API path ####
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLPassword = "sa"
$pfcCompany.SQLServer = "localhost"
$pfcCompany.SQLUserName = "sa"
$pfcCompany.Databasename = "PFDemo"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2012"
        
$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$itmClass = Import-Csv -Delimiter ';' -Path "C:\ItemClassifications.csv";
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($itmC in $itmClass) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OICL"" WHERE ""U_ClsCode"" = N'{0}'",$itmC.ClassificationCode));
	
    #Creating Item Property object
    $itmClassificaiton = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ItemClassification)
    #Checking that the property already exist
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $itmClassificaiton.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
		$itmClassificaiton.U_ClsCode = $itmC.ClassificationCode;
		$exists = 0
	}
   
   	$itmClassificaiton.U_ClsName = $itmC.ClassificationName;
	
	

	$itmClassificaiton.U_GrpCode = $itmC.Group; 
	
	
   
   	if($itmC.ProductionOrders -eq 'Y')
	{
		$itmClassificaiton.U_ProdOrders = 'Y'
	}
	else
	{
		$itmClassificaiton.U_ProdOrders = 'N'
	}
	
	if($itmC.ShipmentsDocumentation -eq 'Y')
	{
		$itmClassificaiton.U_ShipDoc = 'Y'
	}
	else
	{
		$itmClassificaiton.U_ShipDoc = 'N'
	}
	
	if($itmC.PickLists -eq 'Y')
	{
		$itmClassificaiton.U_PickLists = 'Y'
	}
	else
	{
		$itmClassificaiton.U_PickLists = 'N'
	}
	
	if($itmC.MSDS -eq 'Y')
	{
		$itmClassificaiton.U_MSDS = 'Y'
	}
	else
	{
		$itmClassificaiton.U_MSDS = 'N'
	}
	
	if($itmC.PurchaseOrders -eq 'Y')
	{
		$itmClassificaiton.U_PurOrders = 'Y'
	}
	else
	{
		$itmClassificaiton.U_PurOrders = 'N'
	}
	
	if($itmC.Returns -eq 'Y')
	{
		$itmClassificaiton.U_Returns = 'Y'
	}
	else
	{
		$itmClassificaiton.U_Returns = 'N'
	}
	
	if($itmC.Other -eq 'Y')
	{
		$itmClassificaiton.U_Other = 'Y'
	}
	else
	{
		$itmClassificaiton.U_Other = 'N'
	}
	
	$itmClassificaiton.U_Remarks = $itmC.Remarks;
	
	$message = 0
    #Adding or updating Items Properties depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Item Classification: {0}", $itmC.ClassificationCode)
        $message = $itmClassificaiton.Update()
    }
    else
    {
        [System.String]::Format("Adding Item Classification: {0}", $itmC.ClassificationCode)
        $message= $itmClassificaiton.Add()
	}
            
    if($message -lt 0)
    {    
	    $err=$pfcCompany.GetLastErrorDescription()
	    write-host -backgroundcolor red -foregroundcolor white $err
	}
	else
	{
		Write-Host -BackgroundColor Blue -ForegroundColor White "Success"
	}
  }
}
else
{
write-host "Failure"
}
