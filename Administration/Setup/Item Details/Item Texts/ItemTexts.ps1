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
$itemTexts = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Item Details\Texts\ItemTexts.csv";
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($itemText in $itemTexts) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OITX"" WHERE ""U_TxtCode"" = N'{0}'",$itemText.TextCode));
	
    #Creating Item Property object
    $text = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ItemText)
    #Checking that the property already exist
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $text.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
		$text.U_TxtCode = $itemText.TextCode;
		$exists = 0
	}
   
   	$text.U_TxtName = $itemText.TextName;
	
	

	$text.U_GrpCode = $itemText.Group; 
	
	
   
   	if($itemText.ProductionOrders -eq 'Y')
	{
		$text.U_ProdOrders = 'Y'
	}
	else
	{
		$text.U_ProdOrders = 'N'
	}
	
	if($itemText.ShipmentsDocumentation -eq 'Y')
	{
		$text.U_ShipDoc = 'Y'
	}
	else
	{
		$text.U_ShipDoc = 'N'
	}
	
	if($itemText.PickLists -eq 'Y')
	{
		$text.U_PickLists = 'Y'
	}
	else
	{
		$text.U_PickLists = 'N'
	}
	
	if($itemText.MSDS -eq 'Y')
	{
		$text.U_MSDS = 'Y'
	}
	else
	{
		$text.U_MSDS = 'N'
	}
	
	if($itemText.PurchaseOrders -eq 'Y')
	{
		$text.U_PurOrders = 'Y'
	}
	else
	{
		$text.U_PurOrders = 'N'
	}
	
	if($itemText.Returns -eq 'Y')
	{
		$text.U_Returns = 'Y'
	}
	else
	{
		$text.U_Returns = 'N'
	}
	
	if($itemText.Other -eq 'Y')
	{
		$text.U_Other = 'Y'
	}
	else
	{
		$text.U_Other = 'N'
	}
	
	$text.U_Remarks = $itemText.Remarks;
	
	$message = 0
    #Adding or updating Items Texts depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Item Text: {0}", $itemText.TextCode)
        $message = $text.Update()
    }
    else
    {
        [System.String]::Format("Adding Item Text: {0}", $itemText.TextCode)
        $message= $text.Add()
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
