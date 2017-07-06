clear
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
$itmProps = Import-Csv -Delimiter ';' -Path "C:\ItemProperties.csv";
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($prop in $itmProps) 
 {
 	$rs.DoQuery([string]::Format("SELECT Code FROM [@CT_PF_OIPR] WHERE U_PrpCode = N'{0}'",$prop.PropertyCode));
	
    #Creating Item Property object
    $itmProperty = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::"ItemProperty")
    #Checking that the property already exist
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $itmProperty.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
		$itmProperty.U_PrpCode = $prop.PropertyCode;
		$exists = 0
	}
   
   	$itmProperty.U_PrpName = $prop.PropertyName;
	$itmProperty.U_UoM = $prop.UoM;
	
	
	
	if($prop.Group -ne '')
	{
		$rs.DoQuery([string]::Format("SELECT Code FROM [@CT_PF_OIPG] WHERE U_GrpCode = N'{0}'",$prop.Group));
		$itmProperty.U_GrpCode = $rs.Fields.Item(0).Value
	
	
		if($prop.Subgroup -ne '')
		{
			$itmProperty.U_SubGrpLineNo = $prop.Subgroup
		}
	}
	
	if($prop.QualityControlTesting -eq 'Y')
	{
		$itmProperty.U_IsQcTesting = 'Y'
	}
	else
	{
		$itmProperty.U_IsQcTesting = 'N'
	}
   
   	if($prop.ProductionOrders -eq 'Y')
	{
		$itmProperty.U_ProdOrders = 'Y'
	}
	else
	{
		$itmProperty.U_ProdOrders = 'N'
	}
	
	if($prop.ShipmentsDocumentation -eq 'Y')
	{
		$itmProperty.U_ShipDoc = 'Y'
	}
	else
	{
		$itmProperty.U_ShipDoc = 'N'
	}
	
	if($prop.PickLists -eq 'Y')
	{
		$itmProperty.U_PickLists = 'Y'
	}
	else
	{
		$itmProperty.U_PickLists = 'N'
	}
	
	if($prop.MSDS -eq 'Y')
	{
		$itmProperty.U_MSDS = 'Y'
	}
	else
	{
		$itmProperty.U_MSDS = 'N'
	}
	
	if($prop.PurchaseOrders -eq 'Y')
	{
		$itmProperty.U_PurOrders = 'Y'
	}
	else
	{
		$itmProperty.U_PurOrders = 'N'
	}
	
	if($prop.Returns -eq 'Y')
	{
		$itmProperty.U_Returns = 'Y'
	}
	else
	{
		$itmProperty.U_Returns = 'N'
	}
	
	if($prop.Other -eq 'Y')
	{
		$itmProperty.U_Other = 'Y'
	}
	else
	{
		$itmProperty.U_Other = 'N'
	}
	
	#Data loading from the csv file - References for itmes properties
    [array]$references = Import-Csv -Delimiter ';' -Path "c:\ItemPropertiesReferences.csv" | Where-Object {$_.PropertyCode -eq $prop.PropertyCode}
    if($references.count -gt 0)
    {
        #Deleting all exisitng Revisions
        $count = $itmProperty.Words.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $itmProperty.Words.DelRowAtPos(0);
        }
        $itmProperty.Words.SetCurrentLine(0);
         
        #Adding Revisions
        foreach($ref in $references)
        {
			$itmProperty.Words.U_WordCode = $ref.ReferenceCode;
			$itmProperty.Words.Add();
		}
	}
	$message = 0
    #Adding or updating Items Properties depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Item Property: {0}", $prop.PropertyCode)
        $message = $itmProperty.Update()
    }
    else
    {
        [System.String]::Format("Adding Item Property: {0}", $prop.PropertyCode)
        $message= $itmProperty.Add()
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
