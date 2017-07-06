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
$itmPhrases = Import-Csv -Delimiter ';' -Path "C:\ItemPhrases.csv";
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($itmPhrase in $itmPhrases) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OIPH"" WHERE ""U_PhCode"" = N'{0}'",$itmPhrase.PhraseCode));
	
    #Creating Item Property object
    $phrase = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ItemPhrases)
    #Checking that the property already exist
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $phrase.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
		$phrase.U_PhCode = $itmPhrase.PhraseCode;
		$exists = 0
	}
   
   	$phrase.U_PhName = $itmPhrase.PhraseName;
	
	

	$phrase.U_GrpCode = $itmPhrase.Group; 
	
	
   
   	if($itmPhrase.ProductionOrders -eq 'Y')
	{
		$phrase.U_ProdOrders = 'Y'
	}
	else
	{
		$phrase.U_ProdOrders = 'N'
	}
	
	if($itmPhrase.ShipmentsDocumentation -eq 'Y')
	{
		$phrase.U_ShipDoc = 'Y'
	}
	else
	{
		$phrase.U_ShipDoc = 'N'
	}
	
	if($itmPhrase.PickLists -eq 'Y')
	{
		$phrase.U_PickLists = 'Y'
	}
	else
	{
		$phrase.U_PickLists = 'N'
	}
	
	if($itmPhrase.MSDS -eq 'Y')
	{
		$phrase.U_MSDS = 'Y'
	}
	else
	{
		$phrase.U_MSDS = 'N'
	}
	
	if($itmPhrase.PurchaseOrders -eq 'Y')
	{
		$phrase.U_PurOrders = 'Y'
	}
	else
	{
		$phrase.U_PurOrders = 'N'
	}
	
	if($itmPhrase.Returns -eq 'Y')
	{
		$phrase.U_Returns = 'Y'
	}
	else
	{
		$phrase.U_Returns = 'N'
	}
	
	if($itmPhrase.Other -eq 'Y')
	{
		$phrase.U_Other = 'Y'
	}
	else
	{
		$phrase.U_Other = 'N'
	}
	
	$phrase.U_Remarks = $itmPhrase.Remarks;
	
	$message = 0
    #Adding or updating Items Phrases depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Item Phrase: {0}", $itmPhrase.PhraseCode)
        $message = $phrase.Update()
    }
    else
    {
        [System.String]::Format("Adding Item Phrase: {0}", $itmPhrase.PhraseCode)
        $message= $phrase.Add()
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
