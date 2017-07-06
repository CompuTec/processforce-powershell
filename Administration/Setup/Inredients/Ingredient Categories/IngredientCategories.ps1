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
$csvHeaders = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Inventory\Ingredients\IngredientCategories\IngredientCategories.csv";
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($csvHeader in $csvHeaders) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OIGC"" WHERE ""U_Code"" = N'{0}'",$csvHeader.Code));
	
    #Creating object
    $md = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::IngredientCategory)
    #Checking if data already exists
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $md.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
		$md.U_Code = $csvHeader.Code;
		$exists = 0
	}
   
   	$md.U_Name = $csvHeader.Description;
	
	$md.U_Remarks = $csvHeader.Remarks;
	
	$message = 0
    #Adding or updating depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Ingredient Category: {0}", $csvHeader.Code)
        $message = $md.Update()
    }
    else
    {
        [System.String]::Format("Adding Ingredient Category: {0}", $csvHeader.Code)
        $message= $md.Add()
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
