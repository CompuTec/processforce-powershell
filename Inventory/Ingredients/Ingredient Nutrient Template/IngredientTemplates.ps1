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
        
$headerFile = "C:\PS\PF\Inventory\Ingredients\IngredientTemplates\IngredientTemplates.csv"
$ingredientsFile = "C:\PS\PF\Inventory\Ingredients\IngredientTemplates\IngredientTemplateIngredients.csv"
$nutrientsFile = "C:\PS\PF\Inventory\Ingredients\IngredientTemplates\IngredientTemplateNutrients.csv"
$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$csvHeaders = Import-Csv -Delimiter ';' -Path $headerFile;
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($csvHeader in $csvHeaders) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OIGT"" WHERE ""U_Code"" = N'{0}'",$csvHeader.Code));
	
    #Creating object
    $md = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::IngredientTemplate)
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
   
   	$md.U_Name = $csvHeader.Name;
	$md.U_Remarks = $csvHeader.Remarks;
	

    #Data loading from a csv file 
    [array]$csvIngredients = Import-Csv -Delimiter ';' -Path $ingredientsFile | Where-Object {$_.TemplateCode -eq $csvHeader.Code}
    
    if($csvIngredients.count -gt 0)
    {
        #Deleting all existing items
        $count = $md.Ingredients.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $md.Ingredients.DelRowAtPos(0);
        }
         
        #Adding the new data       
        foreach($csvIngredient in $csvIngredients)
        {
            $md.Ingredients.U_IgdCode = $csvIngredient.Code;
            $md.Ingredients.U_Value = $csvIngredient.Value
            $md.Ingredients.Add();
        }
     }

    #Data loading from a csv file 
    [array]$csvNutrients = Import-Csv -Delimiter ';' -Path $nutrientsFile | Where-Object {$_.TemplateCode -eq $csvHeader.Code}
    
    if($csvNutrients.count -gt 0)
    {
        #Deleting all existing items
        $count = $md.Nutrients.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $md.Nutrients.DelRowAtPos(0);
        }
         
        #Adding the new data       
        foreach($csvNutrient in $csvNutrients)
        {
            $md.Nutrients.U_NutCode = $csvNutrient.Code;
            $md.Nutrients.U_Value = $csvNutrient.Value
            $md.Nutrients.Add();
        }
     }

	$message = 0
    #Adding or updating depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Template: {0}", $csvHeader.Code)
        $message = $md.Update()
    }
    else
    {
        [System.String]::Format("Adding Template: {0}", $csvHeader.Code)
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
