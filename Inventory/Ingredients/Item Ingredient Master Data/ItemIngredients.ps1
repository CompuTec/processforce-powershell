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
        
$headerFile = "C:\PS\PF\Inventory\Ingredients\ItemIngredients\ItemIngredients.csv"
$allergensFile = "C:\PS\PF\Inventory\Ingredients\ItemIngredients\ItemIngredientAllergens.csv"
$classificationsFile = "C:\PS\PF\Inventory\Ingredients\ItemIngredients\ItemIngredientClassifications.csv"
$certificatesFile = "C:\PS\PF\Inventory\Ingredients\ItemIngredients\ItemIngredientClassificationCertificates.csv"
$specificationsFile = "C:\PS\PF\Inventory\Ingredients\ItemIngredients\ItemIngredientSpecifications.csv"
$ingredientsFile = "C:\PS\PF\Inventory\Ingredients\ItemIngredients\ItemIngredientIngredients.csv"
$nutrientsFile = "C:\PS\PF\Inventory\Ingredients\ItemIngredients\ItemIngredientNutrients.csv"

$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$csvHeaders = Import-Csv -Delimiter ';' -Path $headerFile;
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($csvHeader in $csvHeaders) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OIID"" WHERE ""U_ItemCode"" = N'{0}' AND ""U_Revision"" = '{1}'",$csvHeader.ItemCode,$csvHeader.Revision));
	
    #Creating object
    $md = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ItemIngredientData)
    #Checking if data already exists
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
        $md.GetByCodeAndRevision($csvHeader.ItemCode,$csvHeader.Revision);
		$exists = 1
   	}
	else
	{
		$md.U_ItemCode = $csvHeader.ItemCode;
        $md.U_Revision = $csvHeader.Revision;
		$exists = 0
	}
   

   	$md.U_Quantity = $csvHeader.Quantity;
	$md.U_AltUoM = $csvHeader.AltUoM;
    $md.U_Category = $csvHeader.Category;
    $md.U_AltCode = $csvHeader.AlternativeCode;
	$md.U_KCal = $csvHeader.EnergyDesity;
    $md.U_IntakeCode = $csvHeader.DailyIntakeCode;
	$md.U_Remarks = $csvHeader.Remarks;
	

    #Data loading from a csv file 
    [array]$csvIngredients = Import-Csv -Delimiter ';' -Path $ingredientsFile | Where-Object {$_.ItemCode -eq $csvHeader.ItemCode -and $_.Revision -eq $csvHeader.Revision}
    
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
            if($csvIngredient.Condition -eq 'gt')
            {
                $md.Ingredients.U_Condition = [CompuTec.ProcessForce.API.Documents.Ingredients.IngredientCondition]::GreaterThan; #EQ - equal, GT - greater then, LT - less then
            }
            elseif($csvIngredient.Condition -eq 'lt')
            {
                $md.Ingredients.U_Condition = [CompuTec.ProcessForce.API.Documents.Ingredients.IngredientCondition]::LessThan; #EQ - equal, GT - greater then, LT - less then
            }
            else
            {
                $md.Ingredients.U_Condition = [CompuTec.ProcessForce.API.Documents.Ingredients.IngredientCondition]::Equal; #EQ - equal, GT - greater then, LT - less then
            }
            $md.Ingredients.U_Value = $csvIngredient.Value
            $md.Ingredients.Add();
        }
     }


     #Data loading from a csv file 
    [array]$csvNutrients = Import-Csv -Delimiter ';' -Path $nutrientsFile | Where-Object {$_.ItemCode -eq $csvHeader.ItemCode -and $_.Revision -eq $csvHeader.Revision}
    
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



    #Data loading from a csv file 
    [array]$csvAllergens = Import-Csv -Delimiter ';' -Path $allergensFile | Where-Object {$_.ItemCode -eq $csvHeader.ItemCode -and $_.Revision -eq $csvHeader.Revision}
    
    if($csvAllergens.count -gt 0)
    {
        #Deleting all existing items
        $count = $md.Allergens.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $md.Allergens.DelRowAtPos(0);
        }
         
        #Adding the new data       
        foreach($csvAllergen in $csvAllergens)
        {
            $md.Allergens.U_AlgCode = $csvAllergen.Code;
            if($csvAllergen.CrossContamination -eq 'Y')
            {
                $md.Allergens.U_CrossCtmn = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
            }
            else
            {
                $md.Allergens.U_CrossCtmn = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
            }
            $md.Allergens.Add();
        }
     }

    #Data loading from a csv file 
    [array]$csvClassifications = Import-Csv -Delimiter ';' -Path $classificationsFile | Where-Object {$_.ItemCode -eq $csvHeader.ItemCode -and $_.Revision -eq $csvHeader.Revision}
    
    if($csvClassifications.count -gt 0)
    {
        #Deleting all existing items
        $count = $md.Classifications.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $md.Classifications.DelRowAtPos(0);
        }
         
        $count = $md.Certificates.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $md.Certificates.DelRowAtPos(0);
        }
        
        #Adding the new data       
        foreach($csvClassification in $csvClassifications)
        {
            $md.Classifications.U_ClassCode = $csvClassification.Code;
            $md.Classifications.Add();

            
            #Data loading from a csv file 
            [array]$csvCertificates = Import-Csv -Delimiter ';' -Path $certificatesFile | Where-Object {$_.ItemCode -eq $csvHeader.ItemCode -and $_.Revision -eq $csvHeader.Revision -and  $_.ClassificationCode -eq $csvClassification.Code}
            
            if($csvCertificates.Count -gt 0){
                
                #Adding the new data       
                foreach($csvCertificate in $csvCertificates)
                {
                    $md.Certificates.U_ClassCode = $csvClassification.Code;
                    $md.Certificates.U_BPCode = $csvCertificate.BusinessPartnerCode;
                    $md.Certificates.U_CertNum = $csvCertificate.CertificateNumber;
                    $md.Certificates.U_CertDate = $csvCertificate.CertificateDate;
                    $md.Certificates.U_Status = $csvCertificate.Status; #NA - not approved, P - pending, A - approved
                    $md.Certificates.U_StatDate = $csvCertificate.StatusDate;
                    $md.Certificates.U_Attachment = $csvCertificate.Attachment;
                    $md.Certificates.U_Remarks = $csvCertificate.Remarks;
                    $md.Certificates.Add();
                }
            }

        }


     }

     #Data loading from a csv file 
    [array]$csvSpecifications = Import-Csv -Delimiter ';' -Path $specificationsFile | Where-Object {$_.ItemCode -eq $csvHeader.ItemCode -and $_.Revision -eq $csvHeader.Revision}
    
    if($csvSpecifications.count -gt 0)
    {
        #Deleting all existing items
        $count = $md.Specifications.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $md.Specifications.DelRowAtPos(0);
        }

        #Adding the new data       
        foreach($csvSpecification in $csvSpecifications)
        {
            $md.Specifications.U_BPCode = $csvSpecification.BusinessPartnerCode;
            $md.Specifications.U_SpecNum = $csvSpecification.SpecificationNumber
            $md.Specifications.U_SpecDate = $csvSpecification.SpecificationDate;
            $md.Specifications.U_Status = $csvSpecification.Status; #NA - not approved, P - pending, A - approved
            $md.Specifications.U_StatDate = $csvSpecification.StatusDate;
            $md.Specifications.U_Attachment = $csvSpecification.Attachment;
            $md.Specifications.U_Remarks = $csvSpecification.Remarks;
            $md.Specifications.Add();
        }


    }
	$message = 0
    #Adding or updating depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Item Ingridnient: {0}", $csvHeader.ItemCode)
        $message = $md.Update()
    }
    else
    {
        [System.String]::Format("Adding Item Ingridnient: {0}", $csvHeader.ItemCode)
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
