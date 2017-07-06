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
        
$headerFile = "C:\PS\PF\Inventory\Ingredients\Ingredients\Ingredients.csv"
$allergensFile = "C:\PS\PF\Inventory\Ingredients\Ingredients\IngredientAllergens.csv"
$classificationsFile = "C:\PS\PF\Inventory\Ingredients\Ingredients\IngredientClassifications.csv"
$certificatesFile = "C:\PS\PF\Inventory\Ingredients\Ingredients\IngredientClassificationCertificates.csv"
$specificationsFile = "C:\PS\PF\Inventory\Ingredients\Ingredients\IngredientSpecifications.csv"
$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$csvHeaders = Import-Csv -Delimiter ';' -Path $headerFile;
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($csvHeader in $csvHeaders) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OIMD"" WHERE ""U_Code"" = N'{0}'",$csvHeader.Code));
	
    #Creating object
    $md = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::IngredientMasterData)
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
   
   	$md.U_Desc = $csvHeader.Description;
	$md.U_UoM = $csvHeader.UoM;
    $md.U_Category = $csvHeader.Category;
    $md.U_AltCode = $csvHeader.AlternativeCode;
	
	$md.U_Remarks = $csvHeader.Remarks;
	

    #Data loading from a csv file 
    [array]$csvAllergens = Import-Csv -Delimiter ';' -Path $allergensFile | Where-Object {$_.IngredientCode -eq $csvHeader.Code}
    
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
            $md.Allergens.Add();
        }
     }

    #Data loading from a csv file 
    [array]$csvClassifications = Import-Csv -Delimiter ';' -Path $classificationsFile | Where-Object {$_.IngredientCode -eq $csvHeader.Code}
    
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
            [array]$csvCertificates = Import-Csv -Delimiter ';' -Path $certificatesFile | Where-Object {$_.IngredientCode -eq $csvHeader.Code -and $_.ClassificationCode -eq $csvClassification.Code}
            
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
    [array]$csvSpecifications = Import-Csv -Delimiter ';' -Path $specificationsFile | Where-Object {$_.IngredientCode -eq $csvHeader.Code}
    
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
        [System.String]::Format("Updating Ingridnient: {0}", $csvHeader.Code)
        $message = $md.Update()
    }
    else
    {
        [System.String]::Format("Adding Ingridnient: {0}", $csvHeader.Code)
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
