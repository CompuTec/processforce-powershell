clear
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLPassword = "password"
$pfcCompany.SQLServer = "localhost"
$pfcCompany.SQLUserName = "sa"
$pfcCompany.Databasename = "PFDemo"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2012"
       
$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$testProperty = Import-Csv -Delimiter ';' -Path "C:\TestProperties.csv";
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($prop in $testProperty) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OTPR"" WHERE ""U_TestPrpCode"" = N'{0}'",$prop.PropertyCode));
	
    #Creating Item Property object
    $testProperty = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::"TestProperty")
    #Checking that the property already exist
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $testProperty.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
		$testProperty.U_TestPrpCode = $prop.PropertyCode;
		$exists = 0
	}
   
   	$testProperty.U_TestPrpName = $prop.PropertyName;
	$testProperty.U_TestPrpGrpCode = $prop.Group;
	$testProperty.U_TestPrpRemarks = $prop.Remarks;
	
	if($prop.Group -ne '')
	{
		$testProperty.U_TestPrpGrpCode = $prop.Group 
	}
	
	
	
	#Data loading from the csv file - References for test properties
    [array]$references = Import-Csv -Delimiter ';' -Path "C:\TestPropertiesReferences.csv" | Where-Object {$_.PropertyCode -eq $prop.PropertyCode}
    if($references.count -gt 0)
    {
        #Deleting all exisitng Revisions
        $count = $testProperty.Words.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $testProperty.Words.DelRowAtPos(0);
        }
        $testProperty.Words.SetCurrentLine(0);
         
        #Adding Revisions
        foreach($ref in $references)
        {
			$testProperty.Words.U_WordCode = $ref.ReferenceCode;
			$testProperty.Words.Add();
		}
	}
	$message = 0
    #Adding or updating Test Properties depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Item Property: {0}", $prop.PropertyCode)
        $message = $testProperty.Update()
    }
    else
    {
        [System.String]::Format("Adding Item Property: {0}", $prop.PropertyCode)
        $message= $testProperty.Add()
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
