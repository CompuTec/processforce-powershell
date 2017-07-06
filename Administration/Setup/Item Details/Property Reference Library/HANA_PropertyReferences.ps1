clear
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLServer = "10.0.0.38:30015"
$pfcCompany.SQLUserName = "SYSTEM"
$pfcCompany.SQLPassword = "password"
$pfcCompany.Databasename = "PFDEMO"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"
    
$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$itmPropRefs = Import-Csv -Delimiter ';' -Path "C:\PropertyReferences.csv";
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($ref in $itmPropRefs) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OWRD"" WHERE ""U_WordCode"" = N'{0}'",$ref.ReferenceCode));
	
    #Creating Property Reference object
    $propertyReference = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::"Wordbook")
    #Checking that the property already exist
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $propertyReference.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
		$propertyReference.U_WordCode = $ref.ReferenceCode;
		$exists = 0
	}
   
   	$propertyReference.U_WordName = $ref.ReferenceName;
	$propertyReference.U_WordRemarks = $ref.Remarks;
	
	if($ref.ItemProperty -eq 'Y')
	{
		$propertyReference.U_ItemProp = 'Y'
	}
	else
	{
		$propertyReference.U_ItemProp = 'N'
	}
   
   	if($ref.TestProperty -eq 'Y')
	{
		$propertyReference.U_TestProp = 'Y'
	}
	else
	{
		$propertyReference.U_TestProp = 'N'
	}
	
	if($ref.ResourceProperty -eq 'Y')
	{
		$propertyReference.U_RscProp = 'Y'
	}
	else
	{
		$propertyReference.U_RscProp = 'N'
	}
	
	if($ref.OperationProperty -eq 'Y')
	{
		$propertyReference.U_OperProp = 'Y'
	}
	else
	{
		$propertyReference.U_OperProp = 'N'
	}
	
	if($ref.OperationIOProperty -eq 'Y')
	{
		$propertyReference.U_OperIoProp = 'Y'
	}
	else
	{
		$propertyReference.U_OperIoProp = 'N'
	}
	
	$message = 0
    #Adding or updating Property Reference Library depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Property Reference: {0}", $ref.ReferenceCode)
        $message = $propertyReference.Update()
    }
    else
    {
        [System.String]::Format("Adding Property Reference: {0}", $ref.ReferenceCode)
        $message= $propertyReference.Add()
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
