clear
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLPassword = "sa"
$pfcCompany.SQLServer = "localhost"
$pfcCompany.SQLUserName = "sa"
$pfcCompany.Databasename = "SBODemoPL"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]"dst_MSSQL2008"
       
$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$oprProps = Import-Csv -Delimiter ';' -Path "C:\OperationProperties.csv"

 foreach($prop in $oprProps) 
 {
    #Creating Operation Property object
    $oprProperty = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"OperationProperty")
    #Checking that the property already exist
    $propValue = $oprProperty.GetByOprPrpCode($prop.PropertyCode)
    if($retValue -ne 0)
   {
    #Loading data
    $oprProperty.U_OprPrpCode = $prop.PropertyCode
	$oprProperty.U_OprPrpName = $prop.PropertyName
	$oprProperty.U_OprPrpRemarks = $prop.Remarks
   }    
    #Adding or updating Operation Properties depends on exists in the database
    if($propValue -eq 0)
    {
        [System.String]::Format("Updating Operation Property: {0}", $prop.PropertyCode)
        $message = $oprProperty.Update()
    }
    else
    {
        [System.String]::Format("Adding Operation Property: {0}", $prop.PropertyCode)
        $message= $oprProperty.Add()
	}
            
    if($message -lt 0)
    {    
	    $err=$pfcCompany.GetLastErrorDescription()
	    write-host -backgroundcolor red -foregroundcolor white $err
	}    
  }
}
else
{
write-host "Failure"
}
