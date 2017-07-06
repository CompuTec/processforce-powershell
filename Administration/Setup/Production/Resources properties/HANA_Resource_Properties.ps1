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
$resProps = Import-Csv -Delimiter ';' -Path "C:\ResourceProperties.csv"

 foreach($prop in $resProps) 
 {
    #Creating Resource Property object
    $resProperty = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"ResourceProperty")
    #Checking that the property already exist
    $propValue = $resProperty.GetByRscPrpCode($prop.PropertyCode)
    if($retValue -ne 0)
   {
    #Loading data
    $resProperty.U_RscPrpCode = $prop.PropertyCode
	$resProperty.U_RscPrpName = $prop.PropertyName
	$resProperty.U_RscPrpRemarks = $prop.Remarks
   }    
    #Adding or updating Resources Properties depends on exists in the database
    if($propValue -eq 0)
    {
        [System.String]::Format("Updating Resource Property: {0}", $prop.PropertyCode)
        $message = $resProperty.Update()
    }
    else
    {
        [System.String]::Format("Adding Resource Property: {0}", $prop.PropertyCode)
        $message= $resProperty.Add()
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
