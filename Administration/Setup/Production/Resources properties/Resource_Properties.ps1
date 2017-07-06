clear
#### DI API path ####
[void] [Reflection.Assembly]::LoadFrom( "C:\Program Files (x86)\SAP\SAP Business One\AddOns\CT\ProcessForce\CompuTec.Core.DLL")
[void] [Reflection.Assembly]::LoadFrom( "C:\Program Files (x86)\SAP\SAP Business One\AddOns\CT\ProcessForce\CompuTec.ProcessForce.API.DLL")

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
