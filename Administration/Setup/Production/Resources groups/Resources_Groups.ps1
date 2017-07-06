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
$csvItems = Import-Csv -Delimiter ';' -Path "C:\Resources_Groups.csv"

 foreach($csvItem in $csvItems) 
 {
    $resGroup = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"ResourceGroup")
    #tCreating Resource Group object    
    $rs = $pfcCompany.SapCompany.GetBusinessObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
    #Checking that the resource already exist
    $val = $rs.DoQuery([System.String]::Format("SELECT Code FROM [@CT_PF_ORGR] WHERE U_RscGrpCode = N'{0}'", $csvItem.ResourceGrpCode))
    $exists = $FALSE
    if ($rs.RecordCount -gt 0)
    {
        $exists = $TRUE
        $code = $rs.Fields.Item("Code").Value.ToString()
        #Loading existed Resource Group Code
        $val = $resGroup.GetByKey($code)
        $val = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rs)
    }
    #Loading data
    $resGroup.U_RscGrpCode = $csvItem.ResourceGrpCode
    $resGroup.U_RscGrpName = $csvItem.GroupName
 
    $message = 0
    
    #Adding or updating Resources Groups depends on exists in the database
    if($exists)
    {
        [System.String]::Format("Updating Resource Group: {0}", $csvItem.ResourceGrpCode)
        $message = $resGroup.Update()
    }
    else
    {
        [System.String]::Format("Adding Resource Group: {0}", $csvItem.ResourceGrpCode)
        $message= $resGroup.Add()
	}
    if($message -lt 0)
    {    
	    $err=$pfcCompany.GetLastErrorDescription()
	    write-host -backgroundcolor red -foregroundcolor white $err
	}    
  }
}
