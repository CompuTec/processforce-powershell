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
$mds = Import-Csv -Delimiter ';' -Path "C:\PS\PF\QC\Inspectors\InspectorsNCMR.csv";
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($md in $mds) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OINN"" WHERE ""U_InsCode"" = N'{0}'",$md.InspectorCode));
	
    #Creating object
    $mdObj = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::InspectorNcmr)
    #Checking if object already exist
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $mdObj.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
        $mdObj.U_InsCode = $md.InspectorCode;
		$exists = 0
	}
   
   	$mdObj.U_EmployeeId = $md.EmployeeId;
	
	$mdObj.U_InsGrpCode = $md.Group; 
    
    $mdObj.U_WhsCode = $md.Warehouse; 
	
	$mdObj.U_InsRemarks = $md.Remarks;
	
	$message = 0
    #Adding or updating depends if it already exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Inspector: {0}", $md.InspectorCode)
        $message = $mdObj.Update()
    }
    else
    {
        [System.String]::Format("Adding Inspector: {0}", $md.InspectorCode)
        $message= $mdObj.Add()
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
