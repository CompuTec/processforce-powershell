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
$Grps = Import-Csv -Delimiter ';' -Path "C:\PS\PF\QC\Reasons\ReasonGroups.csv";
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($grp in $Grps) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_ORSG"" WHERE ""U_RsnGrpCode"" = N'{0}'",$grp.GroupCode));
	
    #Creating Group object
    $group = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ReasonGroup)
    #Checking that the group already exist
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $group.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
		$group.U_RsnGrpCode = $grp.GroupCode;
		$exists = 0
	}
   
   	$group.U_RsnGrpName = $grp.GroupName;
	$group.U_RsnGrpRemarks = $grp.Remarks;
	
	
	
	$message = 0
    #Adding or updating Group depends if it already exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Group: {0}", $grp.GroupCode)
        $message = $group.Update()
    }
    else
    {
        [System.String]::Format("Adding Group: {0}", $grp.GroupCode)
        $message= $group.Add()
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
