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
$pfcCompany.Databasename = "PFDemo915"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2012"
        
$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$Grps = Import-Csv -Delimiter ';' -Path "C:\PS\PF\QC\Complaints\ComplaintGroups.csv";
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($grp in $Grps) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OCGR"" WHERE ""U_GrpCode"" = N'{0}'",$grp.ComplaintGroupCode));
	
    #Creating Group object
    $group = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ComplaintGroup)
    #Checking that the group already exist
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $group.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
		$group.U_GrpCode = $grp.ComplaintGroupCode;
		$exists = 0
	}
   
   	$group.U_GrpName = $grp.GroupName;


   if($grp.Customer -eq 'Y')
	{
		$group.U_Customer = 'Y'
	}
    else
	{
		$group.U_Customer = 'N'
	}
 
 
    if($grp.Supplier -eq 'Y')
	{
		$group.U_Supplier = 'Y'
	}
    else
	{
		$group.U_Supplier = 'N'
	}
	
    if($grp.Internal -eq 'Y')
	{
		$group.U_Internal = 'Y'
	}
    else
	{
		$group.U_Internal = 'N'
	}



	$group.U_Remarks = $grp.Remarks;
	
	
	
	$message = 0
    #Adding or updating Group depends if it already exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Group: {0}", $grp.ComplaintGroupCode)
        $message = $group.Update()
    }
    else
    {
        [System.String]::Format("Adding Group: {0}", $grp.ComplaintGroupCode)
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
