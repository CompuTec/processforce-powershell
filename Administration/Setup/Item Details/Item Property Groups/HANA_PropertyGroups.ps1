clear
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLServer = "10.0.0.38:30015"
$pfcCompany.SQLUserName = "SYSTEM"
$pfcCompany.SQLPassword = "Ab123456"
$pfcCompany.Databasename = "WB_PF_TEST_DB_PL"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"
       
$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$propGrps = Import-Csv -Delimiter ';' -Path "C:\PropertyGroups.csv";
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($grp in $propGrps) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OIPG"" WHERE ""U_GrpCode"" = N'{0}'",$grp.GroupCode));
	
    #Creating Property Group object
    $propertyGrp = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::"ItemPropertyGroup")
    #Checking that the group already exist
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $propertyGrp.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
		$propertyGrp.U_GrpCode = $grp.GroupCode;
		$exists = 0
	}
   
   	$propertyGrp.U_GrpName = $grp.GroupName;
	$propertyGrp.U_GrpDescription = $grp.Remarks;
	
	
	#Data loading from the csv file - Subgroups for Property Group
    [array]$subGrps = Import-Csv -Delimiter ';' -Path "C:\PropertySubgroups.csv" | Where-Object {$_.GroupCode -eq $grp.GroupCode}
    if($subGrps.count -gt 0)
    {
        #Deleting all exisitng Revisions
        $count = $propertyGrp.Subgroups.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $propertyGrp.Subgroups.DelRowAtPos(0);
        }
        $propertyGrp.Subgroups.SetCurrentLine(0);
         
        #Adding Subgroup
        foreach($subGrp in $subGrps)
        {
			$propertyGrp.Subgroups.U_SubGrpCode = $subGrp.SubgroupCode
			$propertyGrp.Subgroups.U_SubGrpName = $subGrp.SubgroupName
			$propertyGrp.Subgroups.U_SubGrpDescription = $subGrp.SubgroupRemarks
			$propertyGrp.Subgroups.Add();
		}
	}
	$message = 0
    #Adding or updating Property Groups depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Property Group: {0}", $grp.GroupCode)
        $message = $propertyGrp.Update()
    }
    else
    {
        [System.String]::Format("Adding Property Group: {0}", $grp.GroupCode)
        $message= $propertyGrp.Add()
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
