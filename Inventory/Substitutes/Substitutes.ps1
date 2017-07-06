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
$pfcCompany.Databasename = "PFDEMO"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2012"
        
$headerFile = "C:\PS\PF\Inventory\Substitutes\Substitutes.csv"
$nutrientsFile = "C:\PS\PF\Inventory\Substitutes\SubstitutesLines.csv"
$code = $pfcCompany.Connect() 
if($code -eq 1)
{

#Data loading from a csv file
$csvHeaders = Import-Csv -Delimiter ';' -Path $headerFile;
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

 foreach($csvHeader in $csvHeaders) 
 {
 	$rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OSIT"" WHERE ""Code"" = N'{0}'",$csvHeader.Code));
	
    #Creating object
    $md = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Substitutes)
    #Checking if data already exists
	$exists = 0
	if($rs.RecordCount -gt 0)
	{
	    $md.GetByKey($rs.Fields.Item(0).Value);
		$exists = 1
   	}
	else
	{
		$md.Code = $csvHeader.Code;
		$exists = 0
	}
   
	$md.U_Remarks = $csvHeader.Remarks;
	

    #Data loading from a csv file 
    [array]$csvItems = Import-Csv -Delimiter ';' -Path $nutrientsFile | Where-Object {$_.Code -eq $csvHeader.Code}
    
    if($csvItems.count -gt 0)
    {
        #Deleting all existing items
        $count = $md.Lines.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $md.Lines.DelRowAtPos(0);
        }
         
        #Adding the new data       
        foreach($csvItem in $csvItems)
        {
            $md.Lines.U_SItemCode = $csvItem.SItemCode;
            $md.Lines.U_SRevision = $csvItem.SRevision;
            $md.Lines.U_Revision = $csvItem.Revision;
            
            if($csvItem.ValidFrom -ne "")
            {
                $md.Lines.U_ValidFrom = $csvItem.ValidFrom;
            }

            if($csvItem.ValidTO -ne "")
            {
                $md.Lines.U_ValidTo = $csvItem.ValidTO;
            }
            $md.Lines.U_Ratio = $csvItems.Ratio;
            $md.Lines.U_Remarks = $csvItem.Remarks
            $md.Lines.Add();
        }
     }

	$message = 0
    #Adding or updating depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Subsitutes for: {0}", $csvHeader.Code)
        $message = $md.Update()
    }
    else
    {
        [System.String]::Format("Adding Substitutes for: {0}", $csvHeader.Code)
        $message= $md.Add()
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
