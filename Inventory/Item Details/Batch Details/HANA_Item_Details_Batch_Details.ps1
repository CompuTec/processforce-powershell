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
#Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
$csvItems = Import-Csv -Delimiter ';' -Path "c:\MP_PS\ItemDetailsBatchDetails.csv"
 
#Checking that Item Details already exist 
 foreach($csvItem in $csvItems) 
 {
  
        $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
     
        $rs.DoQuery([string]::Format( "SELECT T0.""ItemCode"" FROM OITM T0
            LEFT OUTER JOIN ""@CT_PF_OIDT"" T1 ON T0.""ItemCode"" = T1.""U_ItemCode""
            WHERE
            T1.""U_ItemCode"" = N'{0}'", $csvItem.ItemCode))
        $exists = 0;
        if($rs.RecordCount -gt 0)
        {
            $exists = 1
        }
  
    #Creating Item Details 
    $idt = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"ItemDetails")
     
    #Restoring Item Costs and setting Inherit Batch/Serial to 'Yes'
    if($exists -eq 1)
    {
        $idt.GetByItemCode($csvItem.ItemCode)
    }
    else
    {
        $idt.U_ItemCode = $csvItem.ItemCode;
        $idt.CFG_RestoreItemCosting = "Y";
    }
     
	if( $csvItem.BatchInherit -eq 0)
	{
		$idt.U_InheritBatch = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
		$idt.U_BtchTmpl = $csvItem.BatchTemplate;
	}
	else 
	{
		$idt.U_InheritBatch = 1
	}
	
	if( $csvItem.SerialIncherit -eq 0)
	{
    	$idt.U_InheritSerial = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
		$idt.U_SrlTmpl = $csvItem.SerialTemplate;
	}
	else
	{
		$idt.U_InheritSerial = 1;
	}
	
	if( $csvItem.ExpiryInherit -eq 0)
	{
		$idt.U_Inherit = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
		if ( $csvItem.Expiry -eq 1 )
		{
			$idt.U_Expiry	= [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
			$idt.U_ExpWarn = $csvItem.ExpiryWarning;
		}
		
		if ( $csvItem.Consume -eq 1)
		{
			$idt.U_Consume = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
			$idt.U_ConsWarn = $csvItem.ConsWarn;
		}
		if ( $csvItem.ShelfLife -ne "")
		{
			$idt.U_ShelfTime = $csvItem.ShelfLife;
		}
		if( $csvItem.InspectionInterval -ne "")
		{
			$idt.U_InspDays = $csvItem.InspectionInterval;
		}
		
	}
	else
	{
		$idt.U_Inherit = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
	}

	
	if ( $csvItem.BatchQueue -eq "F")
	{
		$idt.U_BatchQueue = [CompuTec.ProcessForce.API.Enumerators.BatchQueueType]::FIFO;
	}
	
	if ( $csv.BatchQueue -eq "E")
	{
		$idt.U_BatchQueue = [CompuTec.ProcessForce.API.Enumerators.BatchQueueType]::FEFO;
	}
		
	if ( $csv.BatchQueue -eq "M")
	{
		$idt.U_BatchQueue = [CompuTec.ProcessForce.API.Enumerators.BatchQueueType]::FMFO;
	}
	
  
    $message = 0
     
    #Adding or updating ItemDetails depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Item Details: {0}", $csvItem.ItemCode)
        $message = $idt.Update()
    }
    else
    {
        [System.String]::Format("Adding Item Details: {0}", $csvItem.ItemCode)
       $message= $idt.Add()
    }
     
    if($message -lt 0)
    {    
        $err=$pfcCompany.GetLastErrorDescription()
        write-host -backgroundcolor red -foregroundcolor white $err
    } 
    else
    {
        write-host "Success"
    }   
  }
}
else
{
write-host "Failure"
}