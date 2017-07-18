Clear-Host
#### DI API path ####
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "maciejp"
$pfcCompany.Password = "1234"
$pfcCompany.SQLServer = "10.0.0.202:30015"
$pfcCompany.Databasename = "PFDEMOGB_MACIEJP"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"
         
$code = $pfcCompany.Connect()
if($code -eq 1)
{
#Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
$csvItems = Import-Csv -Delimiter ';' -Path "C:\MP_PS\PF\powershell-scripts\Inventory\Item Management\Batches\Additional Batch Details\BatchDetails.csv"
   
#Checking that Item Details already exist
 foreach($csvItem in $csvItems)
 {
    $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
    $rs.DoQuery([string]::Format( "SELECT ""Code"" FROM ""@CT_PF_OABT""
    WHERE ""U_DistNumber"" = N'{0}' AND ""U_ItemCode"" =  N'{1}'", $csvItem.BatchCode, $csvItem.ItemCode))
    $exists = 0;
    if($rs.RecordCount -gt 0)
    {
        #Creating Additional Batch Details
        $abd = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::AdditionalBatchDetails)
        $abd.GetByKey($rs.Fields.Item(0).Value);
		
		if($csvItem.BatchAttribute1 -gt '')
		{
			$abd.U_MnfSerial = $csvItem.BatchAttribute1;
        }
		
		if($csvItem.BatchAttribute2 -gt '')
		{
			$abd.U_LotNumber = $csvItem.BatchAttribute2;
        }
		
		if($csvItem.SupplierBatch -gt '')
		{
			$abd.U_SupNumber = $csvItem.SupplierBatch;
        }
		
		if($csvItem.Status -eq 'R')
		{
			$abd.U_Status = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Released
        }
		
		if($csvItem.Status -eq 'A')
		{
			$abd.U_Status = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::NotAccesible;
        }
		
		if($csvItem.Status -eq 'L')
		{
			$abd.U_Status = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Locked
        }
		
		if($csvItem.AddminsionDate -gt '')
		{
			$abd.U_AdmDate = $csvItem.AddminsionDate;
        }
		
		if($csvItem.VendorManufacturingDate -gt '')
		{
			$abd.U_VndDate = $csvItem.VendorManufacturingDate;
        }
		
		if($csvItem.ExpiryDate -gt '')
		{
			$abd.U_ExpiryDate = $csvItem.ExpiryDate;
        }
		
		if($csvItem.ExpiryTime -gt '')
		{
			$abd.U_ExpiryTime = $csvItem.ExpiryTime;
        }
		
		if($csvItem.ConsumeByDate -gt '')
		{
			$abd.U_ConsDate = $csvItem.ConsumeByDate;
        }
		
		if($csvItem.LastInspectionDate -gt '')
		{
			$abd.U_LstInDate = $csvItem.LastInspectionDate;
        }
		
		if($csvItem.InspectionDate -gt '')
		{
			$abd.U_InDate = $csvItem.InspectionDate;
        }
		
		if($csvItem.NextInspectionDate -gt '')
		{
			$abd.U_NxtInDate = $csvItem.NextInspectionDate;
		}
		
		if($csvItem.WarningDatePriorExpiry -gt '')
		{
			$abd.U_WExDate = $csvItem.WarningDatePriorExpiry;
		}
		
		if($csvItem.WarningDatePriorConsume -gt '')
		{
			$abd.U_WCoDate = $csvItem.WarningDatePriorConsume;
		}
		
		if($csvItem.Revision -gt '')
		{
			$abd.U_Revision = $csvItem.Revision;
		}

        if($csvItem.RevisionDesc -gt '')
		{
			$abd.U_RevisionDesc = $csvItem.RevisionDesc;
		}
		
        $message = 0



        [System.String]::Format("Updating Item Details: {0}", $csvItem.ItemCode)
        $message = $abd.Update()
       
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
    else
    {

        $msg = [string]::Format("Additional Batch Details don't exists: {0} - {1}. Run Restore Batch Details.",$csvItem.BatchCode, $csvItem.ItemCode);
        write-host -backgroundcolor red -foregroundcolor white $msg
    }
  }
}
else
{
write-host "Failure"
}