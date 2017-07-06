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
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2012
        
$code = $pfcCompany.Connect()
if($code -eq 1)
{
#Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
$csvItems = Import-Csv -Delimiter ';' -Path "C:\ItemDetailsBatchDetails.csv"
 
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

	
	if( $csvItem.InheritBatchQueue -eq 0)
	{
		$idt.U_InheritQueue = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
		$idt.U_BtchTmpl = $csvItem.BatchTemplate;
		
		
		if ( $csvItem.BatchQueue -eq "F")
		{
			$idt.U_BatchQueue = [CompuTec.ProcessForce.API.Enumerators.BatchQueueType]::FIFO;
		}
		
		if ( $csvItem.BatchQueue -eq "E")
		{
			$idt.U_BatchQueue = [CompuTec.ProcessForce.API.Enumerators.BatchQueueType]::FEFO;
		}
			
		if ( $csvItem.BatchQueue -eq "M")
		{
			$idt.U_BatchQueue = [CompuTec.ProcessForce.API.Enumerators.BatchQueueType]::FMFO;
		}
	}
	else 
	{
		$idt.U_InheritQueue = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
	}
	
    $idt.U_ExpTmpl= "";

    if ( $csvItem.ExpTyp -eq "C")
    { $idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::CreateDate }

    if ( $csvItem.ExpTyp -eq "N" -or $csvItem.ExpTyp -eq "" )
    {
        $idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::CurrentDate 
    }

    if ( $csvItem.ExpTyp -eq "E")
    { $idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::EndDate }

    if ( $csvItem.ExpTyp -eq "Q")
    { 
        $idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::Query 
        $idt.U_ExpTmpl=$csvItem.ExpTempl;
    }

    if ( $csvItem.ExpTyp -eq "R")
    { $idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::RequiredDate }

    if ( $csvItem.ExpTyp -eq "S")
    { $idt.U_ExpTyp = [CompuTec.ProcessForce.API.Enumerators.ExpiryDateEvaluation]::StartDate }
    
    if( $csvItem.InheritStatus -eq 0)
    {
        $idt.U_InheritStatus = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
        if( $csvItem.U_SapDfBS -eq 'R')
        {
            $idt.U_SapDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Released;
        }
        if( $csvItem.U_SapDfBS -eq 'L')
        {
            $idt.U_SapDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Locked;
        }
        if( $csvItem.U_SapDfBS -eq 'A')
        {
            $idt.U_SapDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::NotAccesible;
        }

        if( $csvItem.U_SapDfQCS -eq 'F')
        {
            $idt.U_SapDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Failed
        }
        if( $csvItem.U_SapDfQCS -eq 'H')
        {
            $idt.U_SapDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::OnHold
        }
        if( $csvItem.U_SapDfQCS -eq 'I')
        {
            $idt.U_SapDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Inspection
        }
        if( $csvItem.U_SapDfQCS -eq 'P')
        {
            $idt.U_SapDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Passed
        }
        if( $csvItem.U_SapDfQCS -eq 'T')
        {
            $idt.U_SapDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::QCTesting
        }

        if( $csvItem.U_PFDfBS -eq 'R')
        {
            $idt.U_PFDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Released;
        }
        if( $csvItem.U_PFDfBS -eq 'L')
        {
            $idt.U_PFDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Locked;
        }
        if( $csvItem.U_PFDfBS -eq 'A')
        {
            $idt.U_PFDfBS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::NotAccesible;
        }

         if( $csvItem.U_PFDfQCS -eq 'F')
        {
            $idt.U_PFDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Failed
        }
        if( $csvItem.U_PFDfQCS -eq 'H')
        {
            $idt.U_PFDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::OnHold
        }
        if( $csvItem.U_PFDfQCS -eq 'I')
        {
            $idt.U_PFDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Inspection
        }
        if( $csvItem.U_PFDfQCS -eq 'P')
        {
            $idt.U_PFDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Passed
        }
        if( $csvItem.U_PFDfQCS -eq 'T')
        {
            $idt.U_PFDfQCS = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::QCTesting
        }
    }
    else
    {
        $idt.U_InheritStatus = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
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