clear
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLPassword = "sa"
$pfcCompany.SQLServer = "PLSW006"
$pfcCompany.SQLUserName = "sa"
$pfcCompany.Databasename = "PFDemo03"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2012
#[SAPbobsCOM.BoDataServerTypes]"dst_MSSQL2012"
        
$code = $pfcCompany.Connect();
if($code -eq 1)
{
#Data loading from a csv file - Header information for Test Protocol
$csvTests = Import-Csv -Delimiter ';' -Path "c:\PSDI\Quality_ControlTests.csv"
 
#Checking that Control Test already exist 
 foreach($csvTest in $csvTests)
 {
  
    #Creating ControlTest
    $test = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"QualityControlTest")
    
	$test.U_TestProtocolNo = $csvTest.TestProtocolCode;
	$test.U_ItemCode = $csvTest.ItemCode;
	$test.U_RevCode = $csvTest.RevisionCode;
	$test.U_WhsCode = $csvTest.Warehouse;
	$test.U_ComplaintNo = $csvTest.ComplaintNo;
	$test.U_PrjCode = $csvTest.Project;
	$test.U_InsCode = $csvTest.InspectorCode;
	$test.U_ElectronicSign = $csvTest.ElectronicSign;
	$test.U_Status = $csvTest.Status;
	if($csvTest.CreatedDate -ne "")
	{
		$test.U_Created = $csvTest.CreatedDate;
	}
	if($csvTest.StartDate -ne "")
	{
		$test.U_Start = $csvTest.StartDate;
	}
	if( $csvTest.OnHoldDate -ne "")
	{
		$test.U_OnHold =  $csvTest.OnHoldDate;
	}
	if( $csvTest.WaitingNcmrDate -ne "")
	{
		$test.U_WaitingNcmr = $csvTest.WaitingNcmrDate;
	}
	if( $csvTest.ClosedDate -ne "")
	{
		$test.U_Closed = $csvTest.ClosedDate;
	}
	
	$test.U_TestStatus = $csvTest.TestStatus;
	if ($csvTest.Pass_FailDate -ne "")
	{
		$test.U_PassFailDate = $csvTest.Pass_FailDate;
	}
#Defects
	$test.U_SampleSize = $csvTest.DefSampleSize;
	$test.U_UoM = $csvTest.DefUoM;
	$test.U_PassedQty = $csvTest.DefPassedQty;
	$test.U_DefectQty = $csvTest.DefectQty;
	$test.U_InvMove = $csvTest.InventoryMovements;
	$test.U_Ncmr = $csvTest.NCMR;
	$test.U_NcmrInsCode = $csvTest.NcmrInspectorCode;
	$test.U_Remarks = $csvTest.DefRemarks;
#Transactions
	$test.U_TransType = $csvTest.TransactionType;
	$test.U_BpCode = $csvTest.BPCode;
	$test.U_MnfOprCode = $csvTest.OperationCode;
	
	#Properties
     #Data loading from the csv file - Properties for test from Quality_ControlTestProperties.csv file
	 #Checks if the file exists
	 $qctp_path = "c:\PSDI\Quality_ControlTestProperties.csv"
	 if(Test-Path($qctp_path))
	 {
	    [array]$Properties = Import-Csv -Delimiter ';' -Path $qctp_path | Where-Object {$_.Key -eq $csvTest.Key}
	    if($Properties.count -gt 0)
	    {
	        #Deleting all exisitng Properties
	        $count = $test.TestResults.Count
	        for($i=0; $i -lt $count; $i++)
	        {
	            $dummy = $test.TestResults.DelRowAtPos(0);
	        }
	        $test.TestResults.SetCurrentLine($test.Properties.Count - 1);
	         
	        #Adding Properties
	        foreach($prop in $Properties) 
	        {
				$test.TestResults.U_PrpCode = $prop.PropertyCode;
				$test.TestResults.U_Expression = $prop.Expresion;
				
				if($prop.RangeFrom -ne "")
				{
					$test.TestResults.U_RangeValueFrom = $prop.RangeFrom;
				}
				else
				{
					$test.TestResults.U_RangeValueFrom = 0;
				}
				$test.TestResults.U_RangeValueTo = $prop.RangeTo;
				
				if($prop.UoM -ne "")
				{
					$test.TestResults.U_UnitOfMeasure = $prop.UoM;
				}
				
				$test.TestResults.U_TestedValue = $prop.TestedValue;
				
				if($prop.ReferenceCode -ne "")
				{
					$test.TestResults.U_RefCode = $prop.ReferenceCode;
				}
				$test.TestResults.U_TestedRefCode = $prop.TestedRefCode;
				$test.TestResults.U_PassFail = $prop.Pass_Fail;
				$test.TestResults.U_RsnCode = $prop.ReasonCode;
				$test.TestResults.U_Remarks = $prop.Remarks;
				
	            $dummy = $test.TestResults.Add()
	        }
	    }
  	}
	
	#ItemProperties
	 #Data loading from the csv file - ItemProperties for Test from Quality_ControlTestPropertiesItem.csv file
	 #Checks if the file exists
	 $qctpi_path = "c:\PSDI\Quality_ControlTestPropertiesItem.csv.csv"
	 if(Test-Path($qctpi_path))
	 {
	    [array]$ItemProperties = Import-Csv -Delimiter ';' -Path $qctpi_path | Where-Object {$_.Key -eq $csvTest.Key}
	    if($ItemProperties.count -gt 0)
	    {
	        #Deleting all exisitng ItemProperties
	        $count = $test.ItemProperties.Count
	        for($i=0; $i -lt $count; $i++)
	        {
	            $dummy = $test.ItemProperties.DelRowAtPos(0);
	        }
	        $test.ItemProperties.SetCurrentLine($test.ItemProperties.Count - 1);
	         
		    #Adding Item Properies
	        foreach($itprop in $properties) 
	        {
				$test.ItemProperties.U_PrpCode = $itprop.PropertyCode;
				$test.ItemProperties.U_Expression = $itprop.Expression;
				if($itprop.RangeFrom -ne "")
				{
					$test.ItemProperties.U_RangeValueFrom = $itprop.RangeFrom;
				}
				else
				{
					$test.ItemProperties.U_RangeValueFrom = 0;
				}
				$test.ItemProperties.U_RangeValueTo = $itprop.RangeTo;
				
				$test.ItemProperties.U_TestedValue = $itprop.TestedValue;
				
				if($itprop.ReferenceCode -ne "")
				{
					$test.ItemProperties.U_RefCode = $itprop.ReferenceCode;
				}
				$test.ItemProperties.U_TestedRefCode = $itprop.TestedRefCode;
				$test.ItemProperties.U_PassFail = $itprop.Pass_Fail;
				$test.ItemProperties.U_RsnCode = $itprop.ReasonCode;
				$test.ItemProperties.U_Remarks = $itprop.Remarks;
				
	            $dummy = $test.ItemProperties.Add()
	        }
	    }
  	}
	
	#Resources
	 #Data loading from the csv file - Resources for Test from Quality_TestProtocolsResources.csv file
	 #Checks if the file exists
	 $qctr_path = "c:\PSDI\Quality_ControlTestResources.csv"
	 if(Test-Path($qctr_path))
	 {
	    [array]$Resources = Import-Csv -Delimiter ';' -Path $qctr_path | Where-Object {$_.Key -eq $csvTest.Key}
	    if($Resources.count -gt 0)
	    {
	        #Deleting all exisitng Resources
	        $count = $test.Resources.Count
	        for($i=0; $i -lt $count; $i++)
	        {
	            $dummy = $test.Resources.DelRowAtPos(0);
	        }
	        $test.Resources.SetCurrentLine($test.Resources.Count - 1);
	         
		    #Adding Resources
	        foreach($resource in $Resources) 
	        {
				$test.Resources.U_RscCode = $resource.ResourceCode;
				$test.Resources.U_WhsCode = $resource.Warehouse;
				$test.Resources.U_PlannedQty = $resource.PlanedQuantity;
				$test.Resources.U_ActualQty = $resource.ActualQuantity;
				$test.Resources.U_Remarks = $resource.Remarks;
	            $dummy = $test.Resources.Add()
	        }
	    }
  	}
	
	#Items
	 #Data loading from the csv file - Items for Test from Quality_TestProtocolsResources.csv file
	 #Checks if the file exists
	 $qcti_path = "c:\PSDI\Quality_ControlTestItems.csv"
	 if(Test-Path($qcti_path))
	 {
	    [array]$Items = Import-Csv -Delimiter ';' -Path $qcti_path | Where-Object {$_.Key -eq $csvTest.Key}
	    if($Items.count -gt 0)
	    {
	        #Deleting all exisitng Items
	        $count = $test.Items.Count
	        for($i=0; $i -lt $count; $i++)
	        {
	            $dummy = $test.Items.DelRowAtPos(0);
	        }
	        $test.Items.SetCurrentLine($test.Items.Count - 1);
	         
		    #Adding Items
	        foreach($Item in $Items)
	        {
				$test.Items.U_ItemCode = $Item.ItemCode;
				$test.Items.U_WhsCode =  $Item.Warehouse;
				$test.Items.U_PlannedQty = $Item.PlanedQuantity;
				$test.Items.U_ActualQty = $Item.ActualQuantity;
				$test.Items.U_Remarks = $Item.Remarks;
	            $dummy = $test.Items.Add()
	        }
	    }
  	}
	
	#Defects
	 #Data loading from the csv file - Defects for Test from Quality_ControlTestDefects.csv file
	 #Checks if the file exists
	 $qctd_path = "c:\PSDI\Quality_ControlTestDefects.csv"
	 if(Test-Path($qctd_path))
	 {
	    [array]$defects = Import-Csv -Delimiter ';' -Path $qctd_path | Where-Object {$_.Key -eq $csvTest.Key}
	    if($defects.count -gt 0)
	    {
	        #Deleting all exisitng Items
	        $count = $test.Defects.Count
	        for($i=0; $i -lt $count; $i++)
	        {
	            $dummy = $test.Defects.DelRowAtPos(0);
	        }
	        $test.Defects.SetCurrentLine($test.Defects.Count - 1);
	         
		    #Adding Defects
	        foreach($defect in $defects)
	        {
				$test.Defects.U_DefCode = $defect.DefectCode;
	            $dummy = $test.Defects.Add()
	        }
	    }
  	}
	
	#Batches
	 #Data loading from the csv file - Batches for Test from Quality_ControlTestTransBatches.csv file
	 #Checks if the file exists
	 $qcttb_path = "c:\PSDI\Quality_ControlTestTransBatches.csv"
	 if(Test-Path($qcttb_path))
	 {
	    [array]$batches = Import-Csv -Delimiter ';' -Path $qcttb_path | Where-Object {$_.Key -eq $csvTest.Key}
	    if($batches.count -gt 0)
	    {
	        #Deleting all exisitng Items
	        $count = $test.Batches.Count
	        for($i=0; $i -lt $count; $i++)
	        {
	            $dummy = $test.Batches.DelRowAtPos(0);
	        }
	        $test.Batches.SetCurrentLine($test.Batches.Count - 1);
	         
		    #Adding Defects
	        foreach($batch in $batches)
	        {
				$test.Batches.U_Batch = $batch.Batch;
	            $dummy = $test.Batches.Add()
	        }
	    }
  	}
	
	#SerialNumbers
	 #Data loading from the csv file - Serial Numbers for Test from Quality_ControlTestTransSerialNumbers.csv file
	 #Checks if the file exists
	 $qcttsn_path = "c:\PSDI\Quality_ControlTestTransSerialNumbers.csv"
	 if(Test-Path($qcttsn_path))
	 {
	    [array]$SerilaNumbers = Import-Csv -Delimiter ';' -Path $qcttsn_path | Where-Object {$_.Key -eq $csvTest.Key}
	    if($SerilaNumbers.count -gt 0)
	    {
	        #Deleting all exisitng Items
	        $count = $test.SerialNumbers.Count
	        for($i=0; $i -lt $count; $i++)
	        {
	            $dummy = $test.SerialNumbers.DelRowAtPos(0);
	        }
	        $test.SerialNumbers.SetCurrentLine($test.SerialNumbers.Count - 1);
	         
		    #Adding Defects
	        foreach($sn in $SerilaNumbers)
	        {
				$test.SerialNumbers.U_SerialNo = $sn.SerialNumber;
	            $dummy = $test.SerialNumbers.Add()
	        }
	    }
  	}

	
    $message = 0
     
    #Adding or updating Test depends on exists in the database
    
    write-host [System.String]::Format("Adding Item Details: {0}", $csvTest.Key);
    $message = $test.Add()
    
     
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