clear
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
        
$code = $pfcCompany.Connect();
if($code -eq 1)
{
#Data loading from a csv file - Header information for Test Protocol
$csvTests = Import-Csv -Delimiter ';' -Path "C:\Quality_TestProtocols.csv"
 
#Checking that Test Protocol already exist 
 foreach($csvTest in $csvTests)
 {
  
        $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
     
        $rs.DoQuery([string]::Format( "SELECT ""U_TestPrclCode"", ""Code"" FROM ""@CT_PF_OTCL"" WHERE ""U_TestPrclCode"" = N'{0}'", $csvTest.TestProtocolCode))
        $exists = 0;
        if($rs.RecordCount -gt 0)
        {
            $exists = 1
			$rs.MoveFirst();
        }
  
    #Creating TestProtocol
    $test = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"TestProtocol")
    
    
    if($exists -eq 1)
    {
		
        $test.getByKey($rs.Fields.Item('Code').Value);
    }
    else
    {
        $test.U_TestPrclCode = $csvTest.TestProtocolCode;
		$test.U_TestPrclName = $csvTest.TestProtocolName;
	}
		
		$test.U_ItemCode = $csvTest.ItemCode;
		$test.U_TemplateCode = $csvTest.TemplateCode;
		
		if($csvTest.RevisionCode -ne "")
		{
			$test.U_RevCode = $csvTest.RevisionCode;
		}
		if($csvTest.Warehouse -ne "")
		{
			$test.U_WhsCode = $csvTest.Warehouse;
		}
		if($csvTest.Project -ne "")
		{
			$test.U_Project = $csvTest.Project;
		}
		if($csvTest.ValidFrom -ne "")
		{
			$test.U_ValidFrom = $csvTest.ValidFrom;
		}
        else
        {
            $test.U_ValidFrom = [DateTime]::MinValue;
        }

		if($csvTest.ValidTo -ne "")
		{
			$test.U_ValidTo = $csvTest.ValidTo;
		}
		else
		{
			$test.U_ValidTo = [DateTime]::MinValue 
		}
		
		#Frequency
		$test.U_FrqQuantity = $csvTest.FrqQuantity;
		$test.U_FrqUoM = $csvTest.FrqUoM;
		$test.U_FrqPercentage = $csvTest.FrqPercentage;
		$test.U_FrqTimeBtwnTests = $csvTest.FrqTimeBtwnTests;
		$test.U_FrqAfterNoBatch = $csvTest.FrqAfterNoBatch;
		$test.U_FrqRecInspDate = $csvTest.FrqRecInspDate;
		if($csvTest.FrqSpecDate -ne "")
		{
			$test.U_FrqSpecDate = $csvTest.FrqSpecDate;
		}
		$test.U_FrqRemarks = $csvTest.FrqRemarks;
		
		#Transactions
		$test.U_TrsPurGdsRcptPo = $csvTest.TrsPurGdsRcptPo;
		$test.U_TrsPurApInv = $csvTest.TrsPurApInv;
		$test.U_TrsPurGdsRcptPoBp = $csvTest.TrsPurGdsRcptPoBp;
		$test.U_TrsMnfPickRcpt = $csvTest.TrsMnfPickRcpt;
		$test.U_TrsMnfGdsRcpt = $csvTest.TrsMnfGdsRcpt;
		$test.U_TrsMnfPickRcptBp = $csvTest.TrsMnfPickRcptBp;
		$test.U_TrsMnfOrder = $csvTest.TrsMnfOrder;
		$test.U_TrsOprCode = $csvTest.TrsOprCode;
		$test.U_TrsInvBtchReTest = $csvTest.TrsInvBtchReTest;
		$test.U_TrsInvSnReTest = $csvTest.TrsInvSnReTest;
		$test.U_Instructions = $csvTest.Instructions;
	
	#Properties
     #Data loading from the csv file - Properties for test from Quality_TestProtocolsPropertiesTest.csv file
	 #Checks if the file exists
	 $qtppt_path = "C:\Quality_TestProtocolsPropertiesTest.csv"
	 if(Test-Path($qtppt_path))
	 {
	    [array]$Properties = Import-Csv -Delimiter ';' -Path $qtppt_path | Where-Object {$_.TestProtocolCode -eq $csvTest.TestProtocolCode}
	    if($Properties.count -gt 0)
	    {
	        #Deleting all exisitng Properties
	        $count = $test.Properties.Count
	        for($i=0; $i -lt $count; $i++)
	        {
	            $dummy = $test.Properties.DelRowAtPos(0);
	        }
	        $test.Properties.SetCurrentLine($test.Properties.Count - 1);
	         
	        #Adding Properties
	        foreach($prop in $Properties) 
	        {
				$test.Properties.U_PrpCode = $prop.PropertyCode;
				$test.Properties.U_Expression = $prop.Expression;
				
				if($prop.RangeFrom -ne "")
				{
					$test.Properties.U_RangeValueFrom = $prop.RangeFrom;
				}
				else
				{
					$test.Properties.U_RangeValueFrom = 0;
				}
				$test.Properties.U_RangeValueTo = $prop.RangeTo;
				
				if($prop.UoM -ne "")
				{
					$test.Properties.U_UnitOfMeasure = $prop.UoM;
				}
				
				if($prop.ReferenceCode -ne "")
				{
					$test.Properties.U_RefCode = $prop.ReferenceCode;
				}
				
				if($prop.ValidFrom -ne "")
		        {
		            $test.Properties.U_ValidFromDate = $prop.ValidFrom;
		        }
		        else
		        {
		            $test.Properties.U_ValidFromDate = [DateTime]::MinValue;
		        }
				if($prop.ValidTo -ne "")
		        {
		            $test.Properties.U_ValidToDate = $prop.ValidTo
		        }
		        else
		        {
		            $test.Properties.U_ValidToDate = [DateTime]::MinValue;
		        }
				
				$test.Properties.U_Remarks = $prop.Remarks
				
	            $dummy = $test.Properties.Add()
	        }
	    }
  	}
	
	#ItemProperties
	 #Data loading from the csv file - ItemProperties for Test from Quality_TestProtocolsPropertiesItem.csv file
	 #Checks if the file exists
	 $qtppi_path = "C:\Quality_TestProtocolsPropertiesItem.csv"
	 if(Test-Path($qtppi_path))
	 {
	    [array]$ItemProperties = Import-Csv -Delimiter ';' -Path $qtppi_path | Where-Object {$_.TestProtocolCode -eq $csvTest.TestProtocolCode}
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
	        foreach($itprop in $ItemProperties) 
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
				if($itprop.ReferenceCode -ne "")
				{
					$test.ItemProperties.U_RefCode = $itprop.ReferenceCode;
				}
				
				if($itprop.ValidFrom -ne "")
		        {
		            $test.ItemProperties.U_ValidFromDate = $itprop.ValidFrom;
		        }
		        else
		        {
		            $test.ItemProperties.U_ValidFromDate = [DateTime]::MinValue;
		        }
				if($itprop.ValidTo -ne "")
		        {
		            $test.ItemProperties.U_ValidToDate = $itprop.ValidTo
		        }
		        else
		        {
		            $test.ItemProperties.U_ValidToDate = [DateTime]::MinValue;
		        }
				
				$test.ItemProperties.U_Remarks = $itprop.Remarks
				
	            $dummy = $test.ItemProperties.Add()
	        }
	    }
  	}
	
	#Resources
	 #Data loading from the csv file - Resources for Test from Quality_TestProtocolsResources.csv file
	 #Checks if the file exists
	 $qtpr_path = "C:\Quality_TestProtocolsResources.csv"
	 if(Test-Path($qtpr_path))
	 {
	    [array]$Resources = Import-Csv -Delimiter ';' -Path $qtpr_path | Where-Object {$_.TestProtocolCode -eq $csvTest.TestProtocolCode}
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
				$test.Resources.U_Quantity = $resource.Quantity;
				$test.Resources.U_Remarks = $resource.Remarks;
				
				if($resource.ValidFrom -ne "")
		        {
		            $test.Resources.U_ValidFrom = $resource.ValidFrom;
		        }
		        else
		        {
		            $test.Resources.U_ValidFrom = [DateTime]::MinValue;
		        }
				if($resource.ValidTo -ne "")
		        {
		            $test.Resources.U_ValidTo = $resource.ValidTo
		        }
		        else
		        {
		            $test.Resources.U_ValidTo = [DateTime]::MinValue;
		        }
				
				
				
	            $dummy = $test.Resources.Add()
	        }
	    }
  	}
	
	#Items
	 #Data loading from the csv file - Items for Test from Quality_TestProtocolsItems.csv file
	 #Checks if the file exists
	 $qtpi_path = "C:\Quality_TestProtocolsItems.csv"
	 if(Test-Path($qtpi_path))
	 {
	    [array]$Items = Import-Csv -Delimiter ';' -Path $qtpi_path | Where-Object {$_.TestProtocolCode -eq $csvTest.TestProtocolCode}
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
				$test.Items.U_Quantity = $Item.Quantity;
				if($Item.ValidFrom -ne "")
		        {
		            $test.Items.U_ValidFrom = $Item.ValidFrom;
		        }
		        else
		        {
		            $test.Items.U_ValidFrom = [DateTime]::MinValue;
		        }
				if($Item.ValidTo -ne "")
		        {
		            $test.Items.U_ValidTo = $Item.ValidTo
		        }
		        else
		        {
		            $test.Items.U_ValidTo = [DateTime]::MinValue;
		        }
				
				$test.Items.U_Remarks = $Item.Remarks;
	            $dummy = $test.Items.Add()
	        }
	    }
  	}
	
    $message = 0
     
    #Adding or updating Test depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Test Protocol: {0}", $csvTest.TestProtocolCode)
        $message = $test.Update()
    }
    else
    {
        [System.String]::Format("Adding Test Protocol: {0}", $csvTest.TestProtocolCode)
          $message= $test.Add()
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