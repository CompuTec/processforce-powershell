#region #Script info
########################################################################
# CompuTec PowerShell Script - Import Bill of Materials Structures
########################################################################
# Version: 2.0
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Test Protocol. Script add new or will update existing data.   
#      You need to have all requred files for import. 
#      Sctipt check that Test Properies exists in the system during importing Test Protocol.
#      By default script is using his location/startup path as root path for csv files.
# Warning:
#   Make sure that item & item details was imported before use this script.
#   It's recommended run script when all users all disconnected.
#   Before running this script please do database backup.
# Troubleshooting:
#   https://connect.computec.pl/display/PF930EN/FAQ+PowerShell
#   https://connect.computec.pl/display/PF920EN/FAQ+PowerShell
# Script source:
#   https://connect.computec.pl/display/PF930EN/Quality+Control+scripts
########################################################################
#endregion

#region #PF API library usage
Clear-Host
# You need to check in what architecture PowerShell ISE is running (x64 or x86),
# you need run PowerShell ISE in the same architecture like PF API is installed (check in Windows -> Programs & Features)
# Examples:
#     SAP Client x64 + PF x64 installed on DB/Company => PF API x64 => Windows PowerShell ISE
#     SAP Client x86 + PF x86 installed on DB/Company => PF API x86 => Windows PowerShell ISE x86

[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#endregion

#region #Paths to csv files

$csvImportCatalog = $PSScriptRoot + "\";

#If you are using lower version of PowerShell than PowerShell 4.0 you can use static path
# $csvImportCatalog = "C:\PS\PF\TestProtocols\"; 

#endregion

#region #Datbase/Company connection settings

$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.LicenseServer = "10.0.0.240:40000"
$pfcCompany.SQLServer = "10.0.0.240:30015"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"
$pfcCompany.Databasename = "PFDEMOGB_MICHALB"
$pfcCompany.UserName = "michalb"
$pfcCompany.Password = "1234"
 
# where:
 
# LicenseServer = SAP LicenceServer name or IP Address with port number (see in SAP Client -> Administration -> Licence -> Licence Administration -> Licence Server)
# SQLServer     = Server name or IP Address with port number, should be the same like in System Landscape Dirctory (see https://<Server>:<Port>/ControlCenter) - sometimes best is use IP Address for resolve connection problems.
#
# DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2016"     # For MsSQL Server 2016
#                [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2014"     # For MsSQL Server 2014
#                [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2012"     # For MsSQL Server 2012
#                [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"        # For HANA
#
# Databasename = Database / schema name (check in SAP Company select form/window, or in MsSQL Management Studio or in HANA Studio)
# UserName     = SAP user name ex. manager
# Password     = SAP user password
 
#endregion

#region #Connect to company
 
write-host -backgroundcolor yellow -foregroundcolor black  "Trying connect..."
$version = [CompuTec.Core.CoreConfiguration+DatabaseSetup]::AddonVersion
write-host -backgroundcolor green -foregroundcolor black "PF API Library:" $version';' 'Host:'(Get-WmiObject Win32_OperatingSystem).CSName';' 'OSArchitecture:' (Get-WmiObject Win32_OperatingSystem).OSArchitecture
 
try {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'code')]
    $code = $pfcCompany.Connect()
 
    write-host -backgroundcolor green -foregroundcolor black "Connected to:" $pfcCompany.SapCompany.CompanyName "/ " $pfcCompany.SapCompany.CompanyDB"" "Sap Company version: " $pfcCompany.SapCompany.Version
}
catch {
    #Show error messages & stop the script
    write-host "Connection Failure: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
 
    write-host "LicenseServer:" $pfcCompany.LicenseServer
    write-host "SQLServer:" $pfcCompany.SQLServer
    write-host "DbServerType:" $pfcCompany.DbServerType
    write-host "Databasename" $pfcCompany.Databasename
    write-host "UserName:" $pfcCompany.UserName
}

#If company is not connected - stops the script
if (-not $pfcCompany.IsConnected) {
    write-host -backgroundcolor yellow -foregroundcolor black "Company is not connected"
    return 
}
 
#endregion
        

#Data loading from a csv file - Header information for Test Protocol
$csvTests = Import-Csv -Delimiter ';' -Path ($csvImportCatalog + "Quality_TestProtocols.csv")
 
#Checking that Test Protocol already exist 
 foreach($csvTest in $csvTests)
 {
 write-host $csvTest.TestProtocolCode " - Importing"  -backgroundcolor yellow -foregroundcolor black
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
	 $qtppt_path = $csvImportCatalog + "Quality_TestProtocolsPropertiesTest.csv"
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
                #Check that TestProperty exist in the system
                $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
                $query = [string]::Format("select ""U_TestPrpCode"" from ""@CT_PF_OTPR"" where ""U_TestPrpCode"" = '{0}';", $prop.PropertyCode);
                $rs.DoQuery($query)
     
                if ($rs.RecordCount -eq 0) {
                    write-host "   Test Protocol:" $test.U_TestPrclCode "-> Test Property code:" $prop.PropertyCode " can't be found (Pleas add it in Test Properties or check your import data for importing: Test Properties -> file: TestProperties.csv)" -backgroundcolor red -foregroundcolor white $_.Exception.Message
                    continue;
                }

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
	 $qtppi_path = $csvImportCatalog + "Quality_TestProtocolsPropertiesItem.csv"
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
               

				$test.Properties.U_PrpCode = $prop.PropertyCode;
				$test.Properties.U_Expression = $prop.Expression;
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
	 $qtpr_path = $csvImportCatalog + "Quality_TestProtocolsResources.csv"
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
	 $qtpi_path = $csvImportCatalog + "Quality_TestProtocolsItems.csv"
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
    try
    { 
    
        if($exists -eq 1)
        {
            [System.String]::Format("   Updating Test Protocol: {0}", $csvTest.TestProtocolCode)
            $message = $test.Update()
        }
        else
        {
            [System.String]::Format("   Adding Test Protocol: {0}", $csvTest.TestProtocolCode)
            $message= $test.Add()
        }
       
        if($message -lt 0)
        {    
            $err=$pfcCompany.GetLastErrorDescription()
            write-host -backgroundcolor red -foregroundcolor white "   " $err
        } 
        else
        {
            write-host "   Success"
        }   

    }catch
    {
        $err=$pfcCompany.GetLastErrorDescription()
        write-host -backgroundcolor red -foregroundcolor white "   " $err
    } 
    
  }

#region Close connection
 
if ($pfcCompany.IsConnected) {
    $pfcCompany.Disconnect()
     
    write-host " "
    write-host  –backgroundcolor green –foregroundcolor black "Disconnected from the company"
}
 
#endregion