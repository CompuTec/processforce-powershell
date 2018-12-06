#region #Script info
########################################################################
# CompuTec PowerShell Script - Import Quality Control Test Protocols
########################################################################
$SCRIPT_VERSION = "3.3"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 05 R1 HF1 (64-bit)
# Description:
#      Import Test Protocol. Script add new or will update existing data.
#      You need to have all requred files for import.
#      Sctipt check that Test Properies exists in the system during importing Test Protocol.
#      By default script is using his location/startup path as root path for csv files.
# Warning:
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
Write-Host -backgroundcolor Yellow -foregroundcolor DarkBlue ("Script Version:" + $SCRIPT_VERSION)
# You need to check in what architecture PowerShell ISE is running (x64 or x86),
# you need run PowerShell ISE in the same architecture like PF API is installed (check in Windows -> Programs & Features)
# Examples:
#     SAP Client x64 + PF x64 installed on DB/Company => PF API x64 => Windows PowerShell ISE
#     SAP Client x86 + PF x86 installed on DB/Company => PF API x86 => Windows PowerShell ISE x86

[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#endregion

#region Script parameters

$csvImportCatalog = $PSScriptRoot + "\";

#If you are using lower version of PowerShell than PowerShell 4.0 you can use static path
# $csvImportCatalog = "C:\PS\PF\TestProtocols\";

$csvTestsProtocolsPath = -join ($csvImportCatalog, "Quality_TestProtocols.csv");
$csvTestsProtocolsItemsPath = -join ($csvImportCatalog, "Quality_TestProtocolsItems.csv");
$csvTestsProtocolsPropertiesItemPath = -join ($csvImportCatalog, "Quality_TestProtocolsPropertiesItem.csv");
$csvTestsProtocolsPropertiesTestPath = -join ($csvImportCatalog, "Quality_TestProtocolsPropertiesTest.csv");
$csvTestsProtocolsResourcesPath = -join ($csvImportCatalog, "Quality_TestProtocolsResources.csv");
$csvTestsProtocolsAttachmentsPath = -join ($csvImportCatalog, "Quality_TestProtocolsAttachments.csv");

#endregion

#region #Datbase/Company connection settings
#configuration xml
$configurationXMLFilePath = -join ($csvImportCatalog, "configuration.xml");
if (!(Test-Path $configurationXMLFilePath -PathType Leaf)) {
    Write-Host -BackgroundColor Red ([string]::Format("File: {0} don't exists.", $configurationXMLFilePath));
    return;
}
[xml] $configurationXml = Get-Content -Encoding UTF8 $configurationXMLFilePath
$xmlConnection = $configurationXml.SelectSingleNode("/configuration/connection");

$connectionConfirmation = [string]::Format('You are connecting to Database: {0} on Server: {1} as User: {2}. Do you want to continue [y/n]?:', $xmlConnection.Database, $xmlConnection.SQLServer, $xmlConnection.UserName);
Write-Host $connectionConfirmation -backgroundcolor Yellow -foregroundcolor DarkBlue -NoNewline
$confirmation = Read-Host
if (($confirmation -ne 'y') -and ($confirmation -ne 'Y')) {
    return;
}

$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.LicenseServer = $xmlConnection.LicenseServer;
$pfcCompany.SQLServer = $xmlConnection.SQLServer;
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::[string]$xmlConnection.DbServerType;
$pfcCompany.Databasename = $xmlConnection.Database;
$pfcCompany.UserName = $xmlConnection.UserName;
$pfcCompany.Password = $xmlConnection.Password;
 
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
        
try {
    #Data loading from a csv file
    write-host ""

    [array] $csvTestsProtocols = Import-Csv -Delimiter ';' -Path $csvTestsProtocolsPath
    [array] $csvTestsProtocolsItems = $null
    if ((Test-Path -Path $csvTestsProtocolsItemsPath -PathType leaf) -eq $true) {
        [array] $csvTestsProtocolsItems = Import-Csv -Delimiter ';' -Path $csvTestsProtocolsItemsPath
    }
    else {
        write-host "Items - csv not available."
    }

    [array] $csvTestsProtocolsPropertiesItem = $null
    if ((Test-Path -Path $csvTestsProtocolsPropertiesItemPath -PathType leaf) -eq $true) {
        [array] $csvTestsProtocolsPropertiesItem = Import-Csv -Delimiter ';' -Path $csvTestsProtocolsPropertiesItemPath
    }
    else {
        write-host "Item Properties - csv not available."
    }

    [array] $csvTestsProtocolsPropertiesTest = $null
    if ((Test-Path -Path $csvTestsProtocolsPropertiesTestPath -PathType leaf) -eq $true) {
        [array] $csvTestsProtocolsPropertiesTest = Import-Csv -Delimiter ';' -Path $csvTestsProtocolsPropertiesTestPath
    }
    else {
        write-host "Properties - csv not available."
    }

    [array] $csvTestsProtocolsResources = $null
    if ((Test-Path -Path $csvTestsProtocolsResourcesPath -PathType leaf) -eq $true) {
        [array] $csvTestsProtocolsResources = Import-Csv -Delimiter ';' -Path $csvTestsProtocolsResourcesPath
    }
    else {
        write-host "Resources - csv not available."
    }

    [array] $csvTestsProtocolsAttachments = $null
    if ((Test-Path -Path $csvTestsProtocolsAttachmentsPath -PathType leaf) -eq $true) {
        [array] $csvTestsProtocolsAttachments = Import-Csv -Delimiter ';' -Path $csvTestsProtocolsAttachmentsPath
    }
    else {
        write-host "Attachment - csv not available."
    }


    write-Host 'Preparing data: '
    $totalRows = $csvTestsProtocols.Count + $csvTestsProtocolsItems.Count + $csvTestsProtocolsPropertiesItem.Count + $csvTestsProtocolsPropertiesTest.Count + $csvTestsProtocolsResources.Count + $csvTestsProtocolsAttachments.Count;

    $protocolsList = New-Object 'System.Collections.Generic.List[array]'
    $dictionaryItems = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryPropertiesItem = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryPropertiesTest = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryResources = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryAttachments = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;

    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $csvTestsProtocols) {
        $key = $row.TestProtocolCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

   
        $protocolsList.Add([array]$row);
    }

    foreach ($row in $csvTestsProtocolsItems) {
        $key = $row.TestProtocolCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryItems.ContainsKey($key) -eq $false) {
            $dictionaryItems[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $dictionaryItems[$key];
    
        $list.Add([array]$row);
    }
    foreach ($row in $csvTestsProtocolsPropertiesItem) {
        $key = $row.TestProtocolCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryPropertiesItem.ContainsKey($key) -eq $false) {
            $dictionaryPropertiesItem[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $dictionaryPropertiesItem[$key];
    
        $list.Add([array]$row);
    }
    foreach ($row in $csvTestsProtocolsPropertiesTest) {
        $key = $row.TestProtocolCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryPropertiesTest.ContainsKey($key) -eq $false) {
            $dictionaryPropertiesTest[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $dictionaryPropertiesTest[$key];
    
        $list.Add([array]$row);
    }
    foreach ($row in $csvTestsProtocolsResources) {
        $key = $row.TestProtocolCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryResources.ContainsKey($key) -eq $false) {
            $dictionaryResources[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $dictionaryResources[$key];
    
        $list.Add([array]$row);
    }
    foreach ($row in $csvTestsProtocolsAttachments) {
        $key = $row.TestProtocolCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryAttachments.ContainsKey($key) -eq $false) {
            $dictionaryAttachments[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $dictionaryAttachments[$key];
    
        $list.Add([array]$row);
    }


    Write-Host '';
    Write-Host 'Adding/updating data: ' -NoNewLine;

    #Data loading from a csv file - Header information for Test Protocol
    $csvTests = $protocolsList;
    if ($protocolsList.Count -gt 1) {
        $total = $protocolsList.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;

    $qmTP = New-Object 'CompuTec.Core.DI.Database.QueryManager';
    $qmTP.CommandText = "SELECT ""U_TestPrclCode"", ""Code"" FROM ""@CT_PF_OTCL"" WHERE ""U_TestPrclCode"" = @TestPrclCode";
    $qmTestPrp = New-Object 'CompuTec.Core.DI.Database.QueryManager'
    $qmTestPrp.CommandText = "SELECT ""U_TestPrpCode"" FROM ""@CT_PF_OTPR"" WHERE ""U_TestPrpCode"" = @TestPrpCode;";
    #Checking that Test Protocol already exist 
    foreach ($csvTest in $csvTests) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            
            $qmTP.ClearParameters();
            $qmTP.AddParameter("TestPrclCode",$csvTest.TestProtocolCode);
            $rs = $qmTP.Execute($pfcCompany.Token);   
            $exists = $false;
            if ($rs.RecordCount -gt 0) {
                $exists = $true;
                $dummy = $rs.MoveFirst();
            }
    
       
            #Creating TestProtocol
            $test = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"TestProtocol")
    
    
            if ($exists -eq 1) {
                $dummy = $test.getByKey($rs.Fields.Item('Code').Value);
            }
            else {
                $test.U_TestPrclCode = $csvTest.TestProtocolCode;
                $test.U_TestPrclName = $csvTest.TestProtocolName;
            }
		
            $test.U_ItemCode = $csvTest.ItemCode;
            $test.U_TemplateCode = $csvTest.TemplateCode;
		
            if ($csvTest.RevisionCode -ne "") {
                $test.U_RevCode = $csvTest.RevisionCode;
            }
            if ($csvTest.Warehouse -ne "") {
                $test.U_WhsCode = $csvTest.Warehouse;
            }
            if ($csvTest.Project -ne "") {
                $test.U_Project = $csvTest.Project;
            }
            if ($csvTest.ValidFrom -ne "") {
                $test.U_ValidFrom = $csvTest.ValidFrom;
            }
            else {
                $test.U_ValidFrom = [DateTime]::MinValue;
            }

            if ($csvTest.ValidTo -ne "") {
                $test.U_ValidTo = $csvTest.ValidTo;
            }
            else {
                $test.U_ValidTo = [DateTime]::MinValue 
            }
		
            #Frequency
            $test.U_FrqQuantity = $csvTest.FrqQuantity;
            $test.U_FrqUoM = $csvTest.FrqUoM;
            $test.U_FrqPercentage = $csvTest.FrqPercentage;
            $test.U_FrqTimeBtwnTests = $csvTest.FrqTimeBtwnTests;
            $test.U_FrqAfterNoBatch = $csvTest.FrqAfterNoBatch;
            $test.U_FrqRecInspDate = $csvTest.FrqRecInspDate;
            if ($csvTest.FrqSpecDate -ne "") {
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
            [array]$Properties = $dictionaryPropertiesTest[$csvTest.TestProtocolCode];
            if ($Properties.count -gt 0) {
                #Deleting all exisitng Properties
                $count = $test.Properties.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $test.Properties.DelRowAtPos(0);
                }
                $test.Properties.SetCurrentLine($test.Properties.Count - 1);
	         
                #Adding Properties
                foreach ($prop in $Properties) {
                    #Check that TestProperty exist in the system
                    $qmTestPrp.ClearParameters();
                    $qmTestPrp.AddParameter("TestPrpCode",$prop.PropertyCode);
                    $rsProp = $qmTestPrp.Execute($pfcCompany.Token);
     
                    if ($rsProp.RecordCount -eq 0) {
                        write-host "   Test Protocol:" $test.U_TestPrclCode "-> Test Property code:" $prop.PropertyCode " can't be found (Pleas add it in Test Properties or check your import data for importing: Test Properties -> file: TestProperties.csv)" -backgroundcolor red -foregroundcolor white $_.Exception.Message
                        continue;
                    }

                    $test.Properties.U_PrpCode = $prop.PropertyCode;
                    $test.Properties.U_Expression = $prop.Expression;
				
                    if ($prop.RangeFrom -ne "") {
                        $test.Properties.U_RangeValueFrom = $prop.RangeFrom;
                    }
                    else {
                        $test.Properties.U_RangeValueFrom = 0;
                    }
                    $test.Properties.U_RangeValueTo = $prop.RangeTo;
				
                    if ($prop.UoM -ne "") {
                        $test.Properties.U_UnitOfMeasure = $prop.UoM;
                    }
				
                    if ($prop.ReferenceCode -ne "") {
                        $test.Properties.U_RefCode = $prop.ReferenceCode;
                    }
				
                    if ($prop.ValidFrom -ne "") {
                        $test.Properties.U_ValidFromDate = $prop.ValidFrom;
                    }
                    else {
                        $test.Properties.U_ValidFromDate = [DateTime]::MinValue;
                    }
                    if ($prop.ValidTo -ne "") {
                        $test.Properties.U_ValidToDate = $prop.ValidTo
                    }
                    else {
                        $test.Properties.U_ValidToDate = [DateTime]::MinValue;
                    }
				
                    $test.Properties.U_Remarks = $prop.Remarks
				
                    $dummy = $test.Properties.Add()
                }
            }
  	
	
            #ItemProperties
            [array]$ItemProperties = $dictionaryPropertiesItem[$csvTest.TestProtocolCode];
            if ($ItemProperties.count -gt 0) {
                #Deleting all exisitng ItemProperties
                $count = $test.ItemProperties.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $test.ItemProperties.DelRowAtPos(0);
                }
                $test.ItemProperties.SetCurrentLine($test.ItemProperties.Count - 1);
	         
                #Adding Item Properies
                foreach ($itprop in $ItemProperties) {
                    $test.Properties.U_PrpCode = $prop.PropertyCode;
                    $test.Properties.U_Expression = $prop.Expression;
                    $test.ItemProperties.U_PrpCode = $itprop.PropertyCode;
                    $test.ItemProperties.U_Expression = $itprop.Expression;
                    if ($itprop.RangeFrom -ne "") {
                        $test.ItemProperties.U_RangeValueFrom = $itprop.RangeFrom;
                    }
                    else {
                        $test.ItemProperties.U_RangeValueFrom = 0;
                    }
                    $test.ItemProperties.U_RangeValueTo = $itprop.RangeTo;
                    if ($itprop.ReferenceCode -ne "") {
                        $test.ItemProperties.U_RefCode = $itprop.ReferenceCode;
                    }
				
                    if ($itprop.ValidFrom -ne "") {
                        $test.ItemProperties.U_ValidFromDate = $itprop.ValidFrom;
                    }
                    else {
                        $test.ItemProperties.U_ValidFromDate = [DateTime]::MinValue;
                    }
                    if ($itprop.ValidTo -ne "") {
                        $test.ItemProperties.U_ValidToDate = $itprop.ValidTo
                    }
                    else {
                        $test.ItemProperties.U_ValidToDate = [DateTime]::MinValue;
                    }
				
                    $test.ItemProperties.U_Remarks = $itprop.Remarks
				
                    $dummy = $test.ItemProperties.Add()
                }
            }
  	
	
            #Resources
            [array]$Resources = $dictionaryResources[$csvTest.TestProtocolCode];
            if ($Resources.count -gt 0) {
                #Deleting all exisitng Resources
                $count = $test.Resources.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $test.Resources.DelRowAtPos(0);
                }
                $test.Resources.SetCurrentLine($test.Resources.Count - 1);
	         
                #Adding Resources
                foreach ($resource in $Resources) {
                    $test.Resources.U_RscCode = $resource.ResourceCode;
                    $test.Resources.U_Quantity = $resource.Quantity;
                    $test.Resources.U_Remarks = $resource.Remarks;
				
                    if ($resource.ValidFrom -ne "") {
                        $test.Resources.U_ValidFrom = $resource.ValidFrom;
                    }
                    else {
                        $test.Resources.U_ValidFrom = [DateTime]::MinValue;
                    }
                    if ($resource.ValidTo -ne "") {
                        $test.Resources.U_ValidTo = $resource.ValidTo
                    }
                    else {
                        $test.Resources.U_ValidTo = [DateTime]::MinValue;
                    }
                    $dummy = $test.Resources.Add()
                }
            }
  	
	
            #Items
            [array]$Items = $dictionaryItems[$csvTest.TestProtocolCode];
            if ($Items.count -gt 0) {
                #Deleting all exisitng Items
                $count = $test.Items.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $test.Items.DelRowAtPos(0);
                }
                $test.Items.SetCurrentLine($test.Items.Count - 1);
	         
                #Adding Items
                foreach ($Item in $Items) {
                    $test.Items.U_ItemCode = $Item.ItemCode;
                    $test.Items.U_WhsCode = $Item.Warehouse;
                    $test.Items.U_Quantity = $Item.Quantity;
                    if ($Item.ValidFrom -ne "") {
                        $test.Items.U_ValidFrom = [DateTime] $Item.ValidFrom;
                    }
                    else {
                        $test.Items.U_ValidFrom = [DateTime]::MinValue;
                    }
                    if ($Item.ValidTo -ne "") {
                        $test.Items.U_ValidTo = [DateTime] $Item.ValidTo
                    }
                    else {
                        $test.Items.U_ValidTo = [DateTime]::MinValue;
                    }
				
                    $test.Items.U_Remarks = $Item.Remarks;
                    $dummy = $test.Items.Add()
                }
            }
      
            #Adding attachments to Resources
            [array]$resourcesAttachments = $dictionaryAttachments[$csvTest.TestProtocolCode];
            if ($resourcesAttachments.count -gt 0) {
                #Deleting all existing attachments for protocol
                $count = $test.Attachments.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $test.Attachments.DelRowAtPos(0);
                }
    
                #Adding the new data
                foreach ($att in $resourcesAttachments) {
                    $fileName = [System.IO.Path]::GetFileName($att.AttachmentPath)
                    $path = $att.AttachmentPath.Substring(0, ($att.AttachmentPath.Length - $fileName.Length) - 1);
                    $test.Attachments.U_FileName = $fileName;
                    $test.Attachments.U_AttDate = [System.DateTime]::Today;
                    $test.Attachments.U_Path = $path;
                    $dummy = $test.Attachments.Add();
                }
            }
	
            $message = 0

            #Adding or updating Test depends on exists in the database
    
            if ($exists -eq $true) {
                $message = $test.Update()
            }
            else {
                $message = $test.Add()
            }
       
            if ($message -lt 0) {    
                $err = $pfcCompany.GetLastErrorDescription()
                Throw [System.Exception] ($err)
            }
        }
        Catch {
            $err = $_.Exception.Message;
            if ($exists -eq $true) {
                $taskMsg = "adding";
            }
            else {
                $taskMsg = "updating"
            }
            $ms = [string]::Format("Error when {0} Test Protocol with Code {1} Details: {2}", $taskMsg, $csvTest.TestProtocolCode, $err);
            Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
            if ($pfcCompany.InTransaction) {
                $pfcCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
            } 
        }		 
      
    
    }
}
Catch {
    $err = $_.Exception.Message;
    $ms = [string]::Format("Exception occured: {0}", $err);
    Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
    if ($pfcCompany.InTransaction) {
        $pfcCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
    } 
}
Finally {
    #region Close connection
    if ($pfcCompany.IsConnected) {
        $pfcCompany.Disconnect()
        Write-Host '';
        write-host  –backgroundcolor green –foregroundcolor black "Disconnected from the company"
    }
    #endregion
}