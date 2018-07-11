using module .\lib\CTLogger.psm1;
using module .\lib\CTProgress.psm1;

Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")
#add-type -Path "C:\PS_MP\MP_PS\SAP\DLL\Interop.SAPbobsCOM.dll"
#add-type -Path "C:\PS_MP\MP_PS\SAP\DLL\Interop.SAPbouiCOM.dll"

$ItemsDictionary = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[psobject]]';
$ResourcesList = New-Object 'System.Collections.Generic.List[string]';
$OperationsList = New-Object 'System.Collections.Generic.List[string]';
$RoutingsList = New-Object 'System.Collections.Generic.List[string]';
$ItemsDictionary.Add('MAKE', (New-Object 'System.Collections.Generic.List[psobject]'));
$ItemsDictionary.Add('BUY', (New-Object 'System.Collections.Generic.List[psobject]'));
[xml] $connectionConfigXml = Get-Content -Encoding UTF8 .\conf\Connection.xml
$xmlConnection = $connectionConfigXml.SelectSingleNode("/CT_CONFIG/Connection");
[xml] $TestConfigXml = Get-Content -Encoding UTF8 .\conf\TestConfig.xml
$MDConfigXml = $TestConfigXml.SelectSingleNode("/CT_CONFIG/MasterData");
$UIConfigXML = $TestConfigXml.SelectSingleNode("/CT_CONFIG/UI");
$RESULT_FILE = $PSScriptRoot + "\Results.csv";
$RESULT_FILE_CONF = $PSScriptRoot + "\Results_conf.csv";
function Imports() {

    
    
    [CTLogger] $logJobs = New-Object CTLogger ('DI', 'Import', $RESULT_FILE)

    #region connection
    
    $logJobs.startSubtask('Import');
    $logJobs.startSubtask('Connection');
    $pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany();
    $pfcCompany.LicenseServer = $xmlConnection.LicenseServer;
    $pfcCompany.SQLServer = $xmlConnection.SQLServer;
    $pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::[string]$xmlConnection.DbServerType;
    $pfcCompany.Databasename = $xmlConnection.Database;
    $pfcCompany.UserName = $xmlConnection.UserName;
    $pfcCompany.Password = $xmlConnection.Password;
    

    write-host -backgroundcolor yellow -foregroundcolor black  "Trying connect..."
    $version = [CompuTec.Core.CoreConfiguration+DatabaseSetup]::AddonVersion
    write-host -backgroundcolor green -foregroundcolor black "PF API Library:" $version';' 'Host:'(Get-WmiObject Win32_OperatingSystem).CSName';' 'OSArchitecture:' (Get-WmiObject Win32_OperatingSystem).OSArchitecture
    
    try {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
        $dummy = $pfcCompany.Connect()
    
        write-host -backgroundcolor green -foregroundcolor black "Connected to:" $pfcCompany.SapCompany.CompanyName "/ " $pfcCompany.SapCompany.CompanyDB"" "Sap Company version: " $pfcCompany.SapCompany.Version
    }
    catch {
        #Show error messages & stop the script
        write-host "Connection Failure: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
        #$logJobs.endSubtask('Connection', 'F', $_.Exception.Message);
        write-host "LicenseServer:" $pfcCompany.LicenseServer
        write-host "SQLServer:" $pfcCompany.SQLServer
        write-host "DbServerType:" $pfcCompany.DbServerType
        write-host "Databasename" $pfcCompany.Databasename
        write-host "UserName:" $pfcCompany.UserName
    }

    #If company is not connected - stops the script
    if (-not $pfcCompany.IsConnected) {
        write-host -backgroundcolor yellow -foregroundcolor black "Company is not connected"
        $logJobs.endSubtask('Connection', 'F', 'Company is not connected');
        return 
    }
    $logJobs.endSubtask('Connection', 'S', '');
    $sapCompany = $pfcCompany.SapCompany;
    #endregion


    function importIMD($sapCompany) {
        [CTLogger] $logIMD = New-Object CTLogger ('DI', 'Import Item Master Data', $RESULT_FILE)
        #region import of Item Master Data
        write-host ''
        write-host 'Import of Item Master Data: '
        $xmlItems = $MDConfigXml.SelectSingleNode([string]::Format("ItemMasterData"));

        $numberOfItems = [int] $xmlItems.NumberOfItems;
        $numberOfMakeItems = [int] $xmlItems.NumberOfMakeItems;
        $itemCodeLength = ([string]$numberOfItems).Length;
        $itemPrefix = [string] $xmlItems.Prefix
        $warehouseCode = [string] $xmlItems.WarehouseCode
        [CTProgress] $progress = New-Object CTProgress ($numberOfItems);
        for ($i = 0; $i -lt $numberOfItems; $i++) {
            try {
                $progress.next();
                $logIMD.startSubtask('Get Item Master Data');
                $sapIMD = $sapCompany.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oItems);
        
                $ItemCode = $itemPrefix + ([string]$i).PadLeft($itemCodeLength, '0');

                if ($i -lt $numberOfMakeItems) {
                    $ItemsDictionary['MAKE'].Add([psobject]@{
                            ItemCode  = $ItemCode
                            Revisions = New-Object 'System.Collections.Generic.List[string]';
                        });
                }
                else {
                    $ItemsDictionary['BUY'].Add([psobject]@{
                            ItemCode  = $ItemCode
                            Revisions = New-Object 'System.Collections.Generic.List[string]';
                        });
                }

                $retValue = $sapIMD.GetByKey($ItemCode)

                if ($retValue -eq $true) {
                    $logIMD.endSubtask('Get Item Master Data', 'S', 'Item Already Exists');
                    continue;
                }
                $logIMD.endSubtask('Get Item Master Data', 'S', '');
                $logIMD.startSubtask('Add Item Master Data');

                $sapIMD.ItemCode = $ItemCode;
                $sapIMD.ItemName = $ItemCode;

                $sapIMD.WhsInfo.WarehouseCode = $warehouseCode;
                $sapIMD.DefaultWarehouse = $warehouseCode;
            
                $message = $sapIMD.Add();
            
                if ($message -lt 0) {
                    $err = $sapCompany.GetLastErrorDescription();
                    Throw [System.Exception] ($err);
                }
                $logIMD.endSubtask('Add Item Master Data', 'S', '');
            }
            Catch {
                $err = $_.Exception.Message;
                $logIMD.endSubtask('Add Item Master Data', 'F', $err);
                continue;
            }
        }
        #endregion
    }

    function importItemDetails($pfcCompany) {
        [CTLogger] $logPFIMD = New-Object CTLogger ('DI', 'Import Item Details', $RESULT_FILE)
        write-host ''
        write-host 'Import of Item Details: '
        $xmlItemDetails = $MDConfigXml.SelectSingleNode([string]::Format("ItemDetails"));
        $numberOfRevisions = [int] $xmlItemDetails.NumberOfRevisions;
        $revisionCodeLength = ([string]$numberOfRevisions).Length;
    

        [CTProgress] $progress = New-Object CTProgress (($ItemsDictionary['MAKE'].Count + $ItemsDictionary['BUY'].Count));
        foreach ($itemType in $ItemsDictionary.Keys) {
            foreach ($item in $ItemsDictionary[$itemType]) {
                $progress.next()
                $itemCode = $item.ItemCode;
                try {
                    $logPFIMD.startSubtask('Get Item Details');
                    $itemDetails = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ItemDetails);
                    $itemDetailsExists = $itemDetails.GetByItemCode($itemCode);
                    $logPFIMD.endSubtask('Get Item Details', 'S', '');
                }
                catch {
                    $err = $_.Exception.Message;
                    $logPFIMD.endSubtask('Get Item Details', 'F', $err);
                    continue;
                }
                try {
                    if ($itemDetailsExists) {
                        $logPFIMD.startSubtask('Update Item Details');
                        # $count = $itemDetails.Revisions.Count
                        # for ($i = 0; $i -lt $count; $i++) {
                        #     $dummy = $itemDetails.Revisions.DelRowAtPos(0);
                        # }
                    }
                    else {
                        $logPFIMD.startSubtask('Add Item Details');
                        $itemDetails.U_ItemCode = $itemCode;
                    }
                    $itemDetails.Revisions.SetCurrentLine($itemDetails.Revisions.Count - 1);

                    for ($i = 0; $i -lt $numberOfRevisions; $i++) {
                        $revisionCode = 'code' + ([string]$i).PadLeft($revisionCodeLength, '0');
                        $item.Revisions.Add($revisionCode);
                        if ($itemDetailsExists -eq $false) {
                            $itemDetails.Revisions.U_Code = $revisionCode;
                            $itemDetails.Revisions.U_Description = $revisionCode;
                            $itemDetails.Revisions.U_Status = 1;
                
                            if ($i -eq 0) {
                                $itemDetails.Revisions.U_Default = 1;
                                $itemDetails.Revisions.U_IsMRPDefault = 1;
                                $itemDetails.Revisions.U_IsCostingDefault = 1;
                            }
                            else {
                                $itemDetails.Revisions.U_Default = 2;
                                $itemDetails.Revisions.U_IsMRPDefault = 2;
                                $itemDetails.Revisions.U_IsCostingDefault = 2;
                            }
                        
                            $dummy = $itemDetails.Revisions.Add()
                        }
                    }

                    $message = 0
      
                    if ($itemDetailsExists ) {
                        $message = $itemDetails.Update()
                    }
                    else {
                        $message = $itemDetails.Add()
                    }
      
                    if ($message -lt 0) {  
                        $err = $pfcCompany.GetLastErrorDescription()
                        Throw [System.Exception] ($err);
                    }
                    if ($itemDetailsExists) {
                        $logPFIMD.endSubtask('Update Item Details', 'S', '');
                    }
                    else {
                        $logPFIMD.endSubtask('Add Item Details', 'S', '');
                    }
                
                }
                catch {
                    $err = $_.Exception.Message;
                    if ($itemDetailsExists) {
                        $logPFIMD.endSubtask('Update Item Details', 'F', $err);
                    }
                    else {
                        $logPFIMD.endSubtask('Add Item Details', 'F', $err);
                    }
                    continue;
                }

            }
        }
    }

    function ImportBOMStructure($pfcCompany) {
        [CTLogger] $log = New-Object CTLogger ('DI', 'Import BOM Structure', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Import of BOM:'
        $xmlBOM = $MDConfigXml.SelectSingleNode([string]::Format("BOM"));
        $numberOfItems = [int] $xmlBOM.NumberOfItems;
        $numberOfBoms = [int] $xmlBOM.NumberOfBoms;
        $warehouseCode = [string] $xmlBOM.WarehouseCode;
        $itemsWarehouseCode = [string] $xmlBOM.ItemsWarehouseCode;
        [CTProgress] $progress = New-Object CTProgress ($numberOfBoms);
        for ($iBOM = 0; $iBOM -lt $numberOfBoms; $iBOM++) {
            try {
                $progress.next();
                $bomItemCode = $ItemsDictionary['MAKE'][$iBOM].ItemCode;
                $bomRevisionCode = $ItemsDictionary['MAKE'][$iBOM].Revisions[0];
                $log.startSubtask('Get BOM');
                $bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::BillOfMaterial);
                $exists = $bom.GetByItemCodeAndRevision($bomItemCode, $bomRevisionCode);
                if ($exists -eq -1) {
                    $bomExists = $false;
                }
                else {
                    $bomExists = $true;
                }
                $log.endSubtask('Get BOM', 'S', '');
            }
            catch {
                $err = $_.Exception.Message;
                $log.endSubtask('Get BOM', 'F', $err);
                continue;
            }
            try {
                if ($bomExists) {
                    $log.startSubtask('Update BOM');
                    $count = $bom.Items.Count
                    for ($i = 0; $i -lt $count; $i++) {
                        $dummy = $bom.Items.DelRowAtPos(0);
                    }
                }
                else {
                    $log.startSubtask('Add BOM');
                    $bom.U_ItemCode = $bomItemCode;
                    $bom.U_Revision = $bomRevisionCode;
                    $bom.U_WhsCode = $warehouseCode;
                }

                for ($iItems = 0; $iItems -lt $numberOfItems; $iItems++) {
                    $itemCode = $ItemsDictionary['BUY'][$iItems].ItemCode;
                    $revisionCode = $ItemsDictionary['BUY'][$iItems].Revisions[0];
                    #$bom.Items.U_Sequence = ($iItems * 10);
                    $bom.Items.U_ItemCode = $itemCode;
                    $bom.Items.U_Revision = $revisionCode;
                    $bom.Items.U_WhsCode = $itemsWarehouseCode;
                    $bom.Items.U_Factor = 1
                    $bom.Items.U_Quantity = 1
                    $bom.Items.U_ScrapPercentage = 0
                    $bom.Items.U_IssueType = 'M'
                    $bom.Items.U_SubRecepitItem = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
                    $dummy = $bom.Items.Add()
                }

                $message = 0
      
                if ($bomExists ) {
                    $message = $bom.Update()
                }
                else {
                    $message = $bom.Add()
                }
      
                if ($message -lt 0) {  
                    $err = $pfcCompany.GetLastErrorDescription()
                    Throw [System.Exception] ($err);
                }
                if ($bomExists) {
                    $log.endSubtask('Update BOM', 'S', '');
                }
                else {
                    $log.endSubtask('Add BOM', 'S', '');
                }
                
            }
            catch {
                $err = $_.Exception.Message;
                if ($bomExists) {
                    $log.endSubtask('Update BOM', 'F', $err);
                }
                else {
                    $log.endSubtask('Add BOM', 'F', $err);
                }
                continue;
            }

        }
    }
    function ImportResources($pfcCompany) {
        [CTLogger] $log = New-Object CTLogger ('DI', 'Import Resources', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Import of Resources:' -NoNewline
        $xmlResource = $MDConfigXml.SelectSingleNode([string]::Format("Resource"));
        $numberOfResources = [int] $xmlResource.NumberOfResources;
        $resourceCodeLength = ([string]$numberOfResources).Length;
        $resourcePrefix = [string] $xmlResource.Prefix
        [CTProgress] $progress = New-Object CTProgress ($numberOfResources);
        for ($i = 0; $i -lt $numberOfResources; $i++) {
            try {
                $progress.next();
                $ResourceCode = $resourcePrefix + ([string]$i).PadLeft($resourceCodeLength, '0');
                $ResourcesList.Add($ResourceCode);

                $log.startSubtask('Get Resource');
                $resource = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Resource);
                $exists = $resource.GetByRscCode($ResourceCode);
                if ($exists -eq -1) {
                    $resourceExists = $false;
                }
                else {
                    $resourceExists = $true;
                }
                $log.endSubtask('Get Resource', 'S', '');
            }
            catch {
                $err = $_.Exception.Message;
                $log.endSubtask('Get Resource', 'F', $err);
                continue;
            }
            try {
                if ($resourceExists) {
                    $log.startSubtask('Update Resource');
                }
                else {
                    $log.startSubtask('Add Resource');
                    $resource.U_RscType = 1;
                    $resource.U_RscCode = $ResourceCode
                    $resource.U_RscName = $ResourceCode
                }

                $message = 0
      
                if ($resourceExists ) {
                    $message = $resource.Update()
                }
                else {
                    $message = $resource.Add()
                }
      
                if ($message -lt 0) {  
                    $err = $pfcCompany.GetLastErrorDescription()
                    Throw [System.Exception] ($err);
                }
                if ($resourceExists) {
                    $log.endSubtask('Update Resource', 'S', '');
                }
                else {
                    $log.endSubtask('Add Resource', 'S', '');
                }
                
            }
            catch {
                $err = $_.Exception.Message;
                if ($resourceExists) {
                    $log.endSubtask('Update Resource', 'F', $err);
                }
                else {
                    $log.endSubtask('Add Resource', 'F', $err);
                }
                continue;
            }

        }
    }

    function ImportOperations($pfcCompany) {
        [CTLogger] $log = New-Object CTLogger ('DI', 'Import Operations', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Import of Operations:' -NoNewline
        $xmlOperation = $MDConfigXml.SelectSingleNode([string]::Format("Operation"));
        $numberOfOperations = [int] $xmlOperation.NumberOfOperations;
        $numberOfResources = [int] $xmlOperation.NumberOfResources;
        $operationCodeLength = ([string]$numberOfOperations).Length;
        $operationPrefix = [string] $xmlOperation.Prefix
        [CTProgress] $progress = New-Object CTProgress ($numberOfOperations);
        for ($iOperation = 0; $iOperation -lt $numberOfOperations; $iOperation++) {
            try {
                $progress.next();
                $operationCode = $operationPrefix + ([string]$iOperation).PadLeft($operationCodeLength, '0');
                $OperationsList.Add($operationCode);

                $log.startSubtask('Get Operation');
                $operation = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Operation);
                $exists = $operation.GetByOprCode($operationCode);
                if ($exists -eq -1) {
                    $operationExists = $false;
                }
                else {
                    $operationExists = $true;
                }
                $log.endSubtask('Get Operation', 'S', '');
            }
            catch {
                $err = $_.Exception.Message;
                $log.endSubtask('Get Operation', 'F', $err);
                continue;
            }
            try {
                if ($operationExists) {
                    $log.startSubtask('Update Operation');
                    $count = $operation.OperationResources.Count - 1
                    for ($i = $count - 1; $i -ge 0; $i--) {
                        $dummy = $operation.OperationResources.DelRowAtPos(0);
                    }
                }
                else {
                    $log.startSubtask('Add Operation');
                    $operation.U_OprCode = $operationCode;
                    $operation.U_OprName = $operationCode;
                }


                for ($iResources = 0; $iResources -lt $numberOfResources; $iResources++) {
                    $ResourceCode = $ResourcesList[$iResources];

                    $operation.OperationResources.U_RscCode = $ResourceCode
                    if ($iResources -eq 0) {
                        $operation.OperationResources.U_IsDefault = 'Y';
                    }
                    else {
                        $operation.OperationResources.U_IsDefault = 'N';
                    }
                
                    $operation.OperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No

                    $dummy = $operation.OperationResources.Add()
                }


                $message = 0
      
                if ($operationExists ) {
                    $message = $operation.Update()
                }
                else {
                    $message = $operation.Add()
                }
      
                if ($message -lt 0) {  
                    $err = $pfcCompany.GetLastErrorDescription()
                    Throw [System.Exception] ($err);
                }
                if ($operationExists) {
                    $log.endSubtask('Update Operation', 'S', '');
                }
                else {
                    $log.endSubtask('Add Operation', 'S', '');
                }
                
            }
            catch {
                $err = $_.Exception.Message;
                if ($operationExists) {
                    $log.endSubtask('Update Operation', 'F', $err);
                }
                else {
                    $log.endSubtask('Add Operation', 'F', $err);
                }
                continue;
            }

        }
    }
    function ImportRoutings($pfcCompany) {
        [CTLogger] $log = New-Object CTLogger ('DI', 'Import Routings', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Import of Routings:' -NoNewline
        $xmlRouting = $MDConfigXml.SelectSingleNode([string]::Format("Routing"));
        $numberOfOperations = [int] $xmlRouting.NumberOfOperations;
        $numberOfRoutings = [int] $xmlRouting.NumberOfRoutings;
        $routingCodeLength = ([string]$numberOfRoutings).Length;
        $routingPrefix = [string] $xmlRouting.Prefix
        [CTProgress] $progress = New-Object CTProgress ($numberOfRoutings);
        for ($i = 0; $i -lt $numberOfRoutings; $i++) {
            try {
                $progress.next();
                $routingCode = $routingPrefix + ([string]$i).PadLeft($routingCodeLength, '0');
                $RoutingsList.Add($routingCode);

                $log.startSubtask('Get Routing');
                $routing = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Routing);
                $exists = $routing.GetByRtgCode($routingCode);
                if ($exists -eq -1) {
                    $routingExists = $false;
                }
                else {
                    $routingExists = $true;
                }
                $log.endSubtask('Get Routing', 'S', '');
            }
            catch {
                $err = $_.Exception.Message;
                $log.endSubtask('Get Routing', 'F', $err);
                continue;
            }
            try {
                if ($routingExists) {
                    $log.startSubtask('Update Routing');
                    $count = $routing.Operations.Count
                    for ($j = 0; $j -lt $count; $j++) {
                        $dummy = $routing.Operations.DelRowAtPos(0);
                    }
                    $count = $routing.OperationResources.Count
                    for ($k = 0; $k -lt $count; $k++) {
                        $dummy = $routing.OperationResources.DelRowAtPos(0);
                    }   
                }
                else {
                    $log.startSubtask('Add Routing');
                    $routing.U_RtgCode = $routingCode
                    $routing.U_RtgName = $routingCode
                    $routing.U_Active = 1
                }


                for ($iOperations = 0; $iOperations -lt $numberOfOperations; $iOperations++) {
                    $OperationCode = $OperationsList[$iOperations];

                    $routing.Operations.U_OprCode = $OperationCode
                    $dummy = $routing.Operations.Add()
                }

                $message = 0
      
                if ($routingExists ) {
                    $message = $routing.Update()
                }
                else {
                    $message = $routing.Add()
                }
      
                if ($message -lt 0) {  
                    $err = $pfcCompany.GetLastErrorDescription()
                    Throw [System.Exception] ($err);
                }
                if ($routingExists) {
                    $log.endSubtask('Update Routing', 'S', '');
                }
                else {
                    $log.endSubtask('Add Routing', 'S', '');
                }
                
            }
            catch {
                $err = $_.Exception.Message;
                if ($routingExists) {
                    $log.endSubtask('Update Routing', 'F', $err);
                }
                else {
                    $log.endSubtask('Add Routing', 'F', $err);
                }
                continue;
            }

        }
    }
    function ImportProductionProcesses($pfcCompany) {
        [CTLogger] $log = New-Object CTLogger ('DI', 'Import Production Processes', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Import of Production Processes:' -NoNewline;
        $xmlProductionProcess = $MDConfigXml.SelectSingleNode([string]::Format("ProductionProcess"));
        $numberOfBoms = [int] $xmlProductionProcess.NumberOfBoms;
        $numberOfRoutings = [int] $xmlProductionProcess.NumberOfRoutings;
        [CTProgress] $progress = New-Object CTProgress ($numberOfBoms);
        for ($iBOM = 0; $iBOM -lt $numberOfBoms; $iBOM++) {
            try {
                $progress.next();
                $bomItemCode = $ItemsDictionary['MAKE'][$iBOM].ItemCode;
                $bomRevisionCode = $ItemsDictionary['MAKE'][$iBOM].Revisions[0];
                $log.startSubtask('Get Production Process');
                $bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::BillOfMaterial);
                $exists = $bom.GetByItemCodeAndRevision($bomItemCode, $bomRevisionCode);
                if ($exists -eq -1) {
                    $bomExists = $false;
                }
                else {
                    $bomExists = $true;
                }
                $log.endSubtask('Get Production Process', 'S', '');
            }
            catch {
                $err = $_.Exception.Message;
                $log.endSubtask('Get Production Process', 'F', $err);
                continue;
            }
            try {
                if ($bomExists) {
                    $log.startSubtask('Update Production Process');
                    $count = $bom.Routings.Count
                    for ($i = 0; $i -lt $count; $i++) {
                        $dummy = $bom.Routings.DelRowAtPos(0);
                    }
            
                    $count = $bom.RoutingOperations.Count
                    for ($i = 0; $i -lt $count; $i++) {
                        $dummy = $bom.RoutingOperations.DelRowAtPos(0);
                    }
                    
                    $count = $bom.RoutingOperationResources.Count
                    for ($i = 0; $i -lt $count; $i++) {
                        $dummy = $bom.RoutingOperationResources.DelRowAtPos(0);
                    }    
                }
                else {
                    $log.startSubtask('Add BOM');
                    $bom.U_ItemCode = $bomItemCode;
                    $bom.U_Revision = $bomRevisionCode;
                }

                for ($iRoutings = 0; $iRoutings -lt $numberOfRoutings; $iRoutings++) {
                    $routingCode = $RoutingsList[$iRoutings];
                    $bom.Routings.U_RtgCode = $routingCode;
                    if ($iRoutings -eq 0) {
                        $bom.Routings.U_IsDefault = 'Y';
                        $bom.Routings.U_IsRollUpDefault = 'Y'
                    }
                    else {
                        $bom.Routings.U_IsDefault = 'N'
                        $bom.Routings.U_IsRollUpDefault = 'N'
                    }
                    #  $bom.RoutingOperationResources
                    $dummy = $bom.Routings.Add()
                }

                $message = 0
      
                if ($bomExists ) {
                    $message = $bom.Update()
                }
                else {
                    $message = $bom.Add()
                }
      
                if ($message -lt 0) {  
                    $err = $pfcCompany.GetLastErrorDescription()
                    Throw [System.Exception] ($err);
                }
                if ($bomExists) {
                    $log.endSubtask('Update Production Process', 'S', '');
                }
                else {
                    $log.endSubtask('Add Production Process', 'S', '');
                }
                
            }
            catch {
                $err = $_.Exception.Message;
                if ($bomExists) {
                    $log.endSubtask('Update Production Process', 'F', $err);
                }
                else {
                    $log.endSubtask('Add Production Process', 'F', $err);
                }
                continue;
            }

        }
    }


    $logJobs.startSubtask('Import IMD');
    importIMD $sapCompany
    $logJobs.endSubtask('Import IMD', 'S', '');

    $logJobs.startSubtask('Import Item Details');
    importItemDetails $pfcCompany
    $logJobs.endSubtask('Import Item Details', 'S', '');

    #restore Item Costing

    $logJobs.startSubtask('Import BOM Structure');
    ImportBOMStructure $pfcCompany
    $logJobs.endSubtask('Import BOM Structure', 'S', '');

    $logJobs.startSubtask('Import Resources');
    ImportResources $pfcCompany
    $logJobs.endSubtask('Import Resources', 'S', '');

    $logJobs.startSubtask('Import Operations');
    ImportOperations $pfcCompany
    $logJobs.endSubtask('Import Operations', 'S', '');

    $logJobs.startSubtask('Import Routings');
    ImportRoutings $pfcCompany
    $logJobs.endSubtask('Import Routings', 'S', '');

    $logJobs.startSubtask('Import Production Processes');
    ImportProductionProcesses $pfcCompany
    $logJobs.endSubtask('Import Production Processes', 'S', '');

    $logJobs.endSubtask('Import', 'S', '');

    $pfcCompany.Disconnect();
}

function UITests() {
    [CTLogger] $logJobs = New-Object CTLogger ('UI', 'GET', $RESULT_FILE)

    #region connection
    $logJobs.startSubtask('Get');
    $logJobs.startSubtask('Connection');
    Write-Host -BackgroundColor Blue 'Connecting...'

    $app = $null;
    

    write-host -backgroundcolor yellow -foregroundcolor black  "Trying connect..."
    $version = [CompuTec.Core.CoreConfiguration+DatabaseSetup]::AddonVersion
    write-host -backgroundcolor green -foregroundcolor black "PF API Library:" $version';' 'Host:'(Get-WmiObject Win32_OperatingSystem).CSName';' 'OSArchitecture:' (Get-WmiObject Win32_OperatingSystem).OSArchitecture
    
    try {
        $pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::ConnectUI([ref] $app,$true)
        write-host -backgroundcolor green -foregroundcolor black "Connected to:" $pfcCompany.SapCompany.CompanyName "/ " $pfcCompany.SapCompany.CompanyDB"" "Sap Company version: " $pfcCompany.SapCompany.Version
    }
    catch {
        #Show error messages & stop the script
        write-host "Connection Failure: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
        #$logJobs.endSubtask('Connection', 'F', $_.Exception.Message);
        write-host "LicenseServer:" $pfcCompany.LicenseServer
        write-host "SQLServer:" $pfcCompany.SQLServer
        write-host "DbServerType:" $pfcCompany.DbServerType
        write-host "Databasename" $pfcCompany.Databasename
        write-host "UserName:" $pfcCompany.UserName
    }

    #If company is not connected - stops the script
    if ($pfcCompany.SapCompany.Connected -ne 1) {
        write-host -backgroundcolor yellow -foregroundcolor black "Company is not connected"
        $logJobs.endSubtask('Connection', 'F', 'Company is not connected');
        return 
    }
    #If company is connected to wrong database - stops the script
    if ($pfcCompany.SapCompany.CompanyDB -ne $xmlConnection.Database) {
        write-host -backgroundcolor yellow -foregroundcolor black "Company is connected to wrong database";
        $logJobs.endSubtask('Connection', 'F', 'Company is connected to wrong database');
        return;
    }
    $logJobs.endSubtask('Connection', 'S', '');


    function openItemDetailsForm ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Open Item Details', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Open Item Details:' -NoNewline;
        $xmlOpenItemDetails = $UIConfigXml.SelectSingleNode([string]::Format("ItemDetails"));
        $repeatOpenForm = [int] $xmlOpenItemDetails.repeatOpenForm;
        
        [CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
        for($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Open Item Details Form');
                $formOpenMenu = $app.Menus.Item('CT_PF_1'); 
                $formOpenMenu.Activate();        
                $log.endSubtask('Open Item Details Form', 'S', '');
                $form = $app.Forms.ActiveForm()
                $form.Close();       
            }
            catch {
                $err = $_.Exception.Message;
                $log.endSubtask('Open Item Details Form', 'F', $err);
                continue;
            }
        }
    }

    function loadItemDetails ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Open Item Details', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Load Item Details:' -NoNewline;
        $xmlOpenItemDetails = $UIConfigXml.SelectSingleNode([string]::Format("ItemDetails"));
        $recordsToGoThrough = [int] $xmlOpenItemDetails.recordsToGoThrough;
        $next = $app.Menus.Item('1288');
        [CTProgress] $progress = New-Object CTProgress ($recordsToGoThrough);
        try {
            $firstItemCode = $ItemsDictionary['MAKE'][0].ItemCode;
            $log.startSubtask('Open Item Details Form');
            $formOpenMenu = $app.Menus.Item('CT_PF_1'); 
            $formOpenMenu.Activate();        
            $form = $app.Forms.ActiveForm       
            $log.endSubtask('Open Item Details Form', 'S', '');
        }
        catch {
            $err = $_.Exception.Message;
            $log.endSubtask('Open Item Details Form', 'F', $err);
            continue;
        }
        for($iRecord = 0; $iRecord -lt $recordsToGoThrough; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Item Details Load Data');
                if($iRecord -eq 0){

                    $ItemCodeInputField = $form.Items('idtItCTbx');
                    $ItemCodeInputField.Specific.String = [string]$firstItemCode;
                    $app.SendKeys('{ENTER}');
                } else {
                    $next.Activate();
                }
                $log.endSubtask('Item Details Load Data', 'S', '');
            }
            catch {
                $log.endSubtask('Item Details Load Data', 'F', $err);
                continue;
            }
        }
        $form.Close();
        
    }

    function openBOMForm ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Open BOM', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Open BOM:' -NoNewline;
        $xmlOpenBOM = $UIConfigXml.SelectSingleNode([string]::Format("BOM"));
        $repeatOpenForm = [int] $xmlOpenBOM.repeatOpenForm;
        
        [CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
        for($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Open BOM Form');
                $formOpenMenu = $app.Menus.Item('CT_PF_2'); 
                $formOpenMenu.Activate();        
                $log.endSubtask('Open BOM Form', 'S', '');
                $form = $app.Forms.ActiveForm()
                $form.Close();       
            }
            catch {
                $err = $_.Exception.Message;
                $log.endSubtask('Open BOM Form', 'F', $err);
                continue;
            }
        }
    }

    function loadBOM ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Load BOMs', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Load BOM:' -NoNewline;
        $xmlOpenBOM = $UIConfigXml.SelectSingleNode([string]::Format("BOM"));
        $recordsToGoThrough = [int] $xmlOpenBOM.recordsToGoThrough;
        $next = $app.Menus.Item('1288');
        $find = $app.Menus.Item('1281');

        [CTProgress] $progress = New-Object CTProgress ($recordsToGoThrough);
        try {
            $firstItemCode = $ItemsDictionary['MAKE'][0].ItemCode;
            $firstRevision = $ItemsDictionary['MAKE'][0].Revisions[0]
            $log.startSubtask('Open BOM Form');
            $formOpenMenu = $app.Menus.Item('CT_PF_2'); 
            $formOpenMenu.Activate();      
            if($find.Enabled -eq $true) {
                $find.Activate();
            }
            $form = $app.Forms.ActiveForm       
            $log.endSubtask('Open BOM Form', 'S', '');
        }
        catch {
            $err = $_.Exception.Message;
            $log.endSubtask('Open BOM Form', 'F', $err);
            continue;
        }
        for($iRecord = 0; $iRecord -lt $recordsToGoThrough; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Load BOM Data');
                if($iRecord -eq 0){

                    $ItemCodeInputField = $form.Items('7');
                    $RevisionInputField = $form.Items('13');
                    $ItemCodeInputField.Specific.String = [string]$firstItemCode;
                    $RevisionInputField.Specific.String = [string]$firstRevision;

                    $app.SendKeys('{ENTER}');
                } else {
                    $next.Activate();
                }
                $log.endSubtask('Load BOM Data', 'S', '');
            }
            catch {
                $log.endSubtask('Load BOM Data', 'F', $err);
                continue;
            }
        }
        $form.Close();
        
    }

    
    function openProductionProcessForm ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Open Production Process', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Open Production Process:' -NoNewline;
        $xmlOpenProductionProcess = $UIConfigXml.SelectSingleNode([string]::Format("ProductionProcess"));
        $repeatOpenForm = [int] $xmlOpenProductionProcess.repeatOpenForm;
        
        [CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
        for($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Open Production Process Form');
                $formOpenMenu = $app.Menus.Item('CT_PF_81'); 
                $formOpenMenu.Activate();        
                $log.endSubtask('Open Production Process Form', 'S', '');
                $form = $app.Forms.ActiveForm()
                $form.Close();       
            }
            catch {
                $err = $_.Exception.Message;
                $log.endSubtask('Open Production Process Form', 'F', $err);
                continue;
            }
        }
    }

    function loadProductionProcess ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Load Production Processes', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Load Production Processes:' -NoNewline;
        $xmlOpenProductionProcess = $UIConfigXml.SelectSingleNode([string]::Format("ProductionProcess"));
        $recordsToGoThrough = [int] $xmlOpenProductionProcess.recordsToGoThrough;
        $next = $app.Menus.Item('1288');
        $find = $app.Menus.Item('1281');

        [CTProgress] $progress = New-Object CTProgress ($recordsToGoThrough);
        try {
            $firstItemCode = $ItemsDictionary['MAKE'][0].ItemCode;
            $firstRevision = $ItemsDictionary['MAKE'][0].Revisions[0]
            $log.startSubtask('Open Production Process Form');
            $formOpenMenu = $app.Menus.Item('CT_PF_81'); 
            $formOpenMenu.Activate();      
            if($find.Enabled -eq $true) {
                $find.Activate();
            }
            $form = $app.Forms.ActiveForm       
            $log.endSubtask('Open Production Process Form', 'S', '');
        }
        catch {
            $err = $_.Exception.Message;
            $log.endSubtask('Open Production Process Form', 'F', $err);
            continue;
        }
        for($iRecord = 0; $iRecord -lt $recordsToGoThrough; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Load Production Process Data');
                if($iRecord -eq 0){

                    $ItemCodeInputField = $form.Items('7');
                    $RevisionInputField = $form.Items('RevNameTbx');
                    $ItemCodeInputField.Specific.String = [string]$firstItemCode;
                    $RevisionInputField.Specific.String = [string]$firstRevision;

                    $app.SendKeys('{ENTER}');
                } else {
                    $next.Activate();
                }
                $log.endSubtask('Load Production Process Data', 'S', '');
            }
            catch {
                $log.endSubtask('Load Production Process Data', 'F', $err);
                continue;
            }
        }
        $form.Close();
        
    }

    function openResourceForm ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Open Resources', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Open Production Process:' -NoNewline;
        $xmlOpenResource = $UIConfigXml.SelectSingleNode([string]::Format("Resource"));
        $repeatOpenForm = [int] $xmlOpenResource.repeatOpenForm;
        
        [CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
        for($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Open Resource Form');
                $formOpenMenu = $app.Menus.Item('CT_PF_12'); 
                $formOpenMenu.Activate();        
                $log.endSubtask('Open Resource Form', 'S', '');
                $form = $app.Forms.ActiveForm()
                $form.Close();       
            }
            catch {
                $err = $_.Exception.Message;
                $log.endSubtask('Open Resource Form', 'F', $err);
                continue;
            }
        }
    }

    function loadResources ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Load Production Processes', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Load Production Processes:' -NoNewline;
        $xmlOpenResource = $UIConfigXml.SelectSingleNode([string]::Format("Resource"));
        $recordsToGoThrough = [int] $xmlOpenResource.recordsToGoThrough;
        $next = $app.Menus.Item('1288');
        $find = $app.Menus.Item('1281');

        [CTProgress] $progress = New-Object CTProgress ($recordsToGoThrough);
        try {
            $firstResourceCode = $ResourcesList[0]
            $log.startSubtask('Open Resource Form');
            $formOpenMenu = $app.Menus.Item('CT_PF_12'); 
            $formOpenMenu.Activate();      
            if($find.Enabled -eq $true) {
                $find.Activate();
            }
            $form = $app.Forms.ActiveForm       
            $log.endSubtask('Open Resource Form', 'S', '');
        }
        catch {
            $err = $_.Exception.Message;
            $log.endSubtask('Open Resource Form', 'F', $err);
            continue;
        }
        for($iRecord = 0; $iRecord -lt $recordsToGoThrough; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Load Resource Data');
                if($iRecord -eq 0){
                    $ResourceCodeInputField = $form.Items('rscCodBox');
                    $ResourceCodeInputField.Specific.String = [string]$firstResourceCode;
                    $app.SendKeys('{ENTER}');
                } else {
                    $next.Activate();
                }
                $log.endSubtask('Load Resource Data', 'S', '');
            }
            catch {
                $log.endSubtask('Load Resource Data', 'F', $err);
                continue;
            }
        }
        $form.Close();
        
    }

    function openOperationForm ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Open Operations', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Open Operations:' -NoNewline;
        $xmlOpenOperation = $UIConfigXml.SelectSingleNode([string]::Format("Operation"));
        $repeatOpenForm = [int] $xmlOpenOperation.repeatOpenForm;
        
        [CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
        for($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Open Operation Form');
                $formOpenMenu = $app.Menus.Item('CT_PF_14'); 
                $formOpenMenu.Activate();        
                $log.endSubtask('Open Operation Form', 'S', '');
                $form = $app.Forms.ActiveForm()
                $form.Close();       
            }
            catch {
                $err = $_.Exception.Message;
                $log.endSubtask('Open Operation Form', 'F', $err);
                continue;
            }
        }
    }

    function loadOperations ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Load Operations', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Load Operations:' -NoNewline;
        $xmlOpenOperation = $UIConfigXml.SelectSingleNode([string]::Format("Operation"));
        $recordsToGoThrough = [int] $xmlOpenOperation.recordsToGoThrough;
        $next = $app.Menus.Item('1288');
        $find = $app.Menus.Item('1281');

        [CTProgress] $progress = New-Object CTProgress ($recordsToGoThrough);
        try {
            $firstOprationCode = $OperationsList[0]
            $log.startSubtask('Open Operation Form');
            $formOpenMenu = $app.Menus.Item('CT_PF_14'); 
            $formOpenMenu.Activate();      
            if($find.Enabled -eq $true) {
                $find.Activate();
            }
            $form = $app.Forms.ActiveForm       
            $log.endSubtask('Open Operation Form', 'S', '');
        }
        catch {
            $err = $_.Exception.Message;
            $log.endSubtask('Open Operation Form', 'F', $err);
            continue;
        }
        for($iRecord = 0; $iRecord -lt $recordsToGoThrough; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Load Operation Data');
                if($iRecord -eq 0){
                    $OperationCodeInputField = $form.Items('oprCodBox');
                    $OperationCodeInputField.Specific.String = [string]$firstOprationCode;
                    $app.SendKeys('{ENTER}');
                } else {
                    $next.Activate();
                }
                $log.endSubtask('Load Operation Data', 'S', '');
            }
            catch {
                $log.endSubtask('Load Operation Data', 'F', $err);
                continue;
            }
        }
        $form.Close();
        
    }

    function openRoutingForm ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Open Routings', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Open Routings:' -NoNewline;
        $xmlOpenRouting = $UIConfigXml.SelectSingleNode([string]::Format("Routing"));
        $repeatOpenForm = [int] $xmlOpenRouting.repeatOpenForm;
        
        [CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
        for($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Open Routing Form');
                $formOpenMenu = $app.Menus.Item('CT_PF_13'); 
                $formOpenMenu.Activate();        
                $log.endSubtask('Open Routing Form', 'S', '');
                $form = $app.Forms.ActiveForm()
                $form.Close();       
            }
            catch {
                $err = $_.Exception.Message;
                $log.endSubtask('Open Routing Form', 'F', $err);
                continue;
            }
        }
    }

    function loadRoutings ($app) {
        [CTLogger] $log = New-Object CTLogger ('UI', 'Load Routings', $RESULT_FILE)
        Write-Host '';
        Write-Host 'Load Routings:' -NoNewline;
        $xmlOpenRouting = $UIConfigXml.SelectSingleNode([string]::Format("Routing"));
        $recordsToGoThrough = [int] $xmlOpenRouting.recordsToGoThrough;
        $next = $app.Menus.Item('1288');
        $find = $app.Menus.Item('1281');

        [CTProgress] $progress = New-Object CTProgress ($recordsToGoThrough);
        try {
            $firstRoutingCode = $RoutingsList[0]
            $log.startSubtask('Open Routing Form');
            $formOpenMenu = $app.Menus.Item('CT_PF_13'); 
            $formOpenMenu.Activate();      
            if($find.Enabled -eq $true) {
                $find.Activate();
            }
            $form = $app.Forms.ActiveForm       
            $log.endSubtask('Open Routing Form', 'S', '');
        }
        catch {
            $err = $_.Exception.Message;
            $log.endSubtask('Open Routing Form', 'F', $err);
            continue;
        }
        for($iRecord = 0; $iRecord -lt $recordsToGoThrough; $iRecord++){
            try {
                $progress.next();
                $log.startSubtask('Load Routing Data');
                if($iRecord -eq 0){
                    $RoutingCodeInputField = $form.Items('rtgCodBox');
                    $RoutingCodeInputField.Specific.String = [string]$firstRoutingCode;
                    $app.SendKeys('{ENTER}');
                } else {
                    $next.Activate();
                }
                $log.endSubtask('Load Routing Data', 'S', '');
            }
            catch {
                $log.endSubtask('Load Routing Data', 'F', $err);
                continue;
            }
        }
        $form.Close();
        
    }

   openItemDetailsForm $app;

   loadItemDetails $app;

    openBOMForm $app;

    loadBOM $app;

    openProductionProcessForm $app;

    loadProductionProcess $app;

    openResourceForm $app;

    loadResources $app;

    openOperationForm $app;

    loadOperations $app;
    
    openRoutingForm $app;

    loadRoutings $app;

    $logJobs.endSubtask('Get', 'S', '');
}

function saveTestConfiguration(){

    [CTProgress] $progress = New-Object CTProgress (10);
    Write-Host 'Checking enviroment:' -NoNewline
    
    Add-Content -path $RESULT_FILE_CONF ([string]::Format("Test started at: {0}",(Get-Date)));

    Add-Content -Path $RESULT_FILE_CONF '';
    $os = Get-Ciminstance Win32_OperatingSystem;
    Add-Content -Path $RESULT_FILE_CONF ( [string]::Format("Total Memory: {0} GB",[int]($os.TotalVisibleMemorySize/1mb)) );
    Add-Content -Path $RESULT_FILE_CONF ( [string]::Format("Free Memory: {0} GB",[math]::Round($os.FreePhysicalMemory/1mb,2)) );
    $progress.next();

    Add-Content -Path $RESULT_FILE_CONF '';
    $processor = Get-WmiObject win32_processor
    Add-Content -Path $RESULT_FILE_CONF ( [string]::Format("Processor: {0}",$processor.Name) );
    Add-Content -Path $RESULT_FILE_CONF ( [string]::Format("Processor average usage: {0} %",($processor|Measure-Object -property LoadPercentage -Average | Select Average).Average) );

    Add-Content -Path $RESULT_FILE_CONF '';
    $progress.next();

    # TestConfig.xml 
    Add-Content -Path $RESULT_FILE_CONF 'TestConfig.xml:'
    Add-Content -path $RESULT_FILE_CONF $TestConfigXml.InnerXml;
    Add-Content -Path $RESULT_FILE_CONF ''
    $progress.next();

    # Connection.xml
    Add-Content -Path $RESULT_FILE_CONF 'Connection.xml:'
    Add-Content -path $RESULT_FILE_CONF $connectionConfigXml.InnerXml;
    Add-Content -Path $RESULT_FILE_CONF ''
    $sqlServer = ($xmlConnection.SQLServer).Split(':')[0]
    $licenseServer = ($xmlConnection.LicenseServer).Split(':')[0]
    Add-Content -Path $RESULT_FILE_CONF '';
    $progress.next();

    $pingToDbServer = Test-Connection $sqlServer -Count 20
    $progress.next();
    $pingToDbServer += Test-Connection $sqlServer -Count 20
    $progress.next();
    $pingToDbServer += Test-Connection $sqlServer -Count 20
    
    Add-Content -Path $RESULT_FILE_CONF 'Ping Database Server:'
    foreach($pingResponse in  $pingToDbServer) {
        Add-Content -Path $RESULT_FILE_CONF ([string]::Format("Source:{0}, Destination:{1}, IPV4Address:{2}, IPV6Address{3}, ResponseTime: {4}",
        $pingResponse.PSComputerName,$pingResponse.Address ,$pingResponse.IPV4Address,$pingResponse.IPV6Address,$pingResponse.ResponseTime ))    
    }
    Add-Content -Path $RESULT_FILE_CONF '';
    $progress.next();

    $pingToDbLicenseServer = Test-Connection $licenseServer -Count 20
    $progress.next();
    $pingToDbLicenseServer += Test-Connection $licenseServer -Count 20
    $progress.next();
    $pingToDbLicenseServer += Test-Connection $licenseServer -Count 20
    Add-Content -Path $RESULT_FILE_CONF 'Ping License Server:'
    foreach($pingResponse in  $pingToDbLicenseServer) {
        Add-Content -Path $RESULT_FILE_CONF ([string]::Format("Source:{0}, Destination:{1}, IPV4Address:{2}, IPV6Address{3}, ResponseTime: {4}",
        $pingResponse.PSComputerName,$pingResponse.Address ,$pingResponse.IPV4Address,$pingResponse.IPV6Address,$pingResponse.ResponseTime ))    
    }
    Add-Content -Path $RESULT_FILE_CONF '';
    $progress.next();

    

    
    
}

saveTestConfiguration ;
Write-Host '';
Imports ;
write-host '';
UITests ;


 