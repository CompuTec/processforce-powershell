using module .\lib\CTLogger.psm1;
using module .\lib\CTProgress.psm1;

Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")
add-type -Path "C:\PS_MP\MP_PS\SAP\DLL\Interop.SAPbobsCOM.dll"
add-type -Path "C:\PS_MP\MP_PS\SAP\DLL\Interop.SAPbouiCOM.dll"

$ItemsDictionary = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[psobject]]';
$ResourcesList = New-Object 'System.Collections.Generic.List[string]';
$OperationsList = New-Object 'System.Collections.Generic.List[string]';
$RoutingsList = New-Object 'System.Collections.Generic.List[string]';
$ItemsDictionary.Add('MAKE', (New-Object 'System.Collections.Generic.List[psobject]'));
$ItemsDictionary.Add('BUY', (New-Object 'System.Collections.Generic.List[psobject]'));
[xml] $TestConfigXml = Get-Content -Encoding UTF8 .\conf\TestConfig.xml
$MDConfigXml = $TestConfigXml.SelectSingleNode("/CT_CONFIG/MasterData");
$UIConfigXML = $TestConfigXml.SelectSingleNode("/CT_CONFIG/UI");
$RESULT_FILE = $PSScriptRoot + "\Results.csv";
function Imports() {

    
    
    [CTLogger] $logJobs = New-Object CTLogger ('DI', 'Import', $RESULT_FILE)

    #region connection
    [xml] $connectionConfigXml = Get-Content -Encoding UTF8 .\conf\Connection.xml
    $xmlConnection = $connectionConfigXml.SelectSingleNode("/CT_CONFIG/Connection");

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
                    for ($i = 0; $i -lt $count; $i++) {
                        $dummy = $routing.Operations.DelRowAtPos(0);
                    }
                    $count = $routing.OperationResources.Count
                    for ($i = 0; $i -lt $count; $i++) {
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
    $logJobs.endSubtask('Connection', 'S', '');

    $x = $app.Menus.Item('CT_PF_SIDT')
    $next = $app.Menus.Item('1288');
    

    #10 x 
    $x.Activate();
    
    $next.Activate();




    $logJobs.endSubtask('Get', 'S', '');
}

#Imports ;

UITests ;