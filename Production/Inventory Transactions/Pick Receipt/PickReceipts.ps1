#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Pick Receipts
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.170) PL: 07 R1 Pre-Release (64-bit)
# Description:
#      Import Pick Receipt documents. Script add or update Pick Receipts.
#      You need to have all requred files for import.
# Warning:
#   It's recommended run script when all users all disconnected.
#   Before running this script please do database backup.
# Troubleshooting:
#   https://connect.computec.pl/display/PF100EN/PowerShell+FAQ
#   https://connect.computec.pl/display/PF930EN/PowerShell+FAQ
# Script source:
#   https://code.computec.pl/repos?visibility=public
########################################################################
Write-Host -backgroundcolor Yellow -foregroundcolor DarkBlue ("Script Version:" + $SCRIPT_VERSION)
#endregion

#region #PF API library usage
# You need to check in what architecture PowerShell ISE is running (x64 or x86),
# you need run PowerShell ISE in the same architecture like PF API is installed (check in Windows -> Programs & Features)
# Examples:
#     SAP Client x64 + PF x64 installed on DB/Company => PF API x64 => Windows PowerShell ISE
#     SAP Client x86 + PF x86 installed on DB/Company => PF API x86 => Windows PowerShell ISE x86

[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")
#endregion

#region Script parameters

$csvImportCatalog = $PSScriptRoot + "\"
#If you are using lower version of PowerShell than PowerShell 4.0 you can use static path
# $csvImportCatalog = "C:\PS\PF\";

$csvPickReceiptsPath = -join ($csvImportCatalog, "PickReceipts.csv")
$csvPickReceiptsLinesPath = -join ($csvImportCatalog, "PickReceiptsLines.csv")
$csvPickReceiptsBinsBatchesPath = -join ($csvImportCatalog, "PickReceiptsBinBatch.csv")

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
    write-host -backgroundcolor yellow -foregroundcolor black "Company is not connected";
    return;
}

try {

    #Data loading from a csv file
    write-host ""

    [array]$pickReceipts = Import-Csv -Delimiter ';' -Path $csvPickReceiptsPath
    [array]$pickReceiptsLines = Import-Csv -Delimiter ';' -Path $csvPickReceiptsLinesPath

    [array]$pickReceiptsBinsBatches = $null;
    if ((Test-Path -Path $csvPickReceiptsBinsBatchesPath -PathType leaf) -eq $true) {
        [array]$pickReceiptsBinsBatches = Import-Csv -Delimiter ';' -Path $csvPickReceiptsBinsBatchesPath
    }
    else {
        write-host "Pick Receipts Bins Batches - csv not available."
    }

    write-Host 'Preparing data: '
    $totalRows = $pickReceipts.Count + $pickReceiptsLines.Count + $pickReceiptsBinsBatches.Count 

    $prList = New-Object 'System.Collections.Generic.List[array]'

    $dictionaryPRLines = New-Object 'System.Collections.Generic.Dictionary[string,psobject]'
    $dictionaryPRBinsBatches = New-Object 'System.Collections.Generic.Dictionary[string,psobject]'
    

    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;

    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $pickReceipts) {
        $key = $row.MORDocEntry;
        $progressItterator++;
        $progress = [math]::Round(($progressItterator * 100) / $total);
        if ($progress -gt $beforeProgress) {
            Write-Host $progress"% " -NoNewline
            $beforeProgress = $progress
        }
        $prList.Add([array]$row);
    }

    foreach ($row in $pickReceiptsLines) {
        $key = $row.MORDocEntry;
        $progressItterator++;
        $progress = [math]::Round(($progressItterator * 100) / $total);
        if ($progress -gt $beforeProgress) {
            Write-Host $progress"% " -NoNewline
            $beforeProgress = $progress
        }

        if ($dictionaryPRLines.ContainsKey($key) -eq $false) {
            $dictionaryPRLines[$key] = [psobject]@{
                existingLines = New-Object 'System.Collections.Generic.Dictionary[int,array]';
                newLines      = New-Object 'System.Collections.Generic.Dictionary[string,array]';
            }
        }
        
        if ($row.PickLineNum -gt 0) {
            $dictionaryPRLines[$key].existingLines.Add($row.PickLineNum, [array]$row);
        }
        else {
            $PRLineKey = [string]::Format("{0}___{1}", [string]$row.MORLineNum, [string]$row.MORBaseType);
            $dictionaryPRLines[$key].newLines.Add($PRLineKey, [array]$row);
        }
    }

    foreach ($row in $pickReceiptsBinsBatches) {
        $key = $row.MORDocEntry + '___' + $row.MORBaseType + '___' + $row.MORLineNum;
        $progressItterator++;
        $progress = [math]::Round(($progressItterator * 100) / $total);
        if ($progress -gt $beforeProgress) {
            Write-Host $progress"% " -NoNewline
            $beforeProgress = $progress
        }
        
        if ($dictionaryPRBinsBatches.ContainsKey($key) -eq $false) {
            $dictionaryPRBinsBatches[$key] = [psobject]@{
                Batches  = New-Object 'System.Collections.Generic.Dictionary[string,psobject]';
                Bins     = New-Object 'System.Collections.Generic.List[psobject]';
                Quantity = 0;
            }
        }
        $dictionaryPRBinsBatches[$key].Quantity += $row.Quantity;    

        if ([string]::IsNullOrWhiteSpace($row.Batch) -eq $false) {
            if ($dictionaryPRBinsBatches[$key].Batches.ContainsKey($row.Batch) -eq $false) {
                $dictionaryPRBinsBatches[$key].Batches.Add($row.Batch, [psobject]@{
                        Batch        = $row.Batch;
                        Details      = $row;
                        Quantity     = 0;
                        BatchLineNum = -1;
                    }
                );
            }
            $dictionaryPRBinsBatches[$key].Batches[$row.Batch].Quantity += $row.Quantity;
        }

        if ([string]::IsNullOrWhiteSpace($row.BinAbsEntry) -eq $false) {
            
            $dictionaryPRBinsBatches[$key].Bins.Add([psobject]@{
                    BinAbsEntry = $row.BinAbsEntry;
                    Batch       = [string] $row.Batch;
                    Quantity    = $row.Quantity;
                });
        }
    }
    
    Write-Host '';

    Write-Host 'Adding/updating data: ';
    if ($prList.Count -gt 1) {
        $total = $prList.Count;
    }
    else {
        $total = 1;
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;

    foreach ($csvItem in $prList) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $dictionaryKey = $csvItem.MORDocEntry;
            $pfcCompany.StartTransaction();
            $pickReceipt = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::PickReceipt)

            $docentryPickReceipt = $csvItem.PickDocEntry;
            $newPickReceipt = ([string]::IsNullOrWhiteSpace($docentryPickReceipt));
            
            if ($newPickReceipt) {
                $pickReceiptAction = $pfcCompany.CreatePFAction([CompuTec.ProcessForce.API.Core.ActionType]::CreatePickReceiptForProductionReceipt);
                $dummy = $pickReceiptAction.AddManufacturingOrderDocEntry($csvItem.MORDocEntry);
                $pickReceiptAction.ReceiptType = [CompuTec.ProcessForce.API.Actions.CreatePickReceiptForProductionReceipt.PickOrderdReceiptType]::All
                $docentryPickReceipt = 0;
                $dummy = $pickReceiptAction.DoAction([ref] $docentryPickReceipt);
            }

            try {
                $result = $pickReceipt.GetByKey($docentryPickReceipt);
                # $result = $pickReceipt.GetByKey(87);
                if ($result -ne 0) {
                    $err = [string]$pfcCompany.GetLastErrorDescription();
                    Throw [System.Exception] ($err);
                }
            }
            catch {
                $err = $_.Exception.Message;
                $ms = [string]::Format("Exception when loading Pick Receipt with DocEntry: {0}. Details: {1}", [string]$docentryPickReceipt, [string]$err);
                throw [System.Exception]($ms);
            }

            $status = [string]$csvItem.Status;
            switch ($status) {
                "" { break; }
                "O" { $pickReceipt.Status = [CompuTec.ProcessForce.API.Enumerators.PickStatus]::Open; break; }
                "S" { $pickReceipt.Status = [CompuTec.ProcessForce.API.Enumerators.PickStatus]::Started; break; }
                "C" { $pickReceipt.Status = [CompuTec.ProcessForce.API.Enumerators.PickStatus]::Closed; break; }
                Default {
                    $ms = [string]::Format("Incorrect status code: {0}. Possible values are: '','O','S','C'", $status);
                    throw [System.Exception]($ms);
                }
            }

            $pickReceipt.U_Ref2 = [string]$csvItem.Ref2
            $pickReceipt.U_Remarks = [string]$csvItem.Remarks
            if ($csvItem.Employee -gt 0) {
                $pickReceipt.U_Employee = $csvItem.Employee;
            }

            if ($dictionaryPRLines[$dictionaryKey].Count -gt 0) {
                
                $prItems = $dictionaryPRLines[$dictionaryKey];
                $linesToBeRemoved = New-Object 'System.Collections.Generic.List[int]';
                $linesToBeUpdated = New-Object 'System.Collections.Generic.Dictionary[string,psobject]';

                $itemIndex = 0;
                foreach ($line in $pickReceipt.RequiredItems) {
                    $LineNum = $line.U_LineNum;
                    $BaseLineNum = $line.U_BaseLineNo;
                    $BaseRef = $line.U_BaseRef;
                    $PRLineKey = [string]::Format("{0}___{1}", [string]$BaseLineNum, [string]$BaseRef);

                    if ($prItems.existingLines.ContainsKey($LineNum) -eq $true) {
                        $linesToBeUpdated.Add($itemIndex, [psobject]@{
                                PickUpdate = $true;
                                currentPos = $itemIndex;
                                item       = $prItems.existingLines[$LineNum]
                            });
                    }
                    elseif ($prItems.newLines.ContainsKey($PRLineKey) -eq $true) {
                        $linesToBeUpdated.Add($itemIndex, [psobject]@{
                                PickUpdate = $false;
                                currentPos = $itemIndex;
                                item       = $prItems.newLines[$PRLineKey]
                            });
                    }
                    else {
                        $linesToBeRemoved.Add($itemIndex);
                    }
                    $itemIndex++;
                }

                #update 
                foreach ($linesToBeUpdatedIndex in $linesToBeUpdated.Keys) {
                    $updateItem = $linesToBeUpdated[$linesToBeUpdatedIndex];
                    $prItem = $updateItem.item;
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
                    $dummy = $pickReceipt.RequiredItems.SetCurrentLine($updateItem.currentPos);
                    $line = $pickReceipt.RequiredItems;

                    #if this is only pick update MORDocEntry, MORLineNum and MORBaseType are not required
                    if ($updateItem.PickUpdate -eq $false) {
                        if ($line.U_BaseEntry -ne $prItem.MORDocEntry) {
                            throw [System.Exception](([string]::Format("Incorrect Pick Base Entry {0} and MOR DocEntry {1}", [string]$line.U_BaseEntry, [string]$prItem.MORDocEntry)));
                        }
                        if ($line.U_BaseLineNo -ne $prItem.MORLineNum) {
                            throw [System.Exception](([string]::Format("Incorrect Pick Base Line No {0} and MOR Line Num {1}", [string]$line.U_BaseLineNo, [string]$prItem.MORLineNum)));
                        }
                        if ($line.U_BaseRef -ne $prItem.MORBaseType) {
                            throw [System.Exception](([string]::Format("Incorrect Pick Base Ref {0} and MOR Base Type {1}", [string]$line.U_BaseRef, [string]$prItem.MORBaseType)));
                        }
                    }
                    if ($line.U_ItemCode -ne $prItem.ItemCode) {
                        throw [System.Exception](([string]::Format("Incorrect Pick Item Code {0} and MOR Item Code {1}", [string]$line.U_ItemCode, [string]$prItem.ItemCode)));
                    }
                    if ($line.U_RevisionCode -ne $prItem.Revision) {
                        throw [System.Exception](([string]::Format("Incorrect Pick Revision {0} and MOR Revision {1}", [string]$line.U_RevisionCode, [string]$prItem.Revision)));
                    }

                    $pickReceipt.RequiredItems.U_PickedQty = $prItem.PickedQty;
                    if ([string]::IsNullOrWhiteSpace($prItem.DstWhsCode) -eq $false) {
                        $pickReceipt.RequiredItems.U_DstWhsCode = $prItem.DstWhsCode;
                    }
                    if ($prItem.Price -gt 0) {
                        $pickReceipt.RequiredItems.U_Price = $prItem.Price;
                    }
                    if ([string]::IsNullOrWhiteSpace($prItem.Project) -eq $false) {
                        $pickReceipt.RequiredItems.U_Project = $prItem.Project;
                    }
                    if ([string]::IsNullOrWhiteSpace($prItem.OcrCode) -eq $false) {
                        $pickReceipt.RequiredItems.U_OcrCode = $prItem.OcrCode;
                    }
                    if ([string]::IsNullOrWhiteSpace($prItem.OcrCode2) -eq $false) {
                        $pickReceipt.RequiredItems.U_OcrCode2 = $prItem.OcrCode2;
                    }
                    if ([string]::IsNullOrWhiteSpace($prItem.OcrCode3) -eq $false) {
                        $pickReceipt.RequiredItems.U_OcrCode3 = $prItem.OcrCode3;
                    }
                    if ([string]::IsNullOrWhiteSpace($prItem.OcrCode4) -eq $false) {
                        $pickReceipt.RequiredItems.U_OcrCode4 = $prItem.OcrCode4;
                    }
                    if ([string]::IsNullOrWhiteSpace($prItem.OcrCode5) -eq $false) {
                        $pickReceipt.RequiredItems.U_OcrCode5 = $prItem.OcrCode5;
                    }
                    
                    $keyLineNum = ([string] $line.U_BaseEntry) + '___' + $line.U_BaseRef + '___' + $line.U_BaseLineNo;
                    $batches = $dictionaryPRBinsBatches[$keyLineNum].Batches;
                    $bins = $dictionaryPRBinsBatches[$keyLineNum].Bins;
                    $dummy = $pickReceipt.PickedItems.SetCurrentLine($pickReceipt.PickedItems.Count - 1);
                    $pickReceiptLineNum = [int]$line.U_LineNum;
                    foreach ($batchKey in $batches.Keys) {
                        $batch = $batches[$batchKey];
                        $details = $batch.Details;
                        $pickReceipt.PickedItems.U_ItemCode = [string]$line.U_ItemCode;
                        $pickReceipt.PickedItems.U_BnDistNumber = [string]$batch.Batch;
                        
                        if ([string]::IsNullOrWhiteSpace($details.BatchAttribute1) -eq $false) {
                            $pickReceipt.PickedItems.U_BnMnfSerial = $details.BatchAttribute1
                        }
                        if ([string]::IsNullOrWhiteSpace($details.BatchAttribute2) -eq $false) {
                            $pickReceipt.PickedItems.U_BnLotNumber = $details.BatchAttribute2
                        }
                        if ([string]::IsNullOrWhiteSpace($details.ExpiryDate) -eq $false) {
                            $pickReceipt.PickedItems.U_BnExpDate = $details.ExpiryDate
                        }
                        if ([string]::IsNullOrWhiteSpace($details.ExpiryTime) -eq $false) {
                            $pickReceipt.PickedItems.U_BnExpTime = $details.ExpiryTime
                        }
                        if ([string]::IsNullOrWhiteSpace($details.MnfDate) -eq $false) {
                            $pickReceipt.PickedItems.U_BnMnfDate = $details.MnfDate
                        }
                        if ([string]::IsNullOrWhiteSpace($details.AdmDate) -eq $false) {
                            $pickReceipt.PickedItems.U_BnInDate = $details.AdmDate
                        }
                        if ([string]::IsNullOrWhiteSpace($details.Location) -eq $false) {
                            $pickReceipt.PickedItems.U_Location = $details.Location
                        }
                        if ([string]::IsNullOrWhiteSpace($details.Details) -eq $false) {
                            $pickReceipt.PickedItems.U_BnNotes = $details.Details
                        }
                        if ([string]::IsNullOrWhiteSpace($details.ConsDate) -eq $false) {
                            $pickReceipt.PickedItems.U_BnConsDate = $details.ConsDate
                        }
                        if ([string]::IsNullOrWhiteSpace($details.WCoDate) -eq $false) {
                            $pickReceipt.PickedItems.U_BnWConsDate = $details.WCoDate
                        }
                        if ([string]::IsNullOrWhiteSpace($details.WExDate) -eq $false) {
                            $pickReceipt.PickedItems.U_BnWExpDate = $details.WExDate
                        }
                        if ([string]::IsNullOrWhiteSpace($details.InDate) -eq $false) {
                            $pickReceipt.PickedItems.U_BnInspDate = $details.InDate
                        }
                        if ([string]::IsNullOrWhiteSpace($details.LstInDate) -eq $false) {
                            $pickReceipt.PickedItems.U_BnLInspDate = $details.LstInDate
                        }
                        if ([string]::IsNullOrWhiteSpace($details.NxtInDate) -eq $false) {
                            $pickReceipt.PickedItems.U_BnNInspDate = $details.NxtInDate
                        }
                        if ([string]::IsNullOrWhiteSpace($details.SupNumber) -eq $false) {
                            $pickReceipt.PickedItems.U_SupNumber = $details.SupNumber
                        }

                        $bnStatus = [string]$details.Status;
                        switch ($bnStatus) {
                            "" {  break; }
                            "R" { $pickReceipt.PickedItems.U_BnStatus = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Released; break; }
                            "L" { $pickReceipt.PickedItems.U_BnStatus = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::Locked; break; }
                            "A" { $pickReceipt.PickedItems.U_BnStatus = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.BatchStatus]::NotAccesible; break; }
                            Default {
                                $ms = [string]::Format("Incorrect status code: {0}. Possible values are: '','R','L','A'", $bnStatus);
                                throw [System.Exception]($ms);
                            }
                        }

                        $qcStatus = [string]$details.QCStatus;
                        switch ($qcStatus) {
                            "" {  break; }
                            "F" { $pickReceipt.PickedItems.U_BnQCStatus = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Failed; break; }
                            "H" { $pickReceipt.PickedItems.U_BnQCStatus = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::OnHold; break; }
                            "I" { $pickReceipt.PickedItems.U_BnQCStatus = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Inspection; break; }
                            "P" { $pickReceipt.PickedItems.U_BnQCStatus = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::Passed; break; }
                            "T" { $pickReceipt.PickedItems.U_BnQCStatus = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::QCTesting; break; }
                            "D" { $pickReceipt.PickedItems.U_BnQCStatus = [CompuTec.ProcessForce.API.Documents._Other_.AdditionalBatchDetails.QCStatus]::ToBeDetermined; break; }
                            Default {
                                $ms = [string]::Format("Incorrect QC Status code: {0}. Possible values are: '','F','H','I','P,'T'", $qcStatus);
                                throw [System.Exception]($ms);
                            }
                        }

                        if ([string]::IsNullOrWhiteSpace($details.Remarks) -eq $false) {
                            $pickReceipt.PickedItems.U_Remarks = $details.Remarks
                        }
                        if ([string]::IsNullOrWhiteSpace($details.Origin) -eq $false) {
                            $pickReceipt.PickedItems.U_Origin = $details.Origin
                        }
                        if ([string]::IsNullOrWhiteSpace($details.BestBefDate) -eq $false) {
                            $pickReceipt.PickedItems.U_BestBefDate = $details.BestBefDate
                        }

                        $pickReceipt.PickedItems.U_Quantity = $batch.Quantity;
                        $pickReceipt.PickedItems.U_ReqItmLn = $pickReceiptLineNum;
                        $batch.BatchLineNum = $pickReceipt.PickedItems.U_LineNum;
                        $dummy = $pickReceipt.PickedItems.Add();

                        #Specify relation Beetween picked and Required Line 
                        $pickReceipt.Relations.SetCurrentLine($pickReceipt.Relations.Count - 1);
                        $pickReceipt.Relations.U_ReqItemLineNo = [int]$line.U_LineNum
                        $pickReceipt.Relations.U_PickItemLineNo = [int]$batch.BatchLineNum;
                        $dummy = $pickReceipt.Relations.Add();
                    }

                    foreach ($bin in $bins) {
                        $batch = $bin.Batch;
                        $BatchLineNum = -1;
                        if ($batches.ContainsKey($batch)) {
                            $BatchLineNum = $batches[$batch].BatchLineNum;
                        }

                        $pickReceipt.BinAllocations.SetCurrentLine($pickReceipt.BinAllocations.Count - 1);
                        $pickReceipt.BinAllocations.U_BinAbsEntry = $bin.BinAbsEntry;
                        $pickReceipt.BinAllocations.U_Quantity = $bin.Quantity;
                        if ($BatchLineNum -gt -1) {
                            $pickReceipt.BinAllocations.U_SnAndBnLine = $BatchLineNum;
                        }
                        $dummy = $pickReceipt.BinAllocations.Add()
                    }
                }

                #delete
                for ($idxD = $linesToBeRemoved.Count - 1; $idxD -ge 0; $idxD--) {
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
                    $dummy = $pickReceipt.RequiredItems.DelRowAtPos($linesToBeRemoved[$idxD]);
                }

                $result = $pickReceipt.Update();
                
                if ($result -lt 0) {    
                    $err = $pfcCompany.GetLastErrorDescription()
                    Throw [System.Exception] ($err)
                }

                if ([CompuTec.ProcessForce.API.Configuration.ConfigurationHolder]::Configuration[$pfcCompany.Token].AutoReceiptEnabled -eq $false) {
                    if ($csvItem.CreateGoodsReceipt -eq "Y") {
                        $pickReceiptGRAction = $pfcCompany.CreatePFAction([CompuTec.ProcessForce.API.Core.ActionType]::CreateGoodsReceiptFromPickReceiptBasedOnProductionReceipt);
                        $pickReceiptGRAction.PickReceiptID = $docentryPickReceipt
                        $grDocEntry = 0;
                        $result = $pickReceiptGRAction.DoAction([ref] $grDocEntry)

                        if ($result -lt 0) {    
                            $err = $pfcCompany.GetLastErrorDescription()
                            Throw [System.Exception] ($err)
                        }
                    }
                }
            }
            
            if ($pfcCompany.InTransaction) {
                $pfcCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Commit);
            }
        }
        Catch {
            $err = $_.Exception.Message;
            
            if ($newPickReceipt -eq $true) {
                $taskMsg = "adding";
            }
            else {
                $taskMsg = "updating"
            }
            $ms = [string]::Format("Error when {0} Pick Receipt for MOR {1} Details: {2}", $taskMsg, $csvItem.MORDocEntry, $err);
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