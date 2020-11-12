#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Pick Orders
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.170) PL: 07 R1 Pre-Release (64-bit)
# Description:
#      Import Pick Order documents. Script add or update Pick Orders.
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

$csvPickOrdersPath = -join ($csvImportCatalog, "PickOrders.csv")
$csvPickOrdersLinesPath = -join ($csvImportCatalog, "PickOrdersLines.csv")
$csvPickOrdersBinsBatchesPath = -join ($csvImportCatalog, "PickOrdersBinBatch.csv")

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

    [array]$PickOrders = Import-Csv -Delimiter ';' -Path $csvPickOrdersPath
    [array]$PickOrdersLines = Import-Csv -Delimiter ';' -Path $csvPickOrdersLinesPath

    [array]$PickOrdersBinsBatches = $null;
    if ((Test-Path -Path $csvPickOrdersBinsBatchesPath -PathType leaf) -eq $true) {
        [array]$PickOrdersBinsBatches = Import-Csv -Delimiter ';' -Path $csvPickOrdersBinsBatchesPath
    }
    else {
        write-host "Pick Orders Bins Batches - csv not available."
    }

    write-Host 'Preparing data: '
    $totalRows = $PickOrders.Count + $PickOrdersLines.Count + $PickOrdersBinsBatches.Count 

    $poList = New-Object 'System.Collections.Generic.List[array]'

    $dictionaryPOLines = New-Object 'System.Collections.Generic.Dictionary[string,psobject]'
    $dictionaryPOBinsBatches = New-Object 'System.Collections.Generic.Dictionary[string,psobject]'
    

    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;

    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $PickOrders) {
        $key = $row.MORDocEntry;
        $progressItterator++;
        $progress = [math]::Round(($progressItterator * 100) / $total);
        if ($progress -gt $beforeProgress) {
            Write-Host $progress"% " -NoNewline
            $beforeProgress = $progress
        }
        $poList.Add([array]$row);
    }

    foreach ($row in $PickOrdersLines) {
        $key = $row.MORDocEntry;
        $progressItterator++;
        $progress = [math]::Round(($progressItterator * 100) / $total);
        if ($progress -gt $beforeProgress) {
            Write-Host $progress"% " -NoNewline
            $beforeProgress = $progress
        }

        if ($dictionaryPOLines.ContainsKey($key) -eq $false) {
            $dictionaryPOLines[$key] = [psobject]@{
                existingLines = New-Object 'System.Collections.Generic.Dictionary[int,array]';
                newLines      = New-Object 'System.Collections.Generic.Dictionary[string,array]';
            }
        }
        
        if ($row.PickLineNum -gt 0) {
            $dictionaryPOLines[$key].existingLines.Add($row.PickLineNum, [array]$row);
        }
        else {
            $dictionaryPOLines[$key].newLines.Add($row.MORLineNum, [array]$row);
        }
    }

    foreach ($row in $PickOrdersBinsBatches) {
        $key = $row.MORDocEntry + '___' + $row.MORLineNum;
        $progressItterator++;
        $progress = [math]::Round(($progressItterator * 100) / $total);
        if ($progress -gt $beforeProgress) {
            Write-Host $progress"% " -NoNewline
            $beforeProgress = $progress
        }
        
        if ($dictionaryPOBinsBatches.ContainsKey($key) -eq $false) {
            $dictionaryPOBinsBatches[$key] = [psobject]@{
                Batches  = New-Object 'System.Collections.Generic.Dictionary[string,psobject]';
                Bins     = New-Object 'System.Collections.Generic.List[psobject]';
                Quantity = 0;
            }
        }
        $dictionaryPOBinsBatches[$key].Quantity += $row.Quantity;    

        if ([string]::IsNullOrWhiteSpace($row.Batch) -eq $false) {
            if ($dictionaryPOBinsBatches[$key].Batches.ContainsKey($row.Batch) -eq $false) {
                $dictionaryPOBinsBatches[$key].Batches.Add($row.Batch, [psobject]@{
                        Batch        = $row.Batch;
                        Details      = $row;
                        Quantity     = 0;
                        BatchLineNum = -1;
                    }
                );
            }
            $dictionaryPOBinsBatches[$key].Batches[$row.Batch].Quantity += $row.Quantity;
        }

        if ([string]::IsNullOrWhiteSpace($row.BinAbsEntry) -eq $false) {
            
            $dictionaryPOBinsBatches[$key].Bins.Add([psobject]@{
                    BinAbsEntry = $row.BinAbsEntry;
                    Batch       = [string] $row.Batch;
                    Quantity    = $row.Quantity;
                });
        }
    }
    
    Write-Host '';

    Write-Host 'Adding/updating data: ';
    if ($poList.Count -gt 1) {
        $total = $poList.Count;
    }
    else {
        $total = 1;
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;

    foreach ($csvItem in $poList) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $dictionaryKey = $csvItem.MORDocEntry;
            $pfcCompany.StartTransaction();
            $pickOrder = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::PickOrder)

            $docentryPickOrder = $csvItem.PickDocEntry;
            $newPickOrder = ([string]::IsNullOrWhiteSpace($docentryPickOrder));
            
            if ($newPickOrder) {
                $pickOrderAction = $pfcCompany.CreatePFAction([CompuTec.ProcessForce.API.Core.ActionType]::CreatePickOrderForProductionIssue);
                $dummy = $pickOrderAction.AddMORDocEntry($csvItem.MORDocEntry);
                $docentryPickOrder = 0;
                $dummy = $pickOrderAction.DoAction([ref] $docentryPickOrder);
            }

            try {
                $result = $pickOrder.GetByKey($docentryPickOrder);
                # $result = $pickOrder.GetByKey(87);
                if ($result -ne 0) {
                    $err = [string]$pfcCompany.GetLastErrorDescription();
                    Throw [System.Exception] ($err);
                }
            }
            catch {
                $err = $_.Exception.Message;
                $ms = [string]::Format("Exception when loading Pick Order with DocEntry: {0}. Details: {1}", [string]$docentryPickOrder, [string]$err);
                throw [System.Exception]($ms);
            }

            $status = [string]$csvItem.Status;
            switch ($status) {
                "" { break; }
                "O" { $pickOrder.Status = [CompuTec.ProcessForce.API.Enumerators.PickStatus]::Open; break; }
                "S" { $pickOrder.Status = [CompuTec.ProcessForce.API.Enumerators.PickStatus]::Started; break; }
                "C" { $pickOrder.Status = [CompuTec.ProcessForce.API.Enumerators.PickStatus]::Closed; break; }
                Default {
                    $ms = [string]::Format("Incorrect status code: {0}. Possible values are: '','O','S','C'", $status);
                    throw [System.Exception]($ms);
                }
            }

            $pickOrder.U_Ref2 = [string]$csvItem.Ref2
            $pickOrder.U_Remarks = [string]$csvItem.Remarks
            if ($csvItem.Employee -gt 0) {
                $pickOrder.U_Employee = $csvItem.Employee;
            }

            if ($dictionaryPOLines[$dictionaryKey].Count -gt 0) {
                
                $poItems = $dictionaryPOLines[$dictionaryKey];
                $linesToBeRemoved = New-Object 'System.Collections.Generic.List[int]';
                $linesToBeUpdated = New-Object 'System.Collections.Generic.Dictionary[string,psobject]';

                $itemIndex = 0;
                foreach ($line in $pickOrder.RequiredItems) {
                    $LineNum = $line.U_LineNum;
                    $BaseLineNum = $line.U_BaseLineNo;
                    $PRLineKey = [string]$BaseLineNum;

                    if ($poItems.existingLines.ContainsKey($LineNum) -eq $true) {
                        $linesToBeUpdated.Add($itemIndex, [psobject]@{
                                PickUpdate = $true;
                                currentPos = $itemIndex;
                                item       = $poItems.existingLines[$LineNum]
                            });
                    }
                    elseif ($poItems.newLines.ContainsKey($PRLineKey) -eq $true) {
                        $linesToBeUpdated.Add($itemIndex, [psobject]@{
                                PickUpdate = $false;
                                currentPos = $itemIndex;
                                item       = $poItems.newLines[$PRLineKey]
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
                    $poItem = $updateItem.item;
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
                    $dummy = $pickOrder.RequiredItems.SetCurrentLine($updateItem.currentPos);
                    $line = $pickOrder.RequiredItems;

                    #if this is only pick update MORDocEntry, MORLineNum and MORBaseType are not required
                    if ($updateItem.PickUpdate -eq $false) {
                        if ($line.U_BaseEntry -ne $poItem.MORDocEntry) {
                            throw [System.Exception](([string]::Format("Incorrect Pick Base Entry {0} and MOR DocEntry {1}", [string]$line.U_BaseEntry, [string]$poItem.MORDocEntry)));
                        }
                        if ($line.U_BaseLineNo -ne $poItem.MORLineNum) {
                            throw [System.Exception](([string]::Format("Incorrect Pick Base Line No {0} and MOR Line Num {1}", [string]$line.U_BaseLineNo, [string]$poItem.MORLineNum)));
                        }
                    }
                    if ($line.U_ItemCode -ne $poItem.ItemCode) {
                        throw [System.Exception](([string]::Format("Incorrect Pick Item Code {0} and MOR Item Code {1}", [string]$line.U_ItemCode, [string]$poItem.ItemCode)));
                    }
                    if ($line.U_Revision -ne $poItem.Revision) {
                        throw [System.Exception](([string]::Format("Incorrect Pick Revision {0} and MOR Revision {1}", [string]$line.U_Revision, [string]$poItem.Revision)));
                    }

                    $pickOrder.RequiredItems.U_PickedQty = $poItem.PickedQty;
                    if ([string]::IsNullOrWhiteSpace($poItem.SrcWhsCode) -eq $false) {
                        $pickOrder.RequiredItems.U_SrcWhsCode = $poItem.SrcWhsCode;
                    }
                    
                    if ([string]::IsNullOrWhiteSpace($poItem.Project) -eq $false) {
                        $pickOrder.RequiredItems.U_Project = $poItem.Project;
                    }
                    if ([string]::IsNullOrWhiteSpace($poItem.OcrCode) -eq $false) {
                        $pickOrder.RequiredItems.U_OcrCode = $poItem.OcrCode;
                    }
                    if ([string]::IsNullOrWhiteSpace($poItem.OcrCode2) -eq $false) {
                        $pickOrder.RequiredItems.U_OcrCode2 = $poItem.OcrCode2;
                    }
                    if ([string]::IsNullOrWhiteSpace($poItem.OcrCode3) -eq $false) {
                        $pickOrder.RequiredItems.U_OcrCode3 = $poItem.OcrCode3;
                    }
                    if ([string]::IsNullOrWhiteSpace($poItem.OcrCode4) -eq $false) {
                        $pickOrder.RequiredItems.U_OcrCode4 = $poItem.OcrCode4;
                    }
                    if ([string]::IsNullOrWhiteSpace($poItem.OcrCode5) -eq $false) {
                        $pickOrder.RequiredItems.U_OcrCode5 = $poItem.OcrCode5;
                    }
                    
                    $keyLineNum = ([string] $line.U_BaseEntry) + '___' + $line.U_BaseLineNo;
                    $batches = $dictionaryPOBinsBatches[$keyLineNum].Batches;
                    $bins = $dictionaryPOBinsBatches[$keyLineNum].Bins;
                    $dummy = $pickOrder.PickedItems.SetCurrentLine($pickOrder.PickedItems.Count - 1);
                    $pickOrderLineNum = [int]$line.U_LineNum;
                    foreach ($batchKey in $batches.Keys) {
                        $batch = $batches[$batchKey];
                       
                        $pickOrder.PickedItems.U_ItemCode = [string]$line.U_ItemCode;
                        $pickOrder.PickedItems.U_BnDistNumber = [string]$batch.Batch;
                        
                        $pickOrder.PickedItems.U_Quantity = $batch.Quantity;
                        $pickOrder.PickedItems.U_ReqItmLn = $pickOrderLineNum;
                        $batch.BatchLineNum = $pickOrder.PickedItems.U_LineNum;
                        $dummy = $pickOrder.PickedItems.Add();

                        #Specify relation Beetween picked and Required Line 
                        $pickOrder.Relations.SetCurrentLine($pickOrder.Relations.Count - 1);
                        $pickOrder.Relations.U_ReqItemLineNo = [int]$line.U_LineNum
                        $pickOrder.Relations.U_PickItemLineNo = [int]$batch.BatchLineNum;
                        $dummy = $pickOrder.Relations.Add();
                    }

                    foreach ($bin in $bins) {
                        $batch = $bin.Batch;
                        $BatchLineNum = -1;
                        if ($batches.ContainsKey($batch)) {
                            $BatchLineNum = $batches[$batch].BatchLineNum;
                        }

                        $pickOrder.BinAllocations.SetCurrentLine($pickOrder.BinAllocations.Count - 1);
                        $pickOrder.BinAllocations.U_BinAbsEntry = $bin.BinAbsEntry;
                        $pickOrder.BinAllocations.U_Quantity = $bin.Quantity;
                        if ($BatchLineNum -gt -1) {
                            $pickOrder.BinAllocations.U_SnAndBnLine = $BatchLineNum;
                        }
                        $dummy = $pickOrder.BinAllocations.Add()
                    }
                }

                #delete
                for ($idxD = $linesToBeRemoved.Count - 1; $idxD -ge 0; $idxD--) {
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
                    $dummy = $pickOrder.RequiredItems.DelRowAtPos($linesToBeRemoved[$idxD]);
                }

                $result = $pickOrder.Update();
                
                if ($result -lt 0) {    
                    $err = $pfcCompany.GetLastErrorDescription()
                    Throw [System.Exception] ($err)
                }

                if ([CompuTec.ProcessForce.API.Configuration.ConfigurationHolder]::Configuration[$pfcCompany.Token].AutoIssueEnabled -eq $false) {
                    if ($csvItem.CreateGoodsIssue -eq "Y") {
                        $pickOrderGIAction = $pfcCompany.CreatePFAction([CompuTec.ProcessForce.API.Core.ActionType]::CreateGoodsIssueFromPickOrderBasedOnProductionIssue);
                        $pickOrderGIAction.PickOrderID = $docentryPickOrder
                        $giDocEntry = 0;
                        $result = $pickOrderGIAction.DoAction([ref] $giDocEntry)

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
            
            if ($newPickOrder -eq $true) {
                $taskMsg = "adding";
            }
            else {
                $taskMsg = "updating"
            }
            $ms = [string]::Format("Error when {0} Pick Order for MOR {1} Details: {2}", $taskMsg, $csvItem.MORDocEntry, $err);
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