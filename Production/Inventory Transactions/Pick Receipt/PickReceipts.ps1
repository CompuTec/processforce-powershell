#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Pick Receipts
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.170) PL: 07 R1 Pre-Release (64-bit)
# Description:
#      Import Time Bookings. Script add new Time Bookings.
#      You need to have all requred files for import.
# Warning:
#   It's recommended run script when all users all disconnected.
#   Before running this script please do database backup.
# Troubleshooting:
#   https://connect.computec.pl/display/PF930EN/FAQ+PowerShell
#   https://connect.computec.pl/display/PF920EN/FAQ+PowerShell
# Script source:
#   https://connect.computec.pl/display/PF930EN/Quality+Control+scripts
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
$csvPickReceiptsBatchesPath = -join ($csvImportCatalog, "PickReceiptsBatches.csv")
$csvPickReceiptsBinsPath = -join ($csvImportCatalog, "PickReceiptsBins.csv")

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

    [array]$pickReceipts = Import-Csv -Delimiter ';' -Path $csvPickReceiptsPath
    [array]$pickReceiptsLines = Import-Csv -Delimiter ';' -Path $csvPickReceiptsLinesPath

    [array]$pickReceiptsBatches = $null;
    if ((Test-Path -Path $csvPickReceiptsBatchesPath -PathType leaf) -eq $true) {
        [array]$pickReceiptsBatches = Import-Csv -Delimiter ';' -Path $csvPickReceiptsBatchesPath
    }
    else {
        write-host "Pick Receipts Batches - csv not available."
    }

    [array]$pickReceiptsBins = $null;
    if ((Test-Path -Path $csvPickReceiptsBinsPath -PathType leaf) -eq $true) {
        [array]$pickReceiptsBins = Import-Csv -Delimiter ';' -Path $csvPickReceiptsBinsPath 
    }
    else {
        write-host "Pick Receips Bins - csv not available."
    }
    write-Host 'Preparing data: '
    $totalRows = $pickReceipts.Count + $pickReceiptsLines.Count + $pickReceiptsBatches.Count + $pickReceiptsBins.Count

    $prList = New-Object 'System.Collections.Generic.List[array]'

    $dictionaryPRLines = New-Object 'System.Collections.Generic.Dictionary[string,psobject]'
    $dictionaryPRBatches = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryPRBins = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

    $progressItterator = 0;
    $progres = 0;
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
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        $prList.Add([array]$row);
    }

    foreach ($row in $pickReceiptsLines) {
        $key = $row.MORDocEntry;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryPRLines.ContainsKey($key)) {
            $dictionaryPRLines[$key] = [psobject]@{
                existingLines = New-Object 'System.Collections.Generic.Dictionary[int,array]';
                newLines      = New-Object 'System.Collections.Generic.List[array]';
            }
        }
        
        if($row.PickLineNum -gt 0){
            $dictionaryPRLines[$key].existingLines.Add($row.PickLineNum,[array]$row);
        } else {
            $dictionaryPRLines[$key].newLines.Add([array]$row);
        }
    }

    foreach ($row in $pickReceiptsBatches) {
        $key = $row.MORDocEntry;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryPRBatches.ContainsKey($key)) {
            $list = $dictionaryPRBatches[$key];
        }
        else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryPRBatches[$key] = $list;
        }
        $list.Add([array]$row);
    }

    foreach ($row in $pickReceiptsBins) {
        $key = $row.MORDocEntry;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryPRBins.ContainsKey($key)) {
            $list = $dictionaryPRBins[$key];
        }
        else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryPRBins[$key] = $list;
        }
        $list.Add([array]$row);
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
                $pickReceiptAction.AddManufacturingOrderDocEntry($csvItem.MORDocEntry);
                $pickReceiptAction.AddManufacturingOrderDocEntry(39);
                $pickReceiptAction.ReceiptType = [CompuTec.ProcessForce.API.Actions.CreatePickReceiptForProductionReceipt.PickOrderdReceiptType]::FinalGood
                $docentryPickReceipt = 0;
                $pickReceiptAction.DoAction([ref] $docentryPickReceipt);
            }

            try {
                $result = $pickReceipt.GetByKeyNew($docentryPickReceipt);
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

           
            $requiredItemsCount = $pickReceipt.RequiredItems.Count;
            if($requiredItemsCount -gt 0){
                
                $prItems = $dictionaryPRLines[$dictionaryKey];
                $linesToBeRemoved = New-Object 'System.Collections.Generic.List[int]';
                $currentPosDict = New-Object 'System.Collections.Generic.Dictionary[string,int]';

                $itemIndex = 0;
                foreach ($line in $pickReceipt.RequiredItems) {
                    $LineNum = $line.U_LineNum;
                    if ($newMORItems.existingLines.ContainsKey($LineNum) -eq $false) {
                        $linesToBeRemoved.Add($itemIndex);
                    } else {
                        $currentPosDict.Add($LineNum,$itemIndex);
                    }
                    $itemIndex++;
                }
                
                



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
            $ms = [string]::Format("Error when {0} Pirck Receipt with ItemCode {1} and Revision {2} Details: {3}", $taskMsg, $csvItem.BOM_ItemCode, $csvItem.Revision, $err);
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