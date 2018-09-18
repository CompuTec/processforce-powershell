#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Item Details Revisions
########################################################################
$SCRIPT_VERSION = "3.1"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Item Details Revisions. Script add new or will update existing revisions.
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
# $csvImportCatalog = "C:\PS\PF\TestProtocols\";

$csvItemsRevisionsPath = -join ($csvImportCatalog, "ItemDetails_Revisions.csv")

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
#endregion


try {
    $atLeastOneSuccess = $false;
    #Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
    Write-Host 'Preparing data: '
    [array]$csvItemRevisions = Import-Csv -Delimiter ';' -Path $csvItemsRevisionsPath
    $dictionaryItemsRevisions = New-Object 'System.Collections.Generic.Dictionary[string,psobject]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
    if ($csvItemRevisions.Count -gt 1) {
        $total = $csvItemRevisions.Count
    }
    else {
        $total = 1
    }

    foreach ($row in $csvItemRevisions) {
        $key = $row.Itemcode;
        $revCode = $row.RevisionCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }        

        if ($dictionaryItemsRevisions.ContainsKey($key) -eq $false) {
            $dictionaryItemsRevisions.Add($key, (New-Object 'System.Collections.Generic.Dictionary[string,array]'));
        }

        if ($dictionaryItemsRevisions[$key].ContainsKey($revCode) -eq $false) {
            $dictionaryItemsRevisions[$key].Add($revCode, [array]$row);
        }
    }
    Write-Host '';
    Write-Host 'Add/Update data to SAP: '
    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
    if ($dictionaryItemsRevisions.Count -gt 1) {
        $total = $dictionaryItemsRevisions.Count
    }
    else {
        $total = 1
    }

    #checking state of DirectAccess mode
    $DirectAccessState = [CompuTec.Core.DI.Database.DataLayer]::GetLayerState($pfcCompany.Token);
    if ($DirectAccessState -eq [CompuTec.Core.DI.Database.DataLayerState]::"Connected") {
        $DIRECT_ACCESS_MODE = $true;
        [CompuTec.ProcessForce.API.DynamicCosting.Restoration.ItemCostingRestorationManager]::SuspendItemCostingRestoration = $true
    }
    else {
        $DIRECT_ACCESS_MODE = $false;
        write-host -backgroundcolor yellow -foregroundcolor blue "Please enable Direct Access Mode to speed up import process";
    }
    
    #Checking that Item Details already exist
    foreach ($key in $dictionaryItemsRevisions.Keys) {
        try {
            $progressItterator++;
            $progres = [math]::Round(($progressItterator * 100) / $total);
            if ($progres -gt $beforeProgress) {
                Write-Host $progres"% " -NoNewline
                $beforeProgress = $progres
            }
            $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]::"BoRecordset")
      
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
            $dummy = $rs.DoQuery([string]::Format( "SELECT T0.""ItemCode"" AS ""ItemCode"" FROM OITM T0
                INNER JOIN ""@CT_PF_OIDT"" T1 ON T0.""ItemCode"" = T1.""U_ItemCode"" WHERE T0.""ItemCode"" = N'{0}'", $key))
        
            if ($rs.RecordCount -eq 0) {
                $err = [string]::Format('Item Master Data with ItemCode {0} do not exists. Please restore Item Details', $key);
                Throw [System.Exception] ($err)
            }
   
            #Creating Item Details
            $idt = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::"ItemDetails")
      
            #Restoring Item Costs and setting Inherit Batch/Serial to 'Yes'
            $dummy = $idt.GetByItemCode($key)
          
            $revisions = $dictionaryItemsRevisions[$key];
            if ($revisions.count -gt 0) {
    
                $linesToBeRemoved = New-Object 'System.Collections.Generic.List[int]';
                $currentPosDict = New-Object 'System.Collections.Generic.Dictionary[string,int]';

                $revIndex = 0;
                foreach ($revision in $idt.Revisions) {
                    if ($revision.U_Code -gt "") {
                        if ($revisions.ContainsKey($revision.U_Code)) {
                            $currentPosDict.Add($revision.U_Code, $revIndex);
                        }
                        else {
                            $linesToBeRemoved.Add($revIndex);
                        }
                    }
                    $revIndex++;
                }
          
                #updating existing revision
                foreach ($revCode in $revisions.Keys) {
                    $rev = $revisions[$revCode];
                    if ($currentPosDict.ContainsKey($revCode)) {
                        $idt.Revisions.SetCurrentLine($currentPosDict[$revCode]);
                    }
                    else {
                        $idt.Revisions.SetCurrentLine($idt.Revisions.Count - 1);
                        $idt.Revisions.U_Code = $rev.RevisionCode
                    }

                    $idt.Revisions.U_Description = $rev.RevisionName
                    $idt.Revisions.U_Status = $rev.Status #enum type; Revision Status, Active ACT = 1, BeingPhasedOut BPO = 2, Engineering ENG = 3, Obsolete OBS = 4
                    if ($rev.ValidFrom -gt '') {
                        $idt.Revisions.U_ValidFrom = $rev.ValidFrom
                    }
                    if ($rev.ValidTo -gt '') {
                        $idt.Revisions.U_ValidTo = $rev.ValidTo
                    }
                    $idt.Revisions.U_Remarks = $rev.Remarks
                    $idt.Revisions.U_Default = $rev.IsDefault #enum type; 1 = Yes, 2 = No
                    $idt.Revisions.U_IsMRPDefault = $rev.IsMRPDefault #enum type; 1 = Yes, 2 = No
                    $idt.Revisions.U_IsCostingDefault = $rev.DefaultForCosting #enum type; 1 = Yes, 2 = No
                    if ($currentPosDict.ContainsKey($revCode) -eq $false) {
                        $dummy = $idt.Revisions.Add();
                    }
                }

                #Deleting revision
                for ($idxD = $linesToBeRemoved.Count - 1; $idxD -ge 0; $idxD--) {
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
                    $dummy = $idt.Revisions.DelRowAtPos($linesToBeRemoved[$idxD]);
                }
          
            }
   
            $message = $idt.Update()
        
            if ($message -lt 0) {  
                $err = $pfcCompany.GetLastErrorDescription()
                Throw [System.Exception] ($err);
            }
            else {
                $atLeastOneSuccess = $true;
            }
        }
        Catch {
            $err = $_.Exception.Message;
            $ms = [string]::Format("Error when adding/updating Item Details for ItemCode {0} Details: {1}", $key, $err);
            Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
            if ($pfcCompany.InTransaction) {
                $pfcCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
            } 
        }
    }
}
Catch {
    $err = $_.Exception.Message;
    $ms = [string]::Format("Exception occured:", $err);
    Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
    if ($pfcCompany.InTransaction) {
        $pfcCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
    } 
}
Finally {
    if ($DIRECT_ACCESS_MODE -eq $true) {
        [CompuTec.ProcessForce.API.DynamicCosting.Restoration.ItemCostingRestorationManager]::SuspendItemCostingRestoration = $false
        if ($atLeastOneSuccess -eq $true) { 
            Write-Host '';
            Write-Host 'Restoring cosing data'
            $manager = New-Object  "CompuTec.ProcessForce.API.DynamicCosting.Restoration.ItemCostingRestorationManager" $pfcCompany.Token
            $manager.Initialize();
            $manager.Restore();
        }
    }

    #region Close connection
    
    if ($pfcCompany.IsConnected) {
        $pfcCompany.Disconnect()
        write-host  –backgroundcolor green –foregroundcolor black "Disconnected from the company"
    }
    
    #endregion

}


