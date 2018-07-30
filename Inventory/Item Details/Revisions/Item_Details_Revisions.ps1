#region #PF API library usage
Clear-Host
# You need to check in what architecture PowerShell ISE is running (x64 or x86),
# you need run PowerShell ISE in the same architecture like PF API is installed (check in Windows -> Programs & Features)
# Examples:
#     SAP Client x64 + PF x64 installed on DB/Company => PF API x64 => Windows PowerShell ISE
#     SAP Client x86 + PF x86 installed on DB/Company => PF API x86 => Windows PowerShell ISE x86

[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#endregion

#region Script parameters

$csvImportCatalog = $PSScriptRoot + "\"

$csvItemsRevisionsPath = -join ($csvImportCatalog, "ItemDetails_Revisions.csv")

#endregion

#region #Datbase/Company connection settings
 
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.LicenseServer = "10.0.0.203:40000"
$pfcCompany.SQLServer = "10.0.0.202:30115"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"
$pfcCompany.Databasename = "PFDEMOGB_MACIEJP"
$pfcCompany.UserName = "maciejp"
$pfcCompany.Password = "1234"
 
# where:
 
# LicenseServer = SAP LicenceServer name or IP Address with port number (see in SAP Client -> Administration -> Licence -> Licence Administration -> Licence Server)
# SQLServer     = Server name or IP Address with port number, should be the same like in System Landscape Dirctory (see https://<Server>:<Port>/ControlCenter) - sometimes best is use IP Address for resolve connection problems.
#
# DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2014"     # For MsSQL Server 2014
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
                if($revision.U_Code -gt ""){
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
