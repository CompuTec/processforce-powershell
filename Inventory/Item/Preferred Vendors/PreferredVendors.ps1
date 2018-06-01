Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany();
$pfcCompany.LicenseServer = "10.0.0.2:40000";
$pfcCompany.SQLServer = "10.0.0.1:30115";
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB";
$pfcCompany.Databasename = 'PFDEMOGB';
$pfcCompany.UserName = "manager";
$pfcCompany.Password = "1234";
        
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


#Data loading from a csv file - Items for which Preferred Vendors will be added (each of them has to have Item Master Data)
[array] $csvItemsPreferredVendors = Import-Csv -Delimiter ';' -Path  ($PSScriptRoot + "\Items_PreferredVendors.csv")
 
#region preparing data
write-Host 'Preparing data: '
$totalRows = $csvItemsPreferredVendors.Count;

$dictionaryPV = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    
$progressItterator = 0;
$progres = 0;
$beforeProgress = 0;

if ($totalRows -gt 1) {
    $total = $totalRows
}
else {
    $total = 1
}

foreach ($row in $csvItemsPreferredVendors) {
    $key = $row.ItemCode;
    $progressItterator++;
    $progres = [math]::Round(($progressItterator * 100) / $total);
    if ($progres -gt $beforeProgress) {
        Write-Host $progres"% " -NoNewline
        $beforeProgress = $progres
    }

    if ($dictionaryPV.ContainsKey($key)) {
        $list = $dictionaryPV[$key];
    }
    else {
        $list = New-Object System.Collections.Generic.List[array];
        $dictionaryPV[$key] = $list;
    }

    $list.Add([array]$row);
}

Write-Host '';
#endregion

$progressItterator = 0;
$progres = 0;
$beforeProgress = 0;
$totalRows = $dictionaryPV.Keys.Count;
if ($totalRows -gt 1) {
    $total = $totalRows
}
else {
    $total = 1
}
write-Host 'Adding data to SAP: '
#Checking that Item Details already exist 
foreach ($csvItemCode in $dictionaryPV.Keys) {
    $progressItterator++;
    $progres = [math]::Round(($progressItterator * 100) / $total);
    if ($progres -gt $beforeProgress) {
        Write-Host $progres"% " -NoNewline
        $beforeProgress = $progres
    }
    try {
        $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]::"BoRecordset")
     
        $rs.DoQuery([string]::Format( "SELECT T0.""Code"" FROM ""@CT_PF_OPVR"" T0 WHERE T0.""U_ItemCode"" = N'{0}'", $csvItemCode))
        $exists = 0;
        if ($rs.RecordCount -gt 0) {
            $exists = 1
            $Code = $rs.Fields.Item("Code").Value
        }
  
        #Creating Item Preferred Vendors 
        $PVObject = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::PreferredVendors);
     
           
        if ($exists -eq 1) {
            $PVObject.GetByKey($Code);
        }
        else {
            $PVObject.U_ItemCode = $csvItemCode;
        }
     
        [array]$PVForItemList = $dictionaryPV[$csvItemCode];
        if ($PVForItemList.count -gt 0) {
            #Deleting all existing preferred vendors
            $count = $PVObject.Rows.Count
            for ($i = 0; $i -lt $count; $i++) {
                $dummy = $PVObject.Rows.DelRowAtPos(0);
            }
            $PVObject.Rows.SetCurrentLine( $PVObject.Rows.Count - 1);
         
            #Adding preferred vendors
            foreach ($PVForItem in $PVForItemList) {
                $PVObject.Rows.U_ItemCode = $csvItemCode;
                $PVObject.Rows.U_VendorCode = $PVForItem.VendorCode

                if ($PVForItem.IsDefault -eq 'Y') {
                    $PVObject.Rows.U_IsDefault = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
                }
                else {
                    $PVObject.Rows.U_IsDefault = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
                }

                $PVObject.Rows.U_Percentage = $PVForItem.Percentage;
                $PVObject.Rows.U_StartDate = $PVForItem.StartDate;
                $PVObject.Rows.U_EndDate = $PVForItem.EndDate;

                $dummy = $PVObject.Rows.Add()
            }
        }
  
        $message = 0
     
        #Adding or updating depends if object already exists in the database
        
        if ($exists -eq 1) {
            [System.String]::Format("Updating Preferred Vendors for item: {0}", $csvItemCode)
            $message = $PVObject.Update()
        }
        else {
            [System.String]::Format("Adding Preferred Vendors for item: {0}", $csvItemCode)
            $message = $PVObject.Add()
        }
     
        if ($message -lt 0) {    
            $err = $pfcCompany.GetLastErrorDescription()
            Throw [System.Exception]($err);
        } 
        else {
            write-host "Success"
        }   
    } 
    Catch {
        $err = $_.Exception.Message;
        $content = [string]::Format("Error occured for Item {0}: {1}", $csvItemCode, $err);
        Write-Host -BackgroundColor DarkRed -ForegroundColor White $content;
        continue;
    }
    
}