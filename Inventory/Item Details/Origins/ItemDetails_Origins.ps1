Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.LicenseServer = "10.0.0.2:40000";
$pfcCompany.SQLServer = "10.0.0.1:30115";
$pfcCompany.Databasename = 'PFDEMO';
$pfcCompany.UserName = "manager";
$pfcCompany.Password = "1234";
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"; 
        
$code = $pfcCompany.Connect()
if ($code -eq 1) {
    #Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
    [array] $csvItems = Import-Csv -Delimiter ';' -Path  ($PSScriptRoot + "\ItemDetails.csv")
    [array] $csvItemsOrigins = Import-Csv -Delimiter ';' -Path  ($PSScriptRoot + "\ItemDetails_Origins.csv")
 
    #region preparing data
    write-Host 'Preparing data: '
    $totalRows = $csvItems.Count + $csvItemsOrigins.Count;

    $itemDetailsList = New-Object 'System.Collections.Generic.List[array]'
    $dictionaryOrigins = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    
    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;

    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $csvItems) {
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }


        $itemDetailsList.Add([array]$row);
    }

    foreach ($row in $csvItemsOrigins) {
        $key = $row.ItemCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryOrigins.ContainsKey($key)) {
            $list = $dictionaryOrigins[$key];
        }
        else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryOrigins[$key] = $list;
        }

        $list.Add([array]$row);
    }

    Write-Host '';
    #endregion

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
    $totalRows = $itemDetailsList.Count
    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }
    write-Host 'Adding data to SAP: '
    #Checking that Item Details already exist 
    foreach ($csvItem in $itemDetailsList) {
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        try {
            $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]::"BoRecordset")
     
            $rs.DoQuery([string]::Format( "SELECT T0.""U_ItemCode"" FROM ""@CT_PF_OIDT"" T0 WHERE T0.""U_ItemCode"" = N'{0}'", $csvItem.ItemCode))
            $exists = 0;
            if ($rs.RecordCount -gt 0) {
                $exists = 1
            }
  
            #Creating Item Details 
            $idt = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::"ItemDetails")
     
            #Restoring Item Costs and setting Inherit Batch/Serial to 'Yes'
            if ($exists -eq 1) {
                $idt.GetByItemCode($csvItem.ItemCode)
            }
            else {
                $idt.U_ItemCode = $csvItem.ItemCode;
            }

            $idt.U_DftOrigin = $csvItem.DefaultOrigin;
     
            [array]$origins = $dictionaryOrigins[$idt.U_ItemCode];
            if ($origins.count -gt 0) {
                #Deleting all exisitng Phrases
                $count = $idt.Origins.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $idt.Origins.DelRowAtPos(0);
                }
                $idt.Origins.SetCurrentLine($idt.Origins.Count - 1);
         
                #Adding Origins
                foreach ($origin in $origins) {
                    $idt.Origins.U_CountryCode = $origin.CountryCode;
                    $dummy = $idt.Origins.Add()
                }
            }
  
            $message = 0
     
            #Adding or updating ItemDetails depends if it exists in the database
        
            if ($exists -eq 1) {
                [System.String]::Format("Updating Item Details: {0}", $csvItem.ItemCode)
                $message = $idt.Update()
            }
            else {
                [System.String]::Format("Adding Item Details: {0}", $csvItem.ItemCode)
                $message = $idt.Add()
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
            $content = [string]::Format("Error occured for ItemDeitals {0}: {1}", $idt.U_ItemCode, $err);
            Write-Host -BackgroundColor DarkRed -ForegroundColor White $content;
            continue;
        }
    
    }
}
else {
    write-host "Connection Failure"
}