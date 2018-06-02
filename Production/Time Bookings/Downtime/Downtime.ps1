Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany();
$pfcCompany.LicenseServer = "10.0.0.3:40000";
$pfcCompany.SQLServer = "10.0.0.2:30115";
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

#Data loading from a csv file
[array] $csvDowntime = Import-Csv -Delimiter ';' -Path  ($PSScriptRoot + "\Downtime.csv")
[array] $csvDowntimeReasons = Import-Csv -Delimiter ';' -Path  ($PSScriptRoot + "\DowntimeReasons.csv")
 
#region preparing data
write-Host 'Preparing data: '
$totalRows = $csvDowntime.Count + $csvDowntimeReasons.Count;

$downtimeList = New-Object 'System.Collections.Generic.List[array]'
$downtimeReasonsDict = New-Object 'System.Collections.Generic.Dictionary[String,System.Collections.Generic.List[array]]'
    
$progressItterator = 0;
$progress = 0;
$beforeProgress = 0;

if ($totalRows -gt 1) {
    $total = $totalRows
}
else {
    $total = 1
}

foreach ($row in $csvDowntime) {
    $progressItterator++;
    $progres = [math]::Round(($progressItterator * 100) / $total);
    if ($progres -gt $beforeProgress) {
        Write-Host $progres"% " -NoNewline
        $beforeProgress = $progres
    }

    $downtimeList.Add([array]$row);
}

foreach ($row in $csvDowntimeReasons) {
    $key = $row.Key;
    $progressItterator++;
    $progres = [math]::Round(($progressItterator * 100) / $total);
    if ($progres -gt $beforeProgress) {
        Write-Host $progres"% " -NoNewline
        $beforeProgress = $progres
    }

    if ($downtimeReasonsDict.ContainsKey($key)) {
        $list = $downtimeReasonsDict[$key];
    }
    else {
        $list = New-Object System.Collections.Generic.List[array];
        $downtimeReasonsDict[$key] = $list;
    }

    $list.Add([array]$row);
}


Write-Host '';
#endregion


$progressItterator = 0;
$progress = 0;
$beforeProgress = 0;
$totalRows = $downtimeList.Count;
if ($totalRows -gt 1) {
    $total = $totalRows
}
else {
    $total = 1
}
write-Host 'Adding data to SAP: '
foreach ($item in $downtimeList) {
    $progressItterator++;
    $progress = [math]::Round(($progressItterator * 100) / $total);
    if ($progres -gt $beforeProgress) {
        Write-Host $progress"% " -NoNewline
        $beforeProgress = $progress
    }
    try {
        $exists = 0;
        if ($item.DocEntry -gt 0) {
            $exists = 1
        }
  
        #Creating PF Object
        $PFObject = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::DownTime);
  
        if ($exists -eq 1) {
            $PFObject.GetByKey($item.DocEntry);
        }

        $PFObject.U_RscCode = $item.RscCode;

        $morDocEntry = $item.ManufacturingOrderDocEntry;
        if($morDocEntry -gt 0){
            $PFObject.U_MorDE = $morDocEntry
            $operCode = $item.OperationCode;

            if($operCode -ne $null -and $operCode -ne ''){
                $PFObject.U_OprCode = $operCode;
            }
        }

        $PFObject.U_DocDate = $item.DocDate;
        $PFObject.U_StartDate = $item.StartDate;
        $startTime = $item.StartTime;
        if($startTime -ne $null -and $startTime -gt 0){
            $PFObject.U_StartTime = $startTime;
        }
        $PFObject.U_EndDate = $item.EndDate;
        $endTime = $item.EndTime;
        if($endTime -ne $null -and $endTime -gt 0){
            $PFObject.U_EndTime = $endTime;
        }

        $PFObject.U_ReporterID =  $item.ReporterId;
        $PFObject.U_TechnicID = $item.TechnicanId;
        $PFObject.U_ClosedBy = $item.ClosedById;

        $status = $item.Status;

        switch ($status) {
            'R' { $PFObject.U_Status = [CompuTec.ProcessForce.API.Enumerators.DownTimeStatus]::Reported; break; }
            'I' { $PFObject.U_Status = [CompuTec.ProcessForce.API.Enumerators.DownTimeStatus]::InProgress; break; }
            'W' { $PFObject.U_Status = [CompuTec.ProcessForce.API.Enumerators.DownTimeStatus]::Working; break; }
            'F' { $PFObject.U_Status = [CompuTec.ProcessForce.API.Enumerators.DownTimeStatus]::Fixed; break; }
            Default {$PFObject.U_Status = [CompuTec.ProcessForce.API.Enumerators.DownTimeStatus]::Reported}
        }

        [array] $reasonsList = $downtimeReasonsDict[$item.Key];
        if ($reasonsList.count -gt 0) {
            #Deleting all exisitng reasons
            $count = $PFObject.Reasons.Count
            for ($i = 0; $i -lt $count; $i++) {
                $dummy = $PFObject.Reasons.DelRowAtPos(0);
            }
            $PFObject.Reasons.SetCurrentLine($PFObject.Reasons.Count - 1);
     
            #Adding Origins
            foreach ($reason in $reasonsList) {
                $PFObject.Reasons.U_ReasonCode = $reason.Code;
                $PFObject.Reasons.U_Remarks = $reason.Remarks;
                $dummy = $PFObject.Reasons.Add()
            }
        }

        $message = 0

        #Adding or updating depends if object already exists in the database
        if ($exists -eq 1) {
            [System.String]::Format("Updating Downtime with key: {0}", $item.Key);
            $message = $PFObject.Update();
        }
        else {
            [System.String]::Format("Adding Downtime with key: {0}", $item.Key);
            $message = $PFObject.Add();
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
        $content = [string]::Format("Error occured for Code {0}: {1}", $item.Code, $err);
        Write-Host -BackgroundColor DarkRed -ForegroundColor White $content;
        continue;
    }
    
}

