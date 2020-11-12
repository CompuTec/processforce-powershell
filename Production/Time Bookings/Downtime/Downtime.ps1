#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Downtimes
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Downtimes. Script add new or will update existing Downtimes.
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

$csvDowntimePath = -join ($csvImportCatalog, "Downtime.csv")
$csvDowntimeReasonsPath = -join ($csvImportCatalog, "DowntimeReasonsLines.csv")
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
    [array] $csvDowntime = Import-Csv -Delimiter ';' -Path $csvDowntimePath
    [array] $csvDowntimeReasons = Import-Csv -Delimiter ';' -Path $csvDowntimeReasonsPath
 
    #region preparing data
    write-Host 'Preparing data: ' -NoNewline
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
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progres -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
    
            $exists = $false;
            if ($item.DocEntry -gt 0) {
                $exists = $true;
            }
  
            #Creating PF Object
            $PFObject = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::DownTime);
            if ($exists -eq $true) {
                $retValue = $PFObject.GetByKey($item.DocEntry);
                if ($retValue -ne 0) {
                    $err = [string]::Format("Downtime document with DocEntry: {0} don't exists", $ite.DocEntry);
                    Throw [System.Exception]($err);
                }
            }

            $PFObject.U_RscCode = $item.RscCode;

            $morDocEntry = $item.ManufacturingOrderDocEntry;
            if ($morDocEntry -gt 0) {
                $PFObject.U_MorDE = $morDocEntry
                $operCode = $item.OperationCode;

                if ($null -ne $operCode -and $operCode -ne '') {
                    $PFObject.U_OprCode = $operCode;
                }
            }

            $PFObject.U_DocDate = $item.DocDate;
            $PFObject.U_StartDate = $item.StartDate;
            $startTime = $item.StartTime;
            if ($null -ne $startTime -and $startTime -gt 0) {
                $PFObject.U_StartTime = $startTime;
            }
            $PFObject.U_EndDate = $item.EndDate;
            $endTime = $item.EndTime;
            if ($null -ne $endTime -and $endTime -gt 0) {
                $PFObject.U_EndTime = $endTime;
            }

            $PFObject.U_ReporterID = $item.ReporterId;
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
            if ($exists -eq $true) {
                $message = $PFObject.Update();
            }
            else {
                $message = $PFObject.Add();
            }
     
            if ($message -lt 0) {    
                $err = $pfcCompany.GetLastErrorDescription()
                Throw [System.Exception] ($err)
            }
        }
        Catch {
            $err = $_.Exception.Message;
            if ($exists -eq $false) {
                $taskMsg = "adding";
            }
            else {
                $taskMsg = "updating"
            }
            $ms = [string]::Format("Error when {0} Downtime with Key {1} Details: {2}", $taskMsg, $item.Key, $err);
            Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
            if ($pfcCompany.InTransaction) {
                $pfcCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
            } 
        }		
    
    }
}
Catch {
    $err = $_.Exception.Message;
    $ms = [string]::Format("Exception occured:{0}", $err);
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



