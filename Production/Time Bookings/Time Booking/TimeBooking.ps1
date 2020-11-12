#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Time Bookings
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Time Bookings. Script add new Time Bookings.
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

$csvTimeBookingsPath = -join ($csvImportCatalog, "TimeBookings.csv")
$csvTimeBookingLinesPath = -join ($csvImportCatalog, "TimeBookingLines.csv")

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
    [array] $csvTimeBookings = Import-Csv -Delimiter ';' $csvTimeBookingsPath;
    [array] $csvTimeBookingLines = Import-Csv -Delimiter ';' $csvTimeBookingLinesPath;
	
    write-Host 'Preparing data: '
    $totalRows = $csvTimeBookings.Count + $csvTimeBookingLines.Count;
    $timeBookingsList = New-Object 'System.Collections.Generic.List[array]'
    $timeBookingsLinesDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
	
    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $csvTimeBookings) {
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        $timeBookingsList.Add([array]$row);
    }

    foreach ($row in $csvTimeBookingLines) {
        $key = $row.Key;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
	
        if ($timeBookingsLinesDict.ContainsKey($key) -eq $false) {
            $timeBookingsLinesDict[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $timeBookingsLinesDict[$key];
		
        $list.Add([array]$row);
    }

    Write-Host '';
    Write-Host 'Adding/updating data: ' -NoNewline;
    if ($downtimeList.Count -gt 1) {
        $total = $downtimeList.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;

    foreach ($csvItem in $downtimeList) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            #Creating BOM object
            $otr = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"OperationTimeRecording")
            $otr.U_DocDate = $csvItem.DocDate;
            $otr.U_Remarks = $csvItem.Remarks;
            $otr.U_Ref2 = $csv.Ref2;
	
            #Data loading from a csv file - lines
            [array]$otrLinesCsv = $timeBookingsLinesDict[$csvItem.Key]
            foreach ($otrLineCsv in $otrLinesCsv) {
                $otr.Lines.U_BaseEntry = $otrLineCsv.BaseEntry;
                $otr.Lines.U_BaseDocNum = $ortLineCsv.BaseDocNum;
                $otr.Lines.U_RscCode = $otrLineCsv.ResourceCode;
                $otr.Lines.U_BaseLineNum = $otrLineCsv.BaseLineNum;
                $otr.Lines.U_OprCode = $otrLineCsv.OperationCode;
                $otr.Lines.U_Remarks = $otrLineCsv.Remarks;
                switch ($otrLineCsv.TimeType) {
                    "Q" {
                        $otr.Lines.U_TimeType = [CompuTec.ProcessForce.API.Enumerators.RecordingTimeType]::QueueTime;
                        break
                    }
                    "S" {
                        $otr.Lines.U_TimeType = [CompuTec.ProcessForce.API.Enumerators.RecordingTimeType]::SetupTime;
                        break
                    }
                    "R" {
                        $otr.Lines.U_TimeType = [CompuTec.ProcessForce.API.Enumerators.RecordingTimeType]::RunTime;
                        break
                    }
                    "L" {
                        $otr.Lines.U_TimeType = [CompuTec.ProcessForce.API.Enumerators.RecordingTimeType]::StockTime;
                        break
                    }
                    default {
						$err = [string]::Format("Incorrect TimeType: .Possible values are: Q, S, R, L",$otrLineCsv.TimeType);
                        Throw [System.Exception]($err);
                    }
                }
		
                if ($otrLineCsv.NumberOfResources -gt 1) {
                    $otr.Lines.U_NrOfResources = $otrLineCsv.NumberOfResources;
                }
                else {
                    $otr.Lines.U_NrOfResources = 1;
                }
		
                if ($otrLineCsv.StartDate -gt '') {
                    $otr.Lines.U_StartDate = $otrLineCsv.StartDate;
                }
                if ($otrLineCsv.StartTime -ne '') {
                    $otr.Lines.U_StartTime = $otrLineCsv.StartTime;
                }		
                if ($otrLineCsv.EndDate -ne '') {
                    $otr.Lines.U_EndDate = $otrLineCsv.EndDate;
                }
                if ($otrLineCsv.EndTime -ne '') {
                    $otr.Lines.U_EndTime = $otrLineCsv.EndTime;
                }
                $otr.Lines.U_WorkingHours = $otrLineCsv.WorkingHours;
                $dummy = $otr.Lines.Add();
            }	
	
            $message = 0
            $message = $otr.Add()
            if ($message -lt 0) {    
                $err = $pfcCompany.GetLastErrorDescription()
                Throw [System.Exception] ($err)
            }
        }
        Catch {
            $err = $_.Exception.Message;
	
            $taskMsg = "adding";
	
            $ms = [string]::Format("Error when {0} Time Booking with Key: {1} Details: {2}", $taskMsg, $csvItem.Key, $err);
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

