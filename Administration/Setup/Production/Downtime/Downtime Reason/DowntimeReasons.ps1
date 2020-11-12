#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Properties References
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Properties References. Script add new or will update existing References.
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

$csvDowntimeReasonsPath = -join ($csvImportCatalog, "DowntimeReasons.csv")

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
 
    #endregion
    $UDO_DOWNTIME_REASON = 'CT_PF_DTReasonCodes';
    #Data loading from a csv file
    [array] $csvDowntimeReasons = Import-Csv -Delimiter ';' -Path  $csvDowntimeReasonsPath
 
    #region preparing data
    write-Host 'Preparing data: '
    $totalRows = $csvDowntimeReasons.Count;

    $sourceList = New-Object 'System.Collections.Generic.List[array]'
    
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;

    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $csvDowntimeReasons) {
        $progressItterator++;
        $progress = [math]::Round(($progressItterator * 100) / $total);
        if ($progress -gt $beforeProgress) {
            Write-Host $progress"% " -NoNewline
            $beforeProgress = $progress
        }
        $sourceList.Add([array]$row);
    }

    Write-Host '';
    #endregion


    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;
    $totalRows = $sourceList.Count;
    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }
    write-Host 'Adding data to SAP: '
    foreach ($item in $sourceList) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]::"BoRecordset")
     
            $rs.DoQuery([string]::Format( "SELECT T0.""Code"" FROM ""@CT_PF_ODTRS"" T0 WHERE T0.""Code"" = N'{0}'", $item.Code))
            $exists = $false;
            if ($rs.RecordCount -gt 0) {
                $exists = $true
            }
  
            #Creating PF Object
            #$PFObject = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::DownTimeReason);
            $Company = $pfcCompany.SapCompany;
        
            $cs = $Company.GetCompanyService();
            $gs = $cs.GetGeneralService($UDO_DOWNTIME_REASON);
           
            if ($exists -eq $true) {
                #$PFObject.GetByKey($item.Code);
                $generalParams = $gs.GetDataInterface([SAPbobsCOM.GeneralServiceDataInterfaces]::gsGeneralDataParams)
                $generalParams.SetProperty('Code', $item.Code);
                $PFObject = $gs.GetByParams($generalParams);
            }
            else {
                #$PFObject.Code = $item.Code;
                $PFObject = $gs.GetDataInterface([SAPbobsCOM.GeneralServiceDataInterfaces]::gsGeneralData);
                $PFObject.SetProperty('Code', [string]$item.Code);
            }
            $PFObject.SetProperty('Name', $item.Name);
            $PFObject.SetProperty('U_ReasonGroup', $item.Group);
            $PFObject.SetProperty('U_Remarks', $item.Remarks);
  
            #Adding or updating depends if object already exists in the database
            if ($exists -eq $true) {
                $dummy = $gs.Update($PFObject);
            }
            else {
                $dummy = $gs.Add($PFObject);
            }
     
            # if ($message -lt 0) {    
            #     $err = $pfcCompany.GetLastErrorDescription()
            #     Throw [System.Exception] ($err)
            # }
        }
        Catch {
            $err = $_.Exception.Message;
            if ($exists -eq $false) {
                $taskMsg = "adding";
            }
            else {
                $taskMsg = "updating"
            }
            $ms = [string]::Format("Error when {0} Downtime Reason with Code {1} Details: {2}", $taskMsg, $item.Code, $err);
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


