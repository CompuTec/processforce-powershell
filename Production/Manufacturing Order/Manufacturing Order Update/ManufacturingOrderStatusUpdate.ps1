#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Manufacturing Order Status Update - tutorial 
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Base on SQL Query this script will update Manufacturing Order Status.
#      SQL Query can be easily changed - it must return three columns: 
#		* DocEntry - Manufacturing Order DocEntry
#		* DocNum - Manufacuturing Order Document Number 
#		* StatusCode - Manufacuturing Order Status
# Troubleshooting:
#   https://connect.computec.pl/display/PF930EN/FAQ+PowerShell
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

#region #Datbase/Company connection settings
$csvImportCatalog = $PSScriptRoot + "\"
#If you are using lower version of PowerShell than PowerShell 4.0 you can use static path
# $csvImportCatalog = "C:\PS\PF\";

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
    $SQLQuery = "SELECT ""DocEntry"", ""DocNum"", 'CL' AS ""StatusCode"" FROM ""@CT_PF_OMOR"" WHERE ""U_RequiredDate"" < '2018-07-01' AND ""U_Status"" = 'FI' ";
    $queryManager = New-Object 'CompuTec.Core.DI.Database.QueryManager'
    $queryManager.CommandText = $SQLQuery;
    $recordSet = $queryManager.Execute($pfcCompany.Token);
    $recordCount = $recordSet.RecordCount;
	
    if ($recordCount -gt 0 ) {
        while (!$recordSet.EoF) {
            try {
                $DocEntry = $recordSet.Fields.Item('DocEntry').Value;
                $DocNum = $recordSet.Fields.Item('DocNum').Value;
                $Status = $recordSet.Fields.Item('StatusCode').Value;
                $mo = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ManufacturingOrder)

                $result = $mo.GetByKey($DocEntry);
                if ($result -ne 0) {
                    $err = [string]::Format("Manufacturing Order with DocEntry:{0}, DocNum:{1} don't exists", [string]$DocEntry, [string]$DocNum);
                    Throw [System.Exception]($err);
                }				

                switch ($Status) {
                    "NS" {
                        $MORstatus = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::NotScheduled
                        break;
                    }
                    "SC" {
                        $MORstatus = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Scheduled
                        break;
                    }
                    "RL" {
                        $MORstatus = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Released
                        break;
                    }
                    "ST" {
                        $MORstatus = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Started
                        break;
                    }
                    "FI" {
                        $MORstatus = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Finished
                        break;
                    }
                    "CL" {
                        $MORstatus = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Closed
                        break;
                    }
                    default {
                        $err = [string]::Format("Incorrect Status Code {0}. Possible values are: NS, SC, RL, ST, FI, CL", [string]$Status);
                        Throw [System.Exception]($err);
                        break;
                    }
                }
                $mo.U_Status = $MORstatus;
                $updateResult = $mo.Update();
    
                if ($updateResult -lt 0) {    
                    $err = $pfcCompany.GetLastErrorDescription()
                    Throw [System.Exception] ($err)
                }	
            }
            catch {
                $err = $_.Exception.Message;
                $ms = [string]::Format("Error when updating Manufacturing Order. DocEntry:{0}, DocNum:{1}. Details: {2}", [string]$DocEntry, [string]$DocNum, $err);
                Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
            }
            Finally {
                $recordSet.MoveNext();
            }
        }
		
        Write-Host -BackgroundColor Yellow -ForegroundColor Blue ([string]::Format("{0} Manufacturing Orders updated. Operation completed", [string]$recordCount));
    }
    else {
        Write-Host -BackgroundColor Yellow -ForegroundColor Blue "SQL Query didn't return any records"
    }
}
Catch {
    $err = $_.Exception.Message;
    $ms = [string]::Format("Exception occured:{0}", $err);
    Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
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