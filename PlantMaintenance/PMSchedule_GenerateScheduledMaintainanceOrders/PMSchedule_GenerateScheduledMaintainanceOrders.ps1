#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Test PM Schedule
########################################################################
$SCRIPT_VERSION = "1.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Test PM Schedule. Script triggers creation of Maitenance Orders.
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

#[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")
[System.Reflection.Assembly]::LoadFrom("D:\Praca\ProcessForceGit\Master\Source\CompuTec.ProcessForce.API\bin\x86\Debug\CompuTec.ProcessForce.API.dll")

#endregion

#region Script parameters

$csvImportCatalog = $PSScriptRoot + "\"
#If you are using lower version of PowerShell than PowerShell 4.0 you can use static path
# $csvImportCatalog = "C:\PS\PF\";

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
$xmlParameters = $configurationXml.SelectSingleNode("/configuration/parameters");

$timeBased = $null
$meterBased = $null
$dateTime = $null

$tbString = $xmlParameters.TimeBased
$mbString = $xmlParameters.MeterBased
$dtString = $xmlParameters.CurrentDateTime


    if([bool]::TryParse($tbString, [ref]$timeBased) -ne $true){
        Write-Host -BackgroundColor DarkRed -ForegroundColor Yellow "configuration.xml, parameter TimeBased: Input is not boolean: $tbString"
        return;
    }

    if([bool]::TryParse($mbString, [ref]$meterBased) -ne $true){
        Write-Host -BackgroundColor DarkRed -ForegroundColor Yellow "configuration.xml, parameter MeterBased: Input is not boolean: $mbString"
        return;
    }

try {
    $dateTime = [datetime]$dtString
}
catch {
    $err = $_.Exception.Message;
    Write-Host -BackgroundColor DarkRed -ForegroundColor Yellow "configuration.xml, parameter CurrentDateTime: $err"
    return    
}

$msg = [string]::Format("GenerateScheduledMaintainanceOrders: TimeBased = {0}, MeterBased = {1}, DateTime = {2}", $timeBased, $meterBased, $dateTime);
Write-Host -BackgroundColor Yellow -ForegroundColor Black $msg

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

$DirectAccessState = [CompuTec.Core.DI.Database.DataLayer]::GetLayerState($pfcCompany.Token)
write-host "PF Direct Data Access state: " $DirectAccessState

try {
    $result = $pfcCompany.GenerateScheduledMaintainanceOrders($timeBased, $meterBased, $dateTime);

    $msgKind = "Error"
    $msgBgColorMsg = "DarkRed"
    $msgFontColor = "White"

    if ($result.Success) {
        $msgKind = "OK"
        $msgBgColorMsg = "Green"
        $msgFontColor = "Black"
    }

    Write-Host -BackgroundColor $msgBgColorMsg -ForegroundColor $msgFontColor "Result: $msgKind"

    foreach ($resultError in $result.Errors)
    {
        Write-Host -BackgroundColor Gray -ForegroundColor Black $resultError.Message
    }
}
Catch {
    $err = $_.Exception.Message;
    Write-Host -BackgroundColor DarkRed -ForegroundColor White "Exception occured: $err"
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

