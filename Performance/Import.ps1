using module .\lib\CTLogger.psm1;
Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")
$RESULT_FILE = $PSScriptRoot + "\Results.csv";
[CTLogger] $logJobs = New-Object CTLogger ('DI', 'Import', $RESULT_FILE)
#region connection
[xml] $connectionConfigXml = Get-Content -Encoding UTF8 .\conf\Connection.xml
$xmlConnection = $connectionConfigXml.SelectSingleNode("/CT_CONFIG/Connection");

$logJobs.startSubtask('Import');
$logJobs.startSubtask('Connection');
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany();
$pfcCompany.LicenseServer = $xmlConnection.LicenseServer;
$pfcCompany.SQLServer = $xmlConnection.SQLServer;
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::[string]$xmlConnection.DbServerType;
$pfcCompany.Databasename = $xmlConnection.Database;
$pfcCompany.UserName = $xmlConnection.UserName;
$pfcCompany.Password = $xmlConnection.Password;
    

write-host -backgroundcolor yellow -foregroundcolor black  "Trying connect..."
$version = [CompuTec.Core.CoreConfiguration+DatabaseSetup]::AddonVersion
write-host -backgroundcolor green -foregroundcolor black "PF API Library:" $version';' 'Host:'(Get-WmiObject Win32_OperatingSystem).CSName';' 'OSArchitecture:' (Get-WmiObject Win32_OperatingSystem).OSArchitecture
    
try {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
    $dummy = $pfcCompany.Connect()
    
    write-host -backgroundcolor green -foregroundcolor black "Connected to:" $pfcCompany.SapCompany.CompanyName "/ " $pfcCompany.SapCompany.CompanyDB"" "Sap Company version: " $pfcCompany.SapCompany.Version
}
catch {
    #Show error messages & stop the script
    write-host "Connection Failure: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
    $logJobs.endSubtask('Connection','F',$_.Exception.Message);
    write-host "LicenseServer:" $pfcCompany.LicenseServer
    write-host "SQLServer:" $pfcCompany.SQLServer
    write-host "DbServerType:" $pfcCompany.DbServerType
    write-host "Databasename" $pfcCompany.Databasename
    write-host "UserName:" $pfcCompany.UserName
}

#If company is not connected - stops the script
if (-not $pfcCompany.IsConnected) {
    write-host -backgroundcolor yellow -foregroundcolor black "Company is not connected"
    $logJobs.endSubtask('Connection','F','Company is not connected');
    return 
}
$logJobs.endSubtask('Connection','S','');
$sapCompany = $pfcCompany.SapCompany;
#endregion


function importIMD($sapCompany) {
    [CTLogger] $logIMD = New-Object CTLogger ('DI', 'Import Item Master Data', $RESULT_FILE)
    #region import of Item Master Data
    [xml] $IMDConfigXml = Get-Content -Encoding UTF8 .\conf\ItemMasterData.xml

    $xmlItems = $IMDConfigXml.SelectSingleNode([string]::Format("/CT_CONFIG/ItemMasterData"));

    $numberOfItems = [int] $xmlItems.NumberOfItems
    $itemCodeLength = ([string]$numberOfItems).Length;
    $itemPrefix = [string] $xmlItems.Prefix
    $warehouseCode = [string] $xmlItems.WarehouseCode
    
    for ($i = 0; $i -lt $numberOfItems; $i++) {
        try {
            $logIMD.startSubtask('Preparing to add Item Master Data to SAP');
            $sapIMD = $sapCompany.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oItems);
        
            $ItemCode = $itemPrefix + ([string]$i).PadLeft($itemCodeLength, '0');

            $retValue = $sapIMD.GetByKey($ItemCode)

            if ($retValue -eq $true) {
                $logIMD.endSubtask('Preparing to add Item Master Data to SAP','S','Item Already Exists');
                continue;
            }

            $sapIMD.ItemCode = $ItemCode;
            $sapIMD.ItemName = $ItemCode;
            $sapIMD.WhsInfo.WarehouseCode = $warehouseCode;
            $sapIMD.DefaultWarehouse = $warehouseCode;
            $logIMD.endSubtask('Preparing to add Item Master Data to SAP','S','');
            $logIMD.startSubtask('Adding Item Master Data to SAP');
            $message = $sapIMD.Add();
            
            if ($message -lt 0) {
                $err = $sapCompany.GetLastErrorDescription();
                Throw [System.Exception] ($err);
            }
            $logIMD.endSubtask('Adding Item Master Data to SAP','S','');
        }
        Catch {
            $err = $_.Exception.Message;
            $logIMD.endSubtask('Adding Item Master Data to SAP','F',$err);
            continue;
        }
    }
    #endregion
}
$logJobs.startSubtask('Import IMD');
importIMD $sapCompany
$logJobs.endSubtask('Import IMD','S','');



$logJobs.endSubtask('Import','S','');