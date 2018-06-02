Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany();
$pfcCompany.LicenseServer = "10.0.0.203:40000";
$pfcCompany.SQLServer = "10.0.0.202:30115";
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB";
$pfcCompany.Databasename = 'PFDEMOGB_MACIEJP';
$pfcCompany.UserName = "maciejp";
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

$UDO_DOWNTIME_REASON_GROUP = 'CT_PF_DTReaonGroup';

#Data loading from a csv file
[array] $csvDowntimeReasonGroups = Import-Csv -Delimiter ';' -Path  ($PSScriptRoot + "\DowntimeReasonGroups.csv")
 
#region preparing data
write-Host 'Preparing data: '
$totalRows = $csvDowntimeReasonGroups.Count;

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

foreach ($row in $csvDowntimeReasonGroups) {
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
    $progressItterator++;
    $progress = [math]::Round(($progressItterator * 100) / $total);
    if ($progres -gt $beforeProgress) {
        Write-Host $progress"% " -NoNewline
        $beforeProgress = $progress
    }
    try {
        $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]::"BoRecordset")
     
        $rs.DoQuery([string]::Format( "SELECT T0.""Code"" FROM ""@CT_PF_ODTRG"" T0 WHERE T0.""Code"" = N'{0}'", $item.Code))
        $exists = 0;
        if ($rs.RecordCount -gt 0) {
            $exists = 1
        }
  
        #Creating PF Object
        #$PFObject = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::DownTimeReasonGroup);
        $Company = $pfcCompany.SapCompany;
        
        $cs = $Company.GetCompanyService();
        $gs = $cs.GetGeneralService($UDO_DOWNTIME_REASON_GROUP);
                  
        if ($exists -eq 1) {
            $generalParams = $gs.GetDataInterface([SAPbobsCOM.GeneralServiceDataInterfaces]::gsGeneralDataParams)
            $generalParams.SetProperty('Code', $item.Code);
            $PFObject = $gs.GetByParams($generalParams);
        }
        else {
            $PFObject = $gs.GetDataInterface([SAPbobsCOM.GeneralServiceDataInterfaces]::gsGeneralData);
            $PFObject.SetProperty('Code', [string]$item.Code);
        }
        $PFObject.SetProperty('Name', [string]$item.Name);

    
        #Adding or updating depends if object already exists in the database
        if ($exists -eq 1) {
            [System.String]::Format("Updating Downtime Group: {0}", $item.Code);
            $dummy = $gs.Update($PFObject);
        }
        else {
            [System.String]::Format("Adding Downtime Group: {0}", $item.Code);
            $dummy = $gs.Add($PFObject);
        }
        write-host "Success";

    } 
    Catch {
        $err = $_.Exception.Message;
        $content = [string]::Format("Error occured for Code {0}: {1}", $item.Code, $err);
        Write-Host -BackgroundColor DarkRed -ForegroundColor White $content;
        continue;
    }
    
}

