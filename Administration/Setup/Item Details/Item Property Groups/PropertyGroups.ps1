#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Properties Groups
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Properites Groups. Script add new or will update existing Groups.
#      You need to have all requred files for import.
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

[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#endregion

#region Script parameters

$csvImportCatalog = $PSScriptRoot + "\"
#If you are using lower version of PowerShell than PowerShell 4.0 you can use static path
# $csvImportCatalog = "C:\PS\PF\";

$csvPropertyGroupsPath = -join ($csvImportCatalog, "PropertyGroups.csv")
$csvPropertySubgroupsPath = -join ($csvImportCatalog, "PropertySubgroups.csv")

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
    [array] $csvPropertyGroups = Import-Csv -Delimiter ';' $csvPropertyGroupsPath;
	
    if ((Test-Path -Path $csvPropertySubgroupsPath -PathType leaf) -eq $true) {
        [array] $csvPropertySubgroups = Import-Csv -Delimiter ';' $csvPropertySubgroupsPath;
    }
    else {
        write-host "Property Subgroups - csv not available."
    }

    write-Host 'Preparing data: '
    $totalRows = $csvPropertyGroups.Count + $csvPropertySubgroups.Count;
    $propGroupList = New-Object 'System.Collections.Generic.List[array]'
    $propSubGroupDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
	
    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $csvPropertyGroups) {
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        $propGroupList.Add([array]$row);
    }

    foreach ($row in $csvPropertySubgroups) {
        $key = $row.GroupCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
	
        if ($propSubGroupDict.ContainsKey($key) -eq $false) {
            $propSubGroupDict[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $propSubGroupDict[$key];
		
        $list.Add([array]$row);
    }

    Write-Host '';
    Write-Host 'Adding/updating data: ' -NoNewLine;


    $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
    if ($propGroupList.Count -gt 1) {
        $total = $propGroupList.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;
    foreach ($grp in $propGroupList) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OIPG"" WHERE ""U_GrpCode"" = N'{0}'", $grp.GroupCode));
	
            #Creating Property Group object
            $propertyGrp = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::"ItemPropertyGroup")
            #Checking that the group already exist
            if ($rs.RecordCount -gt 0) {
                $dummy = $propertyGrp.GetByKey($rs.Fields.Item(0).Value);
                $exists = $true
            }
            else {
                $propertyGrp.U_GrpCode = $grp.GroupCode;
                $exists = $false
            }
   
            $propertyGrp.U_GrpName = $grp.GroupName;
            $propertyGrp.U_GrpDescription = $grp.Remarks;
	
            #Data loading from the csv file - Subgroups for Property Group
            [array]$subGrps = $propSubGroupDict[$grp.GroupCode];
            if ($subGrps.Count -gt 0) {
                #Deleting all exisitng Revisions
				$count = $propertyGrp.Subgroups.Count
				for ($i = 1; $i -lt $count; $i++) {
					$dummy = $propertyGrp.Subgroups.DelRowAtPos(1)
				}

				$propertyGrp.Subgroups.SetCurrentLine(0);
				$propertyGrp.Subgroups.U_SubGrpCode = '';
         
                #Adding Subgroup
                foreach ($subGrp in $subGrps) {
                    $propertyGrp.Subgroups.U_SubGrpCode = $subGrp.SubgroupCode
                    $propertyGrp.Subgroups.U_SubGrpName = $subGrp.SubgroupName
                    $propertyGrp.Subgroups.U_SubGrpDescription = $subGrp.SubgroupRemarks
                    $dummy = $propertyGrp.Subgroups.Add();
                }
            }
            $message = 0
            #Adding or updating Property Groups depends on exists in the database
            if ($exists -eq $true) {
                $message = $propertyGrp.Update()
            }
            else {
                $message = $propertyGrp.Add()
            }
	
            if ($message -lt 0) {    
                $err = $pfcCompany.GetLastErrorDescription()
                Throw [System.Exception] ($err)
            }
        }
        Catch {
            $err = $_.Exception.Message;
            if ($exists -eq $true) {
                $taskMsg = "adding";
            }
            else {
                $taskMsg = "updating"
            }
            $ms = [string]::Format("Error when {0} Property Group with Code {1} Details: {2}", $taskMsg, $grp.GroupCode, $err);
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

