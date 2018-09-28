#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Operations
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Operations. Script add new or will update existing Operations.
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

$csvoperationsFilePath = -join ($csvImportCatalog, "Operations.csv")
$csvOperationPropertiesFilePath = -join ($csvImportCatalog, "OperationsProperties.csv")
$csvOperationResourcesFilePath = -join ($csvImportCatalog, "OperationsResources.csv")
$csvOperationResourcesPropertiesFilePath = -join ($csvImportCatalog, "OperationsResourcesProperties.csv")

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
    write-host ""
    [array]$csvOperations = Import-Csv -Delimiter ';' -Path $csvOperationsFilePath
    
    if ((Test-Path -Path $csvOperationPropertiesFilePath -PathType leaf) -eq $true) {
        [array] $csvOperationProperties = Import-Csv -Delimiter ';' $csvOperationPropertiesFilePath;
    }
    else {
        [array] $csvOperationProperties = $null;
        write-host "Operations Properties - csv not available."
    }
    [array]$csvOperationResources = Import-Csv -Delimiter ';' -Path $csvOperationResourcesFilePath

    [array]$csvOperationResourcesProperties = Import-Csv -Delimiter ';' -Path $csvOperationResourcesPropertiesFilePath 
    if ((Test-Path -Path $csvOperationResourcesPropertiesFilePath -PathType leaf) -eq $true) {
        [array] $csvOperationResourcesProperties = Import-Csv -Delimiter ';' $csvOperationResourcesPropertiesFilePath;
    }
    else {
        [array] $csvOperationResourcesProperties = $null;
        write-host "Operation Resources Properties - csv not available."
    }

    write-Host 'Preparing data: ' -NoNewline
    $totalRows = $csvOperations.Count + $csvOperationProperties.Count + $csvOperationResources.Count + $csvOperationResourcesProperties.Count;
    
    $operationsList = New-Object 'System.Collections.Generic.List[array]';
    $dictionaryOperationsProperties = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryOperationsResources = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryResourceProperties = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
    
    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $csvOperations) {
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        $operationsList.Add([array]$row);
    }

    foreach ($row in $csvOperationProperties) {
        $key = $row.OperationCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryOperationsProperties.ContainsKey($key)) {
            $list = $dictionaryOperationsProperties[$key];
        }
        else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryOperationsProperties[$key] = $list;
        }
    
        $list.Add([array]$row);
    }

    foreach ($row in $csvOperationResources) {
        $key = $row.OperationCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryOperationsResources.ContainsKey($key)) {
            $list = $dictionaryOperationsResources[$key];
        }
        else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryOperationsResources[$key] = $list;
        }
    
        $list.Add([array]$row);
    }
    
    foreach ($row in $csvOperationResourcesProperties) {
        $key = $row.OperationCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryResourceProperties.ContainsKey($key)) {
            $list = $dictionaryResourceProperties[$key];
        }
        else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryResourceProperties[$key] = $list;
        }
    
        $list.Add([array]$row);
    }


    write-host ""
    Write-Host 'Adding/updating data: ' -NoNewline;

    $totalRows = $operationsList.Count;
    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
    
    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($csvItem in $operationsList) {
        try {
            $progressItterator++;
            $progres = [math]::Round(($progressItterator * 100) / $total);
            if ($progres -gt $beforeProgress) {
                Write-Host $progres"% " -NoNewline
                $beforeProgress = $progres
            }
            $key = $csvItem.OperationCode;
            #Creating Operation object
            $operation = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"Operation")
            #Checking that the operation already exist  
            $retValue = $operation.GetByOprCode($csvItem.OperationCode)
            if ($retValue -ne 0) {     
                #Adding the new data
                $operation.U_OprCode = $csvItem.OperationCode
                $exists = $false;
            }
            else {
                $exists = $true;
            }
            #Data loading from a csv file - Operation Properties
            $operation.U_OprName = $csvItem.OperationName
            $operation.U_Remarks = $csvItem.Remarks
    
            $operProps = $dictionaryOperationsProperties[$key];
            if ($operProps.count -gt 0) {
                #Deleting all existing properties
                $count = $operation.OperationProperties.Count
                for ($i = 0; $i -lt $count; $i++) {
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
                    $dummy = $operation.OperationProperties.DelRowAtPos(0);
                }
        
                #Adding the new data       
                foreach ($prop in $operProps) {
                    $operation.OperationProperties.U_PrpCode = $prop.PropertiesCode
                    $operation.OperationProperties.U_PrpConType = $prop.Condition #enum ConditionType; Equal EQ = 1, NotEqual NE = 2, GratherThan GT = 3, GratherThanOrEqual GE = 4, LessThan LT = 5, LessThanOrEqual LE = 6, Between BT = 7
                    $operation.OperationProperties.U_PrpConValue = $prop.Value
                    $operation.OperationProperties.U_PrpConValueTo = $prop.ToValue
                    $operation.OperationProperties.U_UnitOfMeasure = $prop.UoM
                    $dummy = $operation.OperationProperties.Add()
                }
        
            }
    
            #Adding resources to Operations
            $opResource = $dictionaryOperationsResources[$key];
            if ($opResource.count -gt 0) {
                #Deleting all existing resources
                $count = $operation.OperationResources.Count - 1
                for ($i = $count - 1; $i -ge 0; $i--) {
                    $dummy = $operation.OperationResources.DelRowAtPos($i); 
                }
                $resourcesDict = New-Object 'System.Collections.Generic.Dictionary[String,int]'
                #Adding the new data
                foreach ($opRes in $opResource) {
      
                    $operation.OperationResources.U_RscCode = $opRes.ResourceCode
                    $operation.OperationResources.U_IsDefault = $opRes.Default # Y = Yes, N = No
                    $operation.OperationResources.U_QueueTime = $opRes.QueTime
                    $operation.OperationResources.U_QueueRate = $opRes.QueTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
                    $operation.OperationResources.U_SetupTime = $opRes.SetupTime
                    $operation.OperationResources.U_SetupRate = $opRes.SetupTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
                    $operation.OperationResources.U_RunTime = $opRes.RunTime
                    $operation.OperationResources.U_RunRate = $opRes.RunTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
                    $operation.OperationResources.U_StockTime = $opRes.StockTime		
                    $operation.OperationResources.U_StockRate = $opRes.StockTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
            
                    if ($opRes.MachineCode -ne '') {
                        if ($operation.OperationResources.U_RscType -eq [CompuTec.ProcessForce.API.Enumerators.ResourceType]::Tool) {
                            $operation.OperationResources.U_MachineCode = $opRes.MachineCode
                        }
                    }

                    if ($opRes.Cycles -eq 'Y') {
                        $operation.OperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes
                        $operation.OperationResources.U_CycleCap = $opRes.CycleCapacity
                    }
                    else {
                        $operation.OperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
                    }
                    $operation.OperationResources.U_Project = $opRes.Project
                    $operation.OperationResources.U_OcrCode = $opRes.Dimension1
                    $operation.OperationResources.U_OcrCode2 = $opRes.Dimension2
                    $operation.OperationResources.U_OcrCode3 = $opRes.Dimension3
                    $operation.OperationResources.U_OcrCode4 = $opRes.Dimension4
                    $operation.OperationResources.U_OcrCode5 = $opRes.Dimension5
			
			
                    $operation.OperationResources.U_Remarks = $opRes.Remarks
                    #$operation.OperationResources.UDFItems.Item("U_UDF1").Value = $opRes.U_UDF1 # how to add UDF
			
                    $resourcesDict.Add($operation.OperationResources.U_RscCode, $operation.OperationResources.U_OprRscCode);
                    $dummy = $operation.OperationResources.Add()
            
                }
            }
	
            #Adding resources properties to Operations
            $opResourceProperties = $dictionaryResourceProperties[$key];
            if ($opResourceProperties.count -gt 0) {
                #Deleting all existing resources
                $count = $operation.OperationResourceProperties.Count - 1 
                for ($i = $count - 1; $i -ge 0; $i--) {
                    $dummy = $operation.OperationResourceProperties.DelRowAtPos($i); 
                }
        
                #Adding the new data
                foreach ($opResProp in $opResourceProperties) {
      
                    $operation.OperationResourceProperties.U_OprCode = $opResProp.OperationCode
                    $operation.OperationResourceProperties.U_OprRscCode = $resourcesDict[$opResProp.ResourceCode]
                    $operation.OperationResourceProperties.U_PrpCode = $opResProp.PropertiesCode
                    $operation.OperationResourceProperties.U_PrpConType = $opResProp.Condition
                    $operation.OperationResourceProperties.U_PrpConValue = $opResProp.Value
                    $operation.OperationResourceProperties.U_PrpConValueTo = $opResProp.ToValue
                    $operation.OperationResourceProperties.U_UnitOfMeasure = $opResProp.UoM
			
                    #$operation.OperationResourceProperties.UDFItems.Item("U_UDF1").Value = $opRes.U_UDF1 # how to add UDF
                    $dummy = $operation.OperationResourceProperties.Add()
            
                }
            }
  
            $message = 0
    
            #Adding or updating Operations depends on exists in the database
            if ($exists -eq $true) {
                $message = $operation.Update()
            }
            else {
                try {
                    #[System.String]::Format("Adding Operation: {0}", )
                    $message = $operation.Add()
                }
                catch [Exception] {
                    Write-Host $_.Exception.InnerException.ToString()
                }
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
            $ms = [string]::Format("Error when {0} Operations with Code {1} Details: {2}", $taskMsg, $csvItem.OperationCode, $err);
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

