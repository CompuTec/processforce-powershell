#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Production Processes
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Production Processes. Script add new or will update existing Production Processes.
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

$opResourceProperties = -join ($csvImportCatalog, "BOMHeader.csv")
$csvBOMRoutingsFilePath = -join ($csvImportCatalog, "BOMRoutings.csv")
$csvBOMRoutingsOperationsFilePath = -join ($csvImportCatalog, "BOMRoutingsOperations.csv")
$csvBOMRoutingsOperationsPropertiesFilePath = -join ($csvImportCatalog, "BOMRoutingsOperationsProperties.csv")
$csvBOMRoutingsOperationsResources = -join ($csvImportCatalog, "BOMRoutingsOperationsResources.csv")
$csvopResourcePropertiesFilePath = -join ($csvImportCatalog, "BOMRoutingsOperationsResourcesProperties.csv")

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
    #region preparing data
    
    #Data loading from a csv file
    [array]$csvItems = Import-Csv -Delimiter ';' -Path $opResourceProperties
    [array]$bomRoutings = Import-Csv -Delimiter ';' -Path $csvBOMRoutingsFilePath
    [array]$bomRoutingsOperations = Import-Csv -Delimiter ';' -Path $csvBOMRoutingsOperationsFilePath

    if ((Test-Path -Path $csvBOMRoutingsOperationsPropertiesFilePath -PathType leaf) -eq $true) {
        [array] $bomRoutingsOperationsProperties = Import-Csv -Delimiter ';' $csvBOMRoutingsOperationsPropertiesFilePath;
    }
    else {
        [array] $bomRoutingsOperationsProperties = $null;
        write-host "BOM Routings Operations Properties - csv not available."
    }

    [array]$bomRoutingsOperationsResources = Import-Csv -Delimiter ';' -Path $csvBOMRoutingsOperationsResources

    if ((Test-Path -Path $csvopResourcePropertiesFilePath -PathType leaf) -eq $true) {
        [array] $opResourceProperties = Import-Csv -Delimiter ';' $csvopResourcePropertiesFilePath;
    }
    else {
        [array] $opResourceProperties = $null;
        write-host "BOM Routings Operations Resources Properties - csv not available."
    }
    write-Host 'Preparing data: ' -NoNewline
    
    $totalRows = $csvItems.Count + $bomRoutings.Count + $bomRoutingsOperations.Count + $bomRoutingsOperationsProperties.Count + $bomRoutingsOperationsResources.Count + $opResourceProperties.Count

    $bomList = New-Object 'System.Collections.Generic.List[array]'

    $dictionaryRoutings = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryRoutingsOperations = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryRoutingsOperationsProperties = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryRoutingsOperationsResources = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
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

    foreach ($row in $csvItems) {
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

  
        $bomList.Add([array]$row);
    }

    foreach ($row in $bomRoutings) {
        $key = $row.BOM_Header + '___' + $row.Revision;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryRoutings.ContainsKey($key)) {
            $list = $dictionaryRoutings[$key];
        }
        else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryRoutings[$key] = $list;
        }
    
        $list.Add([array]$row);
    }

    foreach ($row in $bomRoutingsOperations) {
        $key = $row.BOM_Header + '___' + $row.Revision + '___' + $row.RoutingCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryRoutingsOperations.ContainsKey($key)) {
            $list = $dictionaryRoutingsOperations[$key];
        }
        else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryRoutingsOperations[$key] = $list;
        }
    
        $list.Add([array]$row);
    }

    foreach ($row in $bomRoutingsOperationsProperties) {
        $key = $row.BOM_Header + '___' + $row.Revision + '___' + $row.RoutingCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryRoutingsOperationsProperties.ContainsKey($key)) {
            $list = $dictionaryRoutingsOperationsProperties[$key];
        }
        else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryRoutingsOperationsProperties[$key] = $list;
        }
    
        $list.Add([array]$row);
    }

    
    foreach ($row in $bomRoutingsOperationsResources) {
        $key = $row.BOM_Header + '___' + $row.Revision + '___' + $row.RoutingCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryRoutingsOperationsResources.ContainsKey($key)) {
            $list = $dictionaryRoutingsOperationsResources[$key];
        }
        else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryRoutingsOperationsResources[$key] = $list;
        }
    
        $list.Add([array]$row);
    }
    
    foreach ($row in $opResourceProperties) {
        $key = $row.BOM_Header + '___' + $row.Revision + '___' + $row.RoutingCode;
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
    Write-Host '';
    Write-Host 'Adding/updating data:' -NoNewline;
    #endregion

    if ($bomList.Count -gt 1) {
        $total = $bomList.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;


    foreach ($csvItem in $bomList) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }

            $dictionaryKey = $csvItem.BOM_Header + '___' + $csvItem.Revision;
    

            #Creating BOM object
            $bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"BillOfMaterial");
            #Checking that the BOM already exist
            $retValue = $bom.GetByItemCodeAndRevision($csvItem.BOM_Header, $csvItem.Revision)
            if ($retValue -ne 0) {
                $bom.U_ItemCode = $csvItem.BOM_Header
                $bom.U_Revision = $csvItem.Revision
                $exists = $false;
            }
            else {
                $exists = $true;
            }
            #Data loading from a csv file - Routing
            $bomRoutings = $dictionaryRoutings[$dictionaryKey];
            if ($bomRoutings.count -gt 0) {
                #Deleting all existing routings, operations, resources
        
                $count = $bom.Routings.Count
                for ($i = 0; $i -lt $count; $i++) {
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
                    $dummy = $bom.Routings.DelRowAtPos(0);
                }
        
                $count = $bom.RoutingOperations.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $bom.RoutingOperations.DelRowAtPos(0);
                }
                
                $count = $bom.RoutingOperationResources.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $bom.RoutingOperationResources.DelRowAtPos(0);
                }     
        
                $count = $bom.RoutingOperationProperties.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $bom.RoutingOperationProperties.DelRowAtPos(0);
                }
        
                #Adding a new data - Routings
                foreach ($rtg in $bomRoutings) {  
                    $bom.Routings.U_RtgCode = $rtg.RoutingCode
                    $bom.Routings.U_IsDefault = $rtg.DefaultMRP # Y = Yes, N = No
                    $bom.Routings.U_IsRollUpDefault = $rtg.DefaultCosting # Y = Yes, N = No
                    $dummy = $bom.Routings.Add()
                }
      
                #Deleting defautl operationes copied from Routing
                while ($bom.RoutingOperations.Count -ne 1 ) {  
                    $count = $bom.RoutingOperations.Count
                    $nextint = 0;
        
                    for ($i = 0; $i -lt $count; $i++) {
                        try {
                            $dummy = $bom.RoutingOperations.DelRowAtPos($nextint);
                        }
                        catch {
                            $nextint++
                        }
            
                    }
                }
                $dummy = $bom.RoutingOperations.DelRowAtPos(0);
                $bom.RoutingOperations.SetCurrentLine($bom.RoutingOperations.Count - 1)
       
                $drivers = New-Object 'System.Collections.Generic.Dictionary[String,int]'
                #Adding a new data - Operations for Routings
                foreach ($rtg in $bomRoutings) {
                    $dictionaryKeyRt = $csvItem.BOM_Header + '___' + $csvItem.Revision + '___' + $rtg.RoutingCode;
                    $bomRoutingsOperations = $dictionaryRoutingsOperations[$dictionaryKeyRt];
                    $overlayDict = New-Object 'System.Collections.Generic.Dictionary[int,int]';
                    foreach ($rtgOper in $bomRoutingsOperations) {
                        $bom.RoutingOperations.U_RtgCode = $rtgOper.RoutingCode   
                        $bom.RoutingOperations.U_OprCode = $rtgOper.OperationCode      
                        $bom.RoutingOperations.U_OprSequence = $rtgOper.Sequence

                        if ($rtgOper.OperationOverlayCode -gt '') {
                            $bom.RoutingOperations.U_OprOverlayCode = $rtgOper.OperationOverlayCode;
                            $bom.RoutingOperations.U_OprOverlayId = $overlayDict[$rtgOper.OperationOverlaySequence];
                            $bom.RoutingOperations.U_OprOverlayQty = $rtgOper.OperationOverlayQty;
                        }

                        $overlayDict.Add($rtgOper.Sequence, $bom.RoutingOperations.U_LineNum);
                        $drivers_key = $rtgOper.RoutingCode + '@#@' + $bom.RoutingOperations.U_OprSequence;
                        $drivers.Add($drivers_key, $bom.RoutingOperations.U_RtgOprCode);
                        $dummy = $bom.RoutingOperations.Add()
                    }
			
                    #operation properties
                    $bomRoutingsOperationsProperties = $dictionaryRoutingsOperationsProperties[$dictionaryKeyRt];
                    if ($bomRoutingsOperationsProperties.count -gt 0) {
                        #Deleting all existing properties
                        $count = $bom.RoutingOperationProperties.Count
                        for ($i = 0; $i -lt $count; $i++) {
                            $dummy = $bom.RoutingOperationProperties.DelRowAtPos(0);
                        }
		        
                        #Adding the new data       
                        foreach ($prop in $bomRoutingsOperationsProperties) {
                            $drivers_key = $prop.RoutingCode + '@#@' + $prop.Sequence;
                            $bom.RoutingOperationProperties.U_RtgOprCode = $drivers[$drivers_key];
                            $bom.RoutingOperationProperties.U_RtgCode = $prop.RoutingCode
                            $bom.RoutingOperationProperties.U_OprCode = $prop.OperationCode
                            $bom.RoutingOperationProperties.U_PrpCode = $prop.PropertiesCode
                            $bom.RoutingOperationProperties.U_PrpConType = $prop.Condition #enum ConditionType; Equal EQ = 1, NotEqual NE = 2, GratherThan GT = 3, GratherThanOrEqual GE = 4, LessThan LT = 5, LessThanOrEqual LE = 6, Between BT = 7
                            $bom.RoutingOperationProperties.U_PrpConValue = $prop.Value
                            $bom.RoutingOperationProperties.U_PrpConValueTo = $prop.ToValue
                            $bom.RoutingOperationProperties.U_UnitOfMeasure = $prop.UoM
                            $dummy = $bom.RoutingOperationProperties.Add()
                        }
		        
                    }
			
        
         
                    #Deleting default resources copied from operations for given routing code
                    $count = $bom.RoutingOperationResources.Count
                    for ($i = $count - 1; $i -ge 0; $i--) {
                        $bom.RoutingOperationResources.SetCurrentLine($i);
                        if ($bom.RoutingOperationResources.U_RtgCode -eq $rtg.RoutingCode) {
                            $dummy = $bom.RoutingOperationResources.DelRowAtPos($i);
                        }
                    }
                    $bom.RoutingOperationResources.SetCurrentLine(($bom.RoutingOperationResources.Count - 1));
                    $count = $bom.RoutingsOperationResourceProperties.Count
                    for ($i = $count - 1; $i -ge 0; $i--) {
                        $bom.RoutingsOperationResourceProperties.SetCurrentLine($i);
                        if ($bom.RoutingsOperationResourceProperties.U_RtgCode -eq $rtg.RoutingCode) {
                            $dummy = $bom.RoutingsOperationResourceProperties.DelRowAtPos($i);  
                        }
                    }
                    $bom.RoutingsOperationResourceProperties.SetCurrentLine(($bom.RoutingsOperationResourceProperties.Count - 1));
                    $driversRtgOprRsc = New-Object 'System.Collections.Generic.Dictionary[String,int]'
                    #Adding resources for operations   
        
                    $dictionaryKeyRt = $csvItem.BOM_Header + '___' + $csvItem.Revision + '___' + $rtg.RoutingCode;
                    $bomRoutingsOperationsResources = $dictionaryRoutingsOperationsResources[$dictionaryKeyRt];
                    if ($bomRoutingsOperationsResources.count -gt 0) {
                        foreach ($rtgOperResc in $bomRoutingsOperationsResources) {
                            $drivers_key = $rtgOperResc.RoutingCode + '@#@' + $rtgOperResc.Sequence;
                            $bom.RoutingOperationResources.U_RtgCode = $rtgOperResc.RoutingCode
                            $bom.RoutingOperationResources.U_OprCode = $rtgOperResc.OperationCode
                            $bom.RoutingOperationResources.U_RtgOprCode = $drivers[$drivers_key];
                            $bom.RoutingOperationResources.U_RscCode = $rtgOperResc.ResourceCode

                            if ($rtgOperResc.MachineCode -ne '') {
                                if ($bom.RoutingOperationResources.U_RscType -eq [CompuTec.ProcessForce.API.Enumerators.ResourceType]::Tool) {
                                    $bom.RoutingOperationResources.U_MachineCode = $rtgOperResc.MachineCode;
                                }
                            }

                            $bom.RoutingOperationResources.U_IsDefault = $rtgOperResc.Default
                            $bom.RoutingOperationResources.U_IssueType = $rtgOperResc.IssueType;
                            $bom.RoutingOperationResources.U_OcrCode = $rtgOperResc.DistRule
                            $bom.RoutingOperationResources.U_OcrCode2 = $rtgOperResc.DistRule2
                            $bom.RoutingOperationResources.U_OcrCode3 = $rtgOperResc.DistRule3
                            $bom.RoutingOperationResources.U_OcrCode4 = $rtgOperResc.DistRule4
                            $bom.RoutingOperationResources.U_OcrCode5 = $rtgOperResc.DistRule5
                            $bom.RoutingOperationResources.U_QueueTime = $rtgOperResc.QueTime
                        
                            $queTimeUoM = $rtgOperResc.QueTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3
                            switch ($queTimeUoM) {
                                "1" { $bom.RoutingOperationResources.U_QueueRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedSeconds }
                                "2" { $bom.RoutingOperationResources.U_QueueRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedMinutes }
                                "3" { $bom.RoutingOperationResources.U_QueueRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedHours }

                            }

                            $bom.RoutingOperationResources.U_SetupTime = $rtgOperResc.SetupTime
                            $setupTimeUoM = $rtgOperResc.SetupTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3
                            switch ($setupTimeUoM) {
                                "1" { $bom.RoutingOperationResources.U_SetupRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedSeconds }
                                "2" { $bom.RoutingOperationResources.U_SetupRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedMinutes }
                                "3" { $bom.RoutingOperationResources.U_SetupRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedHours }

                            }

                    

                            $bom.RoutingOperationResources.U_RunTime = $rtgOperResc.RunTime
                            $runtimeUom = $rtgOperResc.RunTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
                            switch ($runtimeUom) {
                                "1" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedSeconds }
                                "2" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedMinutes }
                                "3" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedHours }
                                "4" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::SecondsPerPiece }
                                "5" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::MinutesPerPiece }
                                "6" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::HoursPerPiece }
                                "7" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::PiecesPerSecond }
                                "8" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::PiecesPerMinute }
                                "9" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::PiecesPerHour }

                            }
                            $bom.RoutingOperationResources.U_StockTime = $rtgOperResc.StockTime
                            $stockTimeUoM = $rtgOperResc.StockTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3 
                            switch ($stockTimeUoM) {
                                "1" { $bom.RoutingOperationResources.U_StockRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedSeconds }
                                "2" { $bom.RoutingOperationResources.U_StockRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedMinutes }
                                "3" { $bom.RoutingOperationResources.U_StockRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedHours }

                            }
                            if ($rtgOperResc.NumberOfResources -ne '') {
                                $bom.RoutingOperationResources.U_NrOfResources = $rtgOperResc.NumberOfResources
                            }
					
                            if ($rtgOperResc.HasCycles -ne '') {
						
                                if ($rtgOperResc.HasCycles -eq 'Y') {
							
                                    $bom.RoutingOperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes
							
                                    if ($rtgOperResc.CycleCapacity -ne '') {
                                        $bom.RoutingOperationResources.U_CycleCap = $rtgOperResc.CycleCapacity
                                    }
                                }
                                else {
                                    $bom.RoutingOperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
                                }
                            }
					
                            $bom.RoutingOperationResources.U_Remarks = $rtgOperResc.Remarks
                            if ($rtgOperResc.Project -ne '') {
                                $bom.RoutingOperationResources.U_Project = $rtgOperResc.Project
                            }
					
                            $key = $rtgOperResc.Sequence + '@#@' + $bom.RoutingOperationResources.U_RscCode
					
                            $driversRtgOprRsc.Add($key, $bom.RoutingOperationResources.U_RtgOprRscCode);
                            $dummy = $bom.RoutingOperationResources.Add()
                
                        }
				
                        #Adding resources properties to Operations
                        $dictionaryKeyRt = $csvItem.BOM_Header + '___' + $csvItem.Revision + '___' + $rtg.RoutingCode;
                        $opResourceProperties = $dictionaryResourceProperties[$dictionaryKeyRt];
                        if ($opResourceProperties.count -gt 0) {
                            #Deleting all existing resources
                            $count = $bom.RoutingsOperationResourceProperties.Count - 1
                            if ($count -gt 1) {
                                for ($i = 0; $i -lt $count; $i++) {
			   
			           
                                    $dummy = $bom.RoutingsOperationResourceProperties.DelRowAtPos(0); 
                                }
                            }
			        
                            #Adding the new data
                            foreach ($opResProp in $opResourceProperties) {
			      
                                $key = $opResProp.Sequence + '@#@' + $opResProp.RoutingCode
                                $drivers_key = $opResProp.RoutingCode + '@#@' + $opResProp.Sequence;
					
                                $bom.RoutingsOperationResourceProperties.U_RtgOprCode = $drivers[$drivers_key]
                                $bom.RoutingsOperationResourceProperties.U_RtOpRscCode = $driversRtgOprRsc[$key]
                                $bom.RoutingsOperationResourceProperties.U_RtgCode = $opResProp.RoutingCode
                                $bom.RoutingsOperationResourceProperties.U_OprCode = $opResProp.OperationCode
                                $bom.RoutingsOperationResourceProperties.U_PrpCode = $opResProp.PropertiesCode
                                $bom.RoutingsOperationResourceProperties.U_PrpConType = $opResProp.Condition
                                $bom.RoutingsOperationResourceProperties.U_PrpConValue = $opResProp.Value
                                $bom.RoutingsOperationResourceProperties.U_PrpConValueTo = $opResProp.ToValue
                                $bom.RoutingsOperationResourceProperties.U_UnitOfMeasure = $opResProp.UoM
						
                                #$bom.RoutingsOperationResourceProperties.UDFItems.Item("U_UDF1").Value = $opRes.U_UDF1 # how to add UDF
                                $dummy = $bom.RoutingsOperationResourceProperties.Add()
			            
                            }
                        }
                    }
                }
	
            }
       
            $message = 0
            #Adding or updating BOMs depends on exists in the database
            if ($exists -eq $true) {
                $message = $bom.Update()
            }
            else {
                $message = $bom.Add()
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
            $ms = [string]::Format("Error when {0} Production Process with ItemCode {1} and Revision: {2}. Details: {3}", $taskMsg, $csvItem.BOM_Header, $csvItem.Revision, $err);
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



