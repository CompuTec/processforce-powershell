#region #PF API library usage
Clear-Host
# You need to check in what architecture PowerShell ISE is running (x64 or x86),
# you need run PowerShell ISE in the same architecture like PF API is installed (check in Windows -> Programs & Features)
# Examples:
#     SAP Client x64 + PF x64 installed on DB/Company => PF API x64 => Windows PowerShell ISE
#     SAP Client x86 + PF x86 installed on DB/Company => PF API x86 => Windows PowerShell ISE x86

[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#endregion

#region Script parameters

$csvImportCatalog = $PSScriptRoot + "\";
 
$csvoperationsFilePath = -join ($csvImportCatalog, "Operations.csv")
$csvOperationPropertiesFilePath = -join ($csvImportCatalog, "Operations_Properties.csv")
$csvOperationResourcesFilePath = -join ($csvImportCatalog, "Operations_Resources.csv")
$csvOperationResourcesPropertiesFilePath = -join ($csvImportCatalog, "Operations_ResourcesProperties.csv")
 
#endregion

#region #Datbase/Company connection settings
 
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.LicenseServer = "10.0.0.203:40000"
$pfcCompany.SQLServer = "10.0.0.202:30115"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"
$pfcCompany.Databasename = "PFDEMOGB_MACIEJP"
$pfcCompany.UserName = "maciejp"
$pfcCompany.Password = "1234"
 
# where:
 
# LicenseServer = SAP LicenceServer name or IP Address with port number (see in SAP Client -> Administration -> Licence -> Licence Administration -> Licence Server)
# SQLServer     = Server name or IP Address with port number, should be the same like in System Landscape Dirctory (see https://<Server>:<Port>/ControlCenter) - sometimes best is use IP Address for resolve connection problems.
#
# DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2014"     # For MsSQL Server 2014
#                [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2012"     # For MsSQL Server 2012
#                [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"        # For HANA
#
# Databasename = Database / schema name (check in SAP Company select form/window, or in MsSQL Management Studio or in HANA Studio)
# UserName     = SAP user name ex. manager
# Password     = SAP user password
 
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
 
#endregion

#Data loading from a csv file
write-host ""
[array]$csvOperations = Import-Csv -Delimiter ';' -Path $csvOperationsFilePath
[array]$csvOperationProperties = Import-Csv -Delimiter ';' -Path $csvOperationPropertiesFilePath
[array]$csvOperationResources = Import-Csv -Delimiter ';' -Path $csvOperationResourcesFilePath
[array]$csvOperationResourcesProperties = Import-Csv -Delimiter ';' -Path $csvOperationResourcesPropertiesFilePath 

write-Host 'Preparing data: '
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
        $count = $operation.OperationResources.Count -1
        for ($i = $count-1; $i -ge 0; $i--) {
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
        $count = $operation.OperationResourceProperties.Count -1 
        for ($i = $count-1; $i -ge 0; $i--) {
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
    if ($retValue -eq 0) {
        
        #[System.String]::Format("Updating Opertion: {0}", $csvItem.OperationCode)
        $message = $operation.Update()
    }
    else {
        try {
            #[System.String]::Format("Adding Operation: {0}", $csvItem.OperationCode)
            $message = $operation.Add()
        }
        catch [Exception] {
            Write-Host $_.Exception.InnerException.ToString()
        }
    }
    if ($message -lt 0) {    
        $err = $pfcCompany.GetLastErrorDescription()
        write-host -backgroundcolor red -foregroundcolor white $err
    }    
}
