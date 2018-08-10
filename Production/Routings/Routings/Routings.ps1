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
 
$csvRoutingsFilePath = -join ($csvImportCatalog, "Routings.csv")
$csvRoutingOperationsFilePath = -join ($csvImportCatalog, "Routing_Operations.csv")
$csvRoutingOperationPropertiesFilePath = -join ($csvImportCatalog, "Routing_Operations_Properties.csv")
$csvRoutingOperationResourcesFilePath = -join ($csvImportCatalog, "Routing_Operations_Resources.csv")
$csvRoutingOperationResourcesPropertiesFilePath = -join ($csvImportCatalog, "Routing_Operations_Resources_Properties.csv")
 
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
[array]$csvRoutings = Import-Csv -Delimiter ';' -Path $csvRoutingsFilePath;
[array]$csvRoutingOperations = Import-Csv -Delimiter ';' -Path $csvRoutingOperationsFilePath
[array]$csvRoutingOperationProperties = Import-Csv -Delimiter ';' -Path $csvRoutingOperationPropertiesFilePath
[array]$csvRoutingOperationResources = Import-Csv -Delimiter ';' -Path $csvRoutingOperationResourcesFilePath
[array]$csvRoutingOperationResourcesProperties = Import-Csv -Delimiter ';' -Path $csvRoutingOperationResourcesPropertiesFilePath 

write-Host 'Preparing data: '
$totalRows = $csvRoutings.Count + $csvRoutingOperations.Count + $csvRoutingOperationProperties.Count + $csvRoutingOperationResources.Count + $csvRoutingOperationResourcesProperties.Count;
    
$routingsList = New-Object 'System.Collections.Generic.List[array]';
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

foreach ($row in $csvRoutings) {
    $progressItterator++;
    $progres = [math]::Round(($progressItterator * 100) / $total);
    if ($progres -gt $beforeProgress) {
        Write-Host $progres"% " -NoNewline
        $beforeProgress = $progres
    }
    $routingsList.Add([array]$row);
}

foreach ($row in $csvRoutingOperations) {
    $key = $row.RoutingCode;
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

foreach ($row in $csvRoutingOperationProperties) {
    $key = $row.RoutingCode;
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

foreach ($row in $csvRoutingOperationResources) {
    $key = $row.RoutingCode;
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
    
foreach ($row in $csvRoutingOperationResourcesProperties) {
    $key = $row.RoutingCode;
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
foreach ($csvItem in $routingsList) {
    $key = $csvItem.RoutingCode;
    #Creating Operation object
    $routing = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Routing)
    #Checking that the operation already exist    
    $retValue = $routing.GetByRtgCode($csvItem.RoutingCode)
    if ($retValue -ne 0) { 
        #Adding the new data
        $routing.U_RtgCode = $csvItem.RoutingCode
    }
    $routing.U_RtgName = $csvItem.RoutingName
    $routing.U_Active = $csvItem.Active #enum type; 1 = Yes, 2 = No
    $routing.U_Remarks = $csvItem.Remarks
    #Data loading from a csv file - Routing Operations
    #[array]$routingOperations = Import-Csv -Delimiter ';' -Path "C:\Routing_Operations.csv" | Where-Object {$_.RoutingCode -eq $csvItem.RoutingCode}
    $routingOperations = $dictionaryRoutingsOperations[$key];

    if ($routingOperations.count -gt 0) {
        #Deleting all existing operations
        $count = $routing.Operations.Count
        for ($i = 0; $i -lt $count; $i++) {
            $dummy = $routing.Operations.DelRowAtPos(0);
        }
        $count = $routing.OperationResourceProperties.Count
        for ($i = 0; $i -lt $count; $i++) {
            $dummy = $routing.OperationResourceProperties.DelRowAtPos(0);
        }
        $drivers = New-Object 'System.Collections.Generic.Dictionary[String,int]'
        #Adding the new data       
        foreach ($rtOper in $routingOperations) {
            $routing.Operations.U_OprCode = $rtOper.OperationCode
            $routing.Operations.U_OprOverlayId = $rtOper.OverlayID
            $routing.Operations.U_OprOverlayCode = $rtOper.OperationOverCode
            $routing.Operations.U_OprOverlayQty = $rtOper.OverlayQty
            $routing.Operations.U_OprSequence = $rtOper.Sequence
            $routing.Operations.U_Remarks = $rtOper.Remarks
            $drivers.Add($routing.Operations.U_OprSequence, $routing.Operations.U_RtgOprCode);
            $dummy = $routing.Operations.Add()
        }
		
        #operation properties
        #[array]$routingsOperationsProperties = Import-Csv -Delimiter ';' -Path "C:\Routing_Operations_Properties.csv" | Where-Object {$_.RoutingCode -eq $csvItem.RoutingCode}	
        $routingsOperationsProperties = $dictionaryRoutingsOperationsProperties[$key];
        if ($routingsOperationsProperties.count -gt 0) {
            #Deleting all existing properties
            $count = $routing.OperationProperties.Count
            for ($i = 0; $i -lt $count; $i++) {
                $dummy = $routing.OperationProperties.DelRowAtPos(0);
            }
		        
            #Adding the new data       
            foreach ($prop in $routingsOperationsProperties) {
                $routing.OperationProperties.U_RtgOprCode = $drivers[$prop.Sequence]
                $routing.OperationProperties.U_PrpCode = $prop.PropertiesCode
                $routing.OperationProperties.U_PrpConType = $prop.Condition #enum ConditionType; Equal EQ = 1, NotEqual NE = 2, GratherThan GT = 3, GratherThanOrEqual GE = 4, LessThan LT = 5, LessThanOrEqual LE = 6, Between BT = 7
                $routing.OperationProperties.U_PrpConValue = $prop.Value
                $routing.OperationProperties.U_PrpConValueTo = $prop.ToValue
                $routing.OperationProperties.U_UnitOfMeasure = $prop.UoM
                $dummy = $routing.OperationProperties.Add()
            }
        }
		
        #Deleting default resources copied from operations   
        $count = $routing.OperationResources.Count
        for ($i = $count-1; $i -ge 0; $i--) {
            $dummy = $routing.OperationResources.DelRowAtPos($i);
        }    
        $count = $routing.OperationResourceProperties.Count-1
        for ($i = $count-1; $i -ge 0; $i--) {
            $dummy = $routing.OperationResourceProperties.DelRowAtPos($i);      
        }
        $driversOprRsc = New-Object 'System.Collections.Generic.Dictionary[String,int]'
        #Adding resources for operations   
        
        #[array]$routingsOperationsResources = Import-Csv -Delimiter ';' -Path "C:\Routing_Operations_Resources.csv" | Where-Object {$_.RoutingCode -eq $csvItem.RoutingCode}
        $routingsOperationsResources = $dictionaryRoutingsOperationsResources[$key];
        if ($routingsOperationsResources.count -gt 0) {
            foreach ($rtgOperResc in $routingsOperationsResources) {
                $routing.OperationResources.U_RtgOprCode = $drivers[$rtgOperResc.Sequence];
                $routing.OperationResources.U_RscCode = $rtgOperResc.ResourceCode

                if($rtgOperResc.MachineCode -ne ''){
                    if($routing.OperationResources.U_RscType -eq [CompuTec.ProcessForce.API.Enumerators.ResourceType]::Tool){
                        $routing.OperationResources.U_MachineCode = $rtgOperResc.MachineCode;
                    }
                }

                $routing.OperationResources.U_OcrCode = $rtgOperResc.OcrCode
                $routing.OperationResources.U_OcrCode2 = $rtgOperResc.OcrCode2
                $routing.OperationResources.U_OcrCode3 = $rtgOperResc.OcrCode3
                $routing.OperationResources.U_OcrCode4 = $rtgOperResc.OcrCode4
                $routing.OperationResources.U_OcrCode5 = $rtgOperResc.OcrCode5
                $routing.OperationResources.U_IsDefault = $rtgOperResc.Default
                $routing.OperationResources.U_IssueType = $rtgOperResc.IssueType
                $routing.OperationResources.U_QueueTime = $rtgOperResc.QueTime
                $routing.OperationResources.U_QueueRate = $rtgOperResc.QueTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
                $routing.OperationResources.U_SetupTime = $rtgOperResc.SetupTime
                $routing.OperationResources.U_SetupRate = $rtgOperResc.SetupTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
                $routing.OperationResources.U_RunTime = $rtgOperResc.RunTime
                $routing.OperationResources.U_RunRate = $rtgOperResc.RunTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
                $routing.OperationResources.U_StockTime = $rtgOperResc.StockTime
                $routing.OperationResources.U_StockRate = $rtgOperResc.StockTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
					
					
					
                if ($rtgOperResc.HasCycles -ne '') {
						
                    if ($rtgOperResc.HasCycles -eq 'Y') {
							
                        $routing.OperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes
							
                        if ($rtgOperResc.CycleCapacity -ne '') {
                            $routing.OperationResources.U_CycleCap = $rtgOperResc.CycleCapacity
                        }
                    }
                    else {
                        $routing.OperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
                    }
                }
					
                $routing.OperationResources.U_Remarks = $rtgOperResc.Remarks
                if ($rtgOperResc.Project -ne '') {
                    $routing.OperationResources.U_Project = $rtgOperResc.Project
                }
					
                $key = $rtgOperResc.Sequence + '@#@' + $routing.OperationResources.U_RscCode
					
                $driversOprRsc.Add($key, $routing.OperationResources.U_RtgOprRscCode);
                $dummy = $routing.OperationResources.Add()
                
            }
				
            #Adding resources properties to Operations
            #[array]$opResourceProperties = Import-Csv -Delimiter ';' -Path "C:\Routing_Operations_Resources_Properties.csv" | Where-Object { $_.RoutingCode -eq $csvItem.RoutingCode}
            $opResourceProperties = $dictionaryResourceProperties[$key];
            if ($opResourceProperties.count -gt 0) {
                #Deleting all existing resources
                $count = $routing.OperationResourceProperties.Count - 1
                for ($i = $count-1; $i -ge 0; $i--) {
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
                    $dummy = $routing.OperationResourceProperties.DelRowAtPos($i); 
                }

			        
                #Adding the new data
                foreach ($opResProp in $opResourceProperties) {
			      
                    $key = $opResProp.Sequence + '@#@' + $opResProp.RoutingCode
					
                    $routing.OperationResourceProperties.U_RtgOprCode = $drivers[$opResProp.Sequence]
                    $routing.OperationResourceProperties.U_RtgOprRscCode = $driversOprRsc[$key]
                    $routing.OperationResourceProperties.U_PrpCode = $opResProp.PropertiesCode
                    $routing.OperationResourceProperties.U_PrpConType = $opResProp.Condition
                    $routing.OperationResourceProperties.U_PrpConValue = $opResProp.Value
                    $routing.OperationResourceProperties.U_PrpConValueTo = $opResProp.ToValue
                    $routing.OperationResourceProperties.U_UnitOfMeasure = $opResProp.UoM
						
                    $dummy = $routing.OperationResourceProperties.Add()
			            
                }
            }
          
        }
    }
      
    $message = 0
    
    #Adding or updating Routings depends on exists in the database
    if ($retValue -eq 0) {
        try {
            [System.String]::Format("Updating Routing: {0}", $csvItem.RoutingCode)
     
            $message = $routing.Update()
        }
        catch [Exception] {
            Write-Host $_.Exception.InnerException.ToString()
        }
    }
    else {
        try {
            [System.String]::Format("Adding Routing: {0}", $csvItem.RoutingCode)
            $message = $routing.Add()
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
