Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany();
$pfcCompany.LicenseServer = "10.0.0.3:40000";
$pfcCompany.SQLServer = "10.0.0.2:30115";
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB";
$pfcCompany.Databasename = 'PFDEMOGB';
$pfcCompany.UserName = "manager";
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


#Data loading from a csv file
[array] $csvResources = Import-Csv -Delimiter ';' -Path  ($PSScriptRoot + "\Resources.csv")
[array] $csvResourcesProperties = Import-Csv -Delimiter ';' -Path  ($PSScriptRoot + "\Resources_Properties.csv")
[array] $csvResourcesAtachments = Import-Csv -Delimiter ';' -Path  ($PSScriptRoot + "\Resources_Attachments.csv")
 
#region preparing data
write-Host 'Preparing data: '
$totalRows = $csvResources.Count + $csvResourcesProperties.Count + $csvResourcesAtachments.Count;

$resourcesList = New-Object 'System.Collections.Generic.List[array]'
$resourcesPropertiesDict = New-Object 'System.Collections.Generic.Dictionary[String,System.Collections.Generic.List[array]]'
$resourcesAtachementsDict = New-Object 'System.Collections.Generic.Dictionary[String,System.Collections.Generic.List[array]]'
    
$progressItterator = 0;
$progress = 0;
$beforeProgress = 0;

if ($totalRows -gt 1) {
    $total = $totalRows
}
else {
    $total = 1
}

foreach ($row in $csvResources) {
    $progressItterator++;
    $progress = [math]::Round(($progressItterator * 100) / $total);
    if ($progress -gt $beforeProgress) {
        Write-Host $progress"% " -NoNewline
        $beforeProgress = $progress
    }

    $resourcesList.Add([array]$row);
}

foreach ($row in $csvResourcesProperties) {
    $key = $row.ResourceCode;
    $progressItterator++;
    $progress = [math]::Round(($progressItterator * 100) / $total);
    if ($progress -gt $beforeProgress) {
        Write-Host $progress"% " -NoNewline
        $beforeProgress = $progress
    }

    if ($resourcesPropertiesDict.ContainsKey($key)) {
        $list = $resourcesPropertiesDict[$key];
    }
    else {
        $list = New-Object System.Collections.Generic.List[array];
        $resourcesPropertiesDict[$key] = $list;
    }

    $list.Add([array]$row);
}
Write-Host '';

foreach ($row in $csvResourcesAtachments) {
    $key = $row.ResourceCode;
    $progressItterator++;
    $progress = [math]::Round(($progressItterator * 100) / $total);
    if ($progress -gt $beforeProgress) {
        Write-Host $progress"% " -NoNewline
        $beforeProgress = $progress
    }

    if ($resourcesAtachementsDict.ContainsKey($key)) {
        $list = $resourcesAtachementsDict[$key];
    }
    else {
        $list = New-Object System.Collections.Generic.List[array];
        $resourcesAtachementsDict[$key] = $list;
    }

    $list.Add([array]$row);
}
Write-Host '';
#endregion

$progressItterator = 0;
$progress = 0;
$beforeProgress = 0;
$totalRows = $downtimeList.Count;
if ($totalRows -gt 1) {
    $total = $totalRows
}
else {
    $total = 1
}
foreach ($csvItem in $resourcesList) {
    $progressItterator++;
    $progress = [math]::Round(($progressItterator * 100) / $total);
    if ($progress -gt $beforeProgress) {
        Write-Host $progress"% " -NoNewline
        $beforeProgress = $progress
    }
    try {
        #Creating Resource object
        $res = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"Resource")
        #Checking that the resource already exist
        $retVal = $res.GetByRscCode($csvItem.ResourceCode)
        if ($retValue -ne 0) {
            #Adding the new data
            $res.U_RscType = $csvItem.ResourceType #enum type; Machine = 1 or M, Labour = 2 or L, Tool = 3 or T, Subcontractor = 4 or S 
            $res.U_RscCode = $csvItem.ResourceCode
        }
        $res.U_RscName = $csvItem.ResourceName
        $res.U_RscGrpCode = $csvItem.ResourceGroup
        $res.U_QueueTime = $csvItem.QueTime
        if ($csvItem.QueTimeUoM -ne '') {
            $res.U_QueueRate = $csvItem.QueTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9 
        }
        $res.U_SetupTime = $csvItem.SetupTime
        if ($csvItem.SetupTimeUoM -ne '') {
            $res.U_SetupRate = $csvItem.SetupTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9 
        }
        $res.U_RunTime = $csvItem.RunTime
        if ($csvItem.RunTimeUoM -ne '') {
            $res.U_RunRate = $csvItem.RunTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9 
        }
        $res.U_StockTime = $csvItem.StockTime
        if ($csvItem.StockTimeUoM -ne '') {
            $res.U_StockRate = $csvItem.StockTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9     
        }
        if ($csvItem.ResourceNumber -ne '') {
            $res.U_ResourceCount = $csvItem.ResourceNumber
        }
        if ($csvItem.HasCycle -eq 1) {
            $res.U_HasCycles = $csvItem.HasCycle #enum type; 1 = Yes, 2 = No
            $res.U_CycleCap = $csvItem.CycleCapacity
        }
        
        $res.U_ResActCode = $csvItem.ResourceAccountingCode
        $res.U_Project = $csvItem.Project
        $res.U_OcrCode = $csvItem.Dimension1
        $res.U_OcrCode2 = $csvItem.Dimension2
        $res.U_OcrCode3 = $csvItem.Dimension3
        $res.U_OcrCode4 = $csvItem.Dimension4
        $res.U_OcrCode5 = $csvItem.Dimension5
        $res.U_WhsCode = $csvItem.IssueWhsCode
        $res.U_BinAbs = $csvItem.IssueBinAbs
        $res.U_RWhsCode = $csvItem.ReceiptWhsCode
        $res.U_RBinAbs = $csvItem.ReceiptBinAbs

    
        #$res.UDFItems.Item("U_UDF1").Value = $csvItem.UDF1 ## how to import UDFs
	
        if ($res.U_RscType -eq 'Subcontractor') {
            $res.U_VendorCode = $csvItem.VendorCode
            $res.U_ItemCode = $csvItem.ItemCode
        }
   
    
        #Data loading from a csv file - Resource Properties
        [array]$resProps = $resourcesPropertiesDict[$csvItem.ResourceCode]
        if ($resProps.count -gt 0) {
            #Deleting all existing properties
            $count = $res.Properties.Count
            for ($i = 0; $i -lt $count; $i++) {
                $dummy = $res.Properties.DelRowAtPos(0);
            }
        
            #Adding the new data       
            foreach ($prop in $resProps) {
                $res.Properties.U_PrpCode = $prop.PropertiesCode
                $res.Properties.U_PrpConType = $prop.Condition #enum ConditionType; Equal EQ = 1, NotEqual NE = 2, GratherThan GT = 3, GratherThanOrEqual GE = 4, LessThan LT = 5, LessThanOrEqual LE = 6, Between BT = 7
                $res.Properties.U_PrpConValue = $prop.Value
                $res.Properties.U_PrpConValueTo = $prop.ToValue
                $res.Properties.U_UnitOfMeasure = $prop.UoM
                $dummy = $res.Properties.Add()
            }
        
        
        
        }
    
        #Adding attachments to Resources
        [array]$resAttachments = $resourcesAtachementsDict[$csvItem.ResourceCode];
        if ($resAttachments.count -gt 0) {
            #Deleting all existing attachments
            $count = $res.Attachments.Count
            for ($i = 0; $i -lt $count; $i++) {
                $dummy = $res.Attachments.DelRowAtPos(0);
            }
        
            #Adding the new data
            foreach ($att in $resAttachments) {
                $fileName = [System.IO.Path]::GetFileName($att.AttachmentPath)
                $res.Attachments.U_FileName = $fileName
                $res.Attachments.U_AttDate = [System.DateTime]::Today
                $res.Attachments.U_Path = $att.AttachmentPath
                $dummy = $res.Attachments.Add()
            }
        }
 
        $message = 0
    
        #Adding or updating Resources depends on exists in the database
        if ($retVal -eq 0) {
            [System.String]::Format("Updating Resource: {0}", $csvItem.ResourceCode)
            $message = $res.Update()
        }
        else {
            [System.String]::Format("Adding Resource: {0}", $csvItem.ResourceCode)
            $message = $res.Add()
        }
    
        if ($message -lt 0) {    
            $err = $pfcCompany.GetLastErrorDescription()
            write-host -backgroundcolor red -foregroundcolor white $err
        } 
        else {
            write-host "Success"
        }   
    } 
    Catch {
        $err = $_.Exception.Message;
        $content = [string]::Format("Error occured for ResourceCode {0}: {1}", $csvItem.ResourceCode, $err);
        Write-Host -BackgroundColor DarkRed -ForegroundColor White $content;
        continue;
    }
}

