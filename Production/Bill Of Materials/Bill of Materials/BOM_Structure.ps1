#region #Script info
########################################################################
# CompuTec PowerShell Script - Import Bill of Materials Structures
########################################################################
$SCRIPT_VERSION = "2.1"
# Last tested PF version: PF 9.1 PL13
# Description:
#      Import Bill of Materials Structures. Script add new BOMs or will update existing BOMs.    
#      You need to have all requred files for import. The BOM_Coproducts.csv & BOM_Scraps.csv can be empty except first header line)
#      Sctipt check that Revision for Item Details exists.
#      By default all files needs be stored in catalog C:\PS\PF\BOM\ -Check section Script parameters and update catalog where files .ps1 and csv was saved
# Warning:
#   Make sure that item & item details was imported before use this script.
#   It's recommended run script when all users all disconnected.
#   Before running this script please Make Backup of your database.
# Troubleshooting:
#   https://connect.computec.pl/display/PF920EN/FAQ+PowerShell
#   https://connect.computec.pl/display/PF910EN/FAQ+PowerShell
########################################################################
#endregion

#region #PF API library usage
Clear-Host
Write-Host -backgroundcolor Yellow -foregroundcolor DarkBlue ("Script Version:" + $SCRIPT_VERSION)
# You need to check in what architecture PowerShell ISE is running (x64 or x86),
# you need run PowerShell ISE in the same architecture like PF API is installed (check in Windows -> Programs & Features)
# Examples:
#     SAP Client x64 + PF x64 installed on DB/Company => PF API x64 => Windows PowerShell ISE
#     SAP Client x86 + PF x86 installed on DB/Company => PF API x86 => Windows PowerShell ISE x86

[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#endregion

#region Script parameters

$csvImportCatalog = $PSScriptRoot + "\"

$csvBomFilePath = -join ($csvImportCatalog, "BOMs.csv")
$csvBomItemsFilePath = -join ($csvImportCatalog, "BOM_Items.csv")
$csvBomscrapsFilePath = -join ($csvImportCatalog, "BOM_Scraps.csv")
$csvBomCoproductsFilePath = -join ($csvImportCatalog, "BOM_Coproducts.csv")

#endregion

#region #Datbase/Company connection settings
 
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.LicenseServer = "10.0.0.xx:40000"
$pfcCompany.SQLServer = "10.0.0.xx:30015"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"
$pfcCompany.Databasename = "MICHALB_PFDEMOGB"
$pfcCompany.UserName = "michalb"
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

$csvItems = Import-Csv -Delimiter ';' -Path $csvBomFilePath
[array]$bomItems = Import-Csv -Delimiter ';' -Path $csvBomItemsFilePath 
[array]$bomScraps = $null;
if ((Test-Path -Path $csvBomscrapsFilePath -PathType leaf) -eq $true) {
    [array]$bomScraps = Import-Csv -Delimiter ';' -Path $csvBomscrapsFilePath
}
[array]$bomCoproducts = $null;
if ((Test-Path -Path $csvBomCoproductsFilePath -PathType leaf) -eq $true) {
    [array]$bomCoproducts = Import-Csv -Delimiter ';' -Path $csvBomCoproductsFilePath 
}
write-Host 'Preparing data: '
$totalRows = $csvItems.Count + $bomItems.Count + $bomScraps.Count + $bomCoproducts.Count

$bomList = New-Object 'System.Collections.Generic.List[array]'

$dictionaryItems = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
$dictionaryScraps = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
$dictionaryCoproducts = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

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
    $key = $row.BOM_ItemCode + '___' + $row.Revision;
    $progressItterator++;
    $progres = [math]::Round(($progressItterator * 100) / $total);
    if ($progres -gt $beforeProgress) {
        Write-Host $progres"% " -NoNewline
        $beforeProgress = $progres
    }

   
    $bomList.Add([array]$row);
}

foreach ($row in $bomItems) {
    $key = $row.BOM_ItemCode + '___' + $row.Revision;
    $progressItterator++;
    $progres = [math]::Round(($progressItterator * 100) / $total);
    if ($progres -gt $beforeProgress) {
        Write-Host $progres"% " -NoNewline
        $beforeProgress = $progres
    }

    if ($dictionaryItems.ContainsKey($key)) {
        $list = $dictionaryItems[$key];
    }
    else {
        $list = New-Object System.Collections.Generic.List[array];
        $dictionaryItems[$key] = $list;
    }
    
    $list.Add([array]$row);
}

foreach ($row in $bomScraps) {
    $key = $row.BOM_ItemCode + '___' + $row.Revision;
    $progressItterator++;
    $progres = [math]::Round(($progressItterator * 100) / $total);
    if ($progres -gt $beforeProgress) {
        Write-Host $progres"% " -NoNewline
        $beforeProgress = $progres
    }

    if ($dictionaryScraps.ContainsKey($key)) {
        $list = $dictionaryScraps[$key];
    }
    else {
        $list = New-Object System.Collections.Generic.List[array];
        $dictionaryScraps[$key] = $list;
    }
    
    $list.Add([array]$row);
}


foreach ($row in $bomCoproducts) {
    $key = $row.BOM_ItemCode + '___' + $row.Revision;
    $progressItterator++;
    $progres = [math]::Round(($progressItterator * 100) / $total);
    if ($progres -gt $beforeProgress) {
        Write-Host $progres"% " -NoNewline
        $beforeProgress = $progres
    }

    if ($dictionaryCoproducts.ContainsKey($key)) {
        $list = $dictionaryCoproducts[$key];
    }
    else {
        $list = New-Object System.Collections.Generic.List[array];
        $dictionaryCoproducts[$key] = $list;
    }
    
    $list.Add([array]$row);
}
Write-Host '';



foreach ($csvItem in $bomList) {
    $dictionaryKey = $csvItem.BOM_ItemCode + '___' + $csvItem.Revision;
    write-host "Importing BOM for item: "$csvItem.BOM_ItemCode ",Revision:" $csvItem.Revision
    #Check that Item & Item Details & Revision exist
    $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
    $rs.DoQuery([string]::Format("SELECT ""T0"".""ItemCode"", ""T2"".""U_Code"" FROM  ""OITM"" AS ""T0""
            INNER JOIN ""@CT_PF_OIDT"" AS ""T1"" ON ""T0"".""ItemCode"" = ""T1"".""U_ItemCode""
            INNER JOIN ""@CT_PF_IDT1"" AS ""T2"" ON ""T2"".""Code"" = ""T1"".""Code""
            WHERE
            ""T1"".""U_ItemCode"" = '{0}'
            and ""T2"".""U_Code"" = '{1}'", $csvItem.BOM_ItemCode, $csvItem.Revision))
    
    if ($rs.RecordCount -eq 0) {
        write-host "   Item:" $csvItem.BOM_ItemCode "Revision:" $csvItem.Revision "Can't be found." -backgroundcolor red -foregroundcolor white $_.Exception.Message
        continue;
    }

    #Creating BOM object
    $bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"BillOfMaterial")
    
    #Checking that the BOM already exist
    $retValue = $bom.GetByItemCodeAndRevision($csvItem.BOM_ItemCode, $csvItem.Revision)
    
    if ($retValue -ne 0) {
        $bom.U_ItemCode = $csvItem.BOM_ItemCode
        $bom.U_Revision = $csvItem.Revision
    }
    
    $bom.U_Quantity = $csvItem.Quantity
    $bom.U_Factor = $csvItem.Factor
    $bom.U_WhsCode = $csvItem.Warehouse
    $bom.U_OcrCode = $csvItem.DistRule
    $bom.U_OcrCode2 = $csvItem.DistRule2
    $bom.U_OcrCode3 = $csvItem.DistRule3
    $bom.U_OcrCode4 = $csvItem.DistRule4
    $bom.U_OcrCode5 = $csvItem.DistRule5
    $bom.U_Project = $csvItem.Projec
    $bom.U_ProdType = $csvItem.ProdType # I = Internal, E = External
    #$bom.UDFItems.Item("U_UDF1").Value = $csvItem.UDF1 # how to import UDF
        
    #Data loading from a csv file - BOM Items
   
    #[array]$bomItems = Import-Csv -Delimiter ';' -Path $csvBomItemsFilePath | Where-Object {$_.BOM_ItemCode -eq $csvItem.BOM_ItemCode -and $_.Revision -eq $csvItem.Revision}

    $bomItems = $dictionaryItems[$dictionaryKey]

    write-host "   Trying add: " $bomItems.count " items"

    if ($bomItems.count -gt 0) {
        #Deleting all existing items
        $count = $bom.Items.Count
        for ($i = 0; $i -lt $count; $i++) {
            $dummy = $bom.Items.DelRowAtPos(0);
        }
         
        #Adding the new data       
        foreach ($item in $bomItems) {
            $bom.Items.U_Sequence = $item.Sequence
            $bom.Items.U_ItemCode = $item.ItemCode
            $bom.Items.U_Revision = $item.Item_Revision
            $bom.Items.U_WhsCode = $item.Warehouse
            $bom.Items.U_Factor = $item.Factor
            $bom.Items.U_FactorDescription = $item.FactorDesc
            $bom.Items.U_Quantity = $item.Quantity
            $bom.Items.U_ScrapPercentage = $item.ScrapPercent
            $bom.Items.U_IssueType = $item.IssueType # M = Manual, B = Backflush
            $bom.Items.U_OcrCode = $item.OcrCode
            $bom.Items.U_OcrCode2 = $item.OcrCode2
            $bom.Items.U_OcrCode3 = $item.OcrCode3
            $bom.Items.U_OcrCode4 = $item.OcrCode4
            $bom.Items.U_OcrCode5 = $item.OcrCode5
            $bom.Items.U_Project = $item.Project

            if ($item.SubcontractingItem -eq 'Y') {
                $bom.Items.U_SubRecepitItem = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
            }
            else {
                $bom.Items.U_SubRecepitItem = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
            }
            $bom.Items.U_Remarks = $item.Remarks
            if ($item.Formula -ne "") {
                $bom.Items.U_Formula = $item.Formula
            }
             
            $dummy = $bom.Items.Add()
        }
    }
    
    #Data loading from a csv file - BOM Coproducts
    [array]$bomCoproducts = @();
    $bomCoproducts = $dictionaryCoproducts[$dictionaryKey];
    if ($bomCoproducts.Count -gt 0) {
        write-host "   Trying add: " $bomCoproducts.count " Coproducts"   
    }
    else {
        write-host "   Warning! Item:" $csvItem.BOM_ItemCode "Revision:" $csvItem.Revision "Can't find Coproducts." -backgroundcolor yellow -foregroundcolor black
    }

    if ($bomCoproducts.Count -gt 0) {
        #Deleting all existing items
        $count = $bom.Coproducts.Count
        for ($i = 0; $i -lt $count; $i++) {
            $dummy = $bom.Coproducts.DelRowAtPos(0);
        }
         
        #Adding the new data       
        foreach ($coproducts in $bomCoproducts) {
            $bom.Coproducts.U_Sequence = $coproducts.Sequence
            $bom.Coproducts.U_ItemCode = $coproducts.ItemCode
            $bom.Coproducts.U_Revision = $coproducts.Item_Revision
            $bom.Coproducts.U_WhsCode = $coproducts.Warehouse
            $bom.Coproducts.U_Factor = $coproducts.Factor
            $bom.Coproducts.U_FactorDescription = $coproducts.FactorDesc
            $bom.Coproducts.U_Quantity = $coproducts.Quantity
            $bom.Coproducts.U_IssueType = $coproducts.IssueType # M = Manual, B = Backflush
            $bom.Coproducts.U_OcrCode = $coproducts.OcrCode
            $bom.Coproducts.U_OcrCode2 = $coproducts.OcrCode2
            $bom.Coproducts.U_OcrCode3 = $coproducts.OcrCode3
            $bom.Coproducts.U_OcrCode4 = $coproducts.OcrCode4
            $bom.Coproducts.U_OcrCode5 = $coproducts.OcrCode5
            $bom.Coproducts.U_Project = $coproducts.Project
            $bom.Coproducts.U_Remarks = $coproducts.Remarks
            if ($coproducts.Formula -ne "") {
                $bom.Coproducts.U_Formula = $coproducts.Formula
            }
            $dummy = $bom.Coproducts.Add()
        }
    }
      
    #Data loading from a csv file - BOM Scraps
    [array]$bomScraps = @();
    $bomScraps = $dictionaryScraps[$dictionaryKey];
    if ($bomScraps.Count -gt 0) {
        [array]$bomScraps = Import-Csv -Delimiter ';' -Path $csvBomscrapsFilePath | Where-Object {$_.BOM_ItemCode -eq $csvItem.BOM_ItemCode -and $_.Revision -eq $csvItem.Revision}
        write-host "   Trying add: " $bomScraps.count " Scraps"          
    }
    else {
        write-host "   Warning! Item:" $csvItem.BOM_ItemCode "Revision:" $csvItem.Revision "Can't find Scraps." -backgroundcolor yellow -foregroundcolor black
    }

    if ($bomScraps.count -gt 0) {
        #Deleting all existing items
        $count = $bom.Scraps.Count
        for ($i = 0; $i -lt $count; $i++) {
            $dummy = $bom.Scraps.DelRowAtPos(0);
        }
         
        #Adding the new data       
        foreach ($scraps in $bomScraps) {
            $bom.Scraps.U_Sequence = $scraps.Sequence
            $bom.Scraps.U_ItemCode = $scraps.ItemCode
            $bom.Scraps.U_Revision = $scraps.Item_Revision
            $bom.Scraps.U_WhsCode = $scraps.Warehouse
            $bom.Scraps.U_Factor = $scraps.Factor
            $bom.Scraps.U_FactorDescription = $scraps.FactorDesc
            $bom.Scraps.U_Quantity = $scraps.Quantity
            $bom.Scraps.U_Type = $scraps.Type #enum type; Technological = 1, UseFul = 2
            $bom.Scraps.U_IssueType = $scraps.IssueType # M = Manual, B = Backflush
            $bom.Scraps.U_OcrCode = $scraps.OcrCode
            $bom.Scraps.U_OcrCode2 = $scraps.OcrCode2
            $bom.Scraps.U_OcrCode3 = $scraps.OcrCode3
            $bom.Scraps.U_OcrCode4 = $scraps.OcrCode4
            $bom.Scraps.U_OcrCode5 = $scraps.OcrCode5
            $bom.Scraps.U_Project = $scraps.Project
            $bom.Scraps.U_Remarks = $scraps.Remarks
            if ($scraps.Formula -ne "") {
                $bom.Scraps.U_Formula = $scraps.Formula
            }
            $dummy = $bom.Scraps.Add()
        }
    }
    $bom.U_BatchSize = $csvItem.BatchSize
    
    #Adding or updating BOMs depends on exists in the database
    $message = 0;

    try {
        if ($retValue -eq 0) {
            [System.String]::Format("   Updating BOM: {0} Revision: {1}", $csvItem.BOM_ItemCode, $csvItem.Revision)
            $message = $bom.Update()
        }
        else {
            [System.String]::Format("   Adding BOM: {0}  Revision: {1}", $csvItem.BOM_ItemCode, $csvItem.Revision)
            $message = $bom.Add()
        }

        if ($message -lt 0) {   
            $err = $pfcCompany.GetLastErrorDescription()
           
            write-host -backgroundcolor red "Failed -" -foregroundcolor white $err
        }
        else {
            write-host -backgroundcolor green "Success"-foregroundcolor black
        }  
    }
    catch {
        write-host "   Item:" $csvItem.BOM_ItemCode "Revision:" $csvItem.Revision "Import failed: " -backgroundcolor red -foregroundcolor white $_.Exception.Message $pfcCompany.GetLastErrorDescription()
        write-host -backgroundcolor red "Failed"-foregroundcolor white
    }
}

#region Close connection

if ($pfcCompany.IsConnected) {
    $pfcCompany.Disconnect()
    
    write-host " "
    write-host  –backgroundcolor green –foregroundcolor black "Disconnected from the company"
}

#endregion