#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Ingredient Templates
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Ingredient Templates. Script add new or will update existing Ingredient Templates.
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

$csvIngredientTemplatesPath = -join ($csvImportCatalog, "IngredientTemplates.csv")
$csvIngredientTemplateIngredientsPath = -join ($csvImportCatalog, "IngredientTemplateIngredients.csv")
$csvIngredientTemplateNutrientsPath = -join ($csvImportCatalog, "IngredientTemplateNutrients.csv")

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
    [array] $csvIngredientTemplates = Import-Csv -Delimiter ';' $csvIngredientTemplatesPath;
    if ((Test-Path -Path $csvIngredientTemplateIngredientsPath -PathType leaf) -eq $true) {
        [array] $csvIngredientTemplateIngredients = Import-Csv -Delimiter ';' $csvIngredientTemplateIngredientsPath;
    }
    else {
        [array] $csvIngredientTemplateIngredients = $null;
        write-host "Ingredient Template Ingredients - csv not available."
    }
    if ((Test-Path -Path $csvIngredientTemplateNutrientsPath -PathType leaf) -eq $true) {
        [array] $csvIngredientTemplateNutrients = Import-Csv -Delimiter ';' $csvIngredientTemplateNutrientsPath;
    }
    else {
        [array] $csvIngredientTemplateNutrients = $null;
        write-host "Ingredient Template Nutrients - csv not available."
    }

    write-Host 'Preparing data: ' -NoNewline
    $totalRows = $csvIngredientTemplates.Count + $csvIngredientTemplateIngredients.Count + $csvIngredientTemplateNutrients.Count;
    $ingredientTemplatesList = New-Object 'System.Collections.Generic.List[array]'
    $ingredientTemplateIngredientsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $ingredientTemplateNutrientsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
	
    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $csvIngredientTemplates) {
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        $ingredientTemplatesList.Add([array]$row);
    }

    foreach ($row in $csvIngredientTemplateIngredients) {
        $key = $row.TemplateCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
	
        if ($ingredientTemplateIngredientsDict.ContainsKey($key) -eq $false) {
            $ingredientTemplateIngredientsDict[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $ingredientTemplateIngredientsDict[$key];
		
        $list.Add([array]$row);
    }
    foreach ($row in $csvIngredientTemplateNutrients) {
        $key = $row.TemplateCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
	
        if ($ingredientTemplateNutrientsDict.ContainsKey($key) -eq $false) {
            $ingredientTemplateNutrientsDict[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $ingredientTemplateNutrientsDict[$key];
		
        $list.Add([array]$row);
    }

    Write-Host '';
    Write-Host 'Adding/updating data: ' -NoNewline;
   

    $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
    if ($ingredientTemplatesList.Count -gt 1) {
        $total = $ingredientTemplatesList.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;
    foreach ($csvHeader in $ingredientTemplatesList) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OIGT"" WHERE ""U_Code"" = N'{0}'", $csvHeader.Code));
	
            #Creating object
            $md = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::IngredientTemplate)
            #Checking if data already exists

            if ($rs.RecordCount -gt 0) {
                $dummy = $md.GetByKey($rs.Fields.Item(0).Value);
                $exists = $true;
            }
            else {
                $md.U_Code = $csvHeader.Code;
                $exists = $false;
            }
   
            $md.U_Name = $csvHeader.Name;
            $md.U_Remarks = $csvHeader.Remarks;
	
            [array]$csvIngredients = $ingredientTemplateIngredientsDict[$csvHeader.Code];
    
            if ($csvIngredients.count -gt 0) {
                #Deleting all existing items
                $count = $md.Ingredients.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $md.Ingredients.DelRowAtPos(0);
                }
         
                #Adding the new data       
                foreach ($csvIngredient in $csvIngredients) {
                    $md.Ingredients.U_IgdCode = $csvIngredient.Code;
                    $md.Ingredients.U_Value = $csvIngredient.Value
                    $dummy = $md.Ingredients.Add();
                }
            }

            [array]$csvNutrients = $ingredientTemplateNutrientsDict[$csvHeader.Code];
            if ($csvNutrients.count -gt 0) {
                #Deleting all existing items
                $count = $md.Nutrients.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $md.Nutrients.DelRowAtPos(0);
                }
         
                #Adding the new data       
                foreach ($csvNutrient in $csvNutrients) {
                    $md.Nutrients.U_NutCode = $csvNutrient.Code;
                    $md.Nutrients.U_Value = $csvNutrient.Value
                    $dummy = $md.Nutrients.Add();
                }
            }

            $message = 0
            #Adding or updating depends on exists in the database
            if ($exists -eq $true) {
                $message = $md.Update()
            }
            else {
                $message = $md.Add()
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
            $ms = [string]::Format("Error when {0} Ingredient Template with Code {1} Details: {2}", $taskMsg, $csvHeader.Code, $err);
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

