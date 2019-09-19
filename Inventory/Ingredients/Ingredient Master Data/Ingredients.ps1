#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Ingredients
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Ingredients. Script add new or will update existing Ingredients.
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

$csvIngredientsPath = -join ($csvImportCatalog, "Ingredients.csv")
$csvIngredientAllergensPath = -join ($csvImportCatalog, "IngredientAllergens.csv")
$csvIngredientClassificationsPath = -join ($csvImportCatalog, "IngredientClassifications.csv")
$csvIngredientClassificationCertificatesPath = -join ($csvImportCatalog, "IngredientClassificationCertificates.csv")
$csvIngredientSpecificationsPath = -join ($csvImportCatalog, "IngredientSpecifications.csv")

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
    $csvIngredientsPath = -join ($csvImportCatalog, "Ingredients.csv")
    $csvIngredientAllergensPath = -join ($csvImportCatalog, "IngredientAllergens.csv")
    $csvIngredientClassificationsPath = -join ($csvImportCatalog, "IngredientClassifications.csv")
    $csvIngredientClassificationCertificatesPath = -join ($csvImportCatalog, "IngredientClassificationCertificates.csv")
    $csvIngredientSpecificationsPath = -join ($csvImportCatalog, "IngredientSpecifications.csv")

    [array] $csvIngredients = Import-Csv -Delimiter ';' $csvIngredientsPath;
	
    if ((Test-Path -Path $csvIngredientAllergensPath -PathType leaf) -eq $true) {
        [array] $csvIngredientAllergens = Import-Csv -Delimiter ';' $csvIngredientAllergensPath;
    }
    else {
        [array] $csvIngredientAllergens = $null;
        write-host "Item Properties References - csv not available."
    }
    if ((Test-Path -Path $csvIngredientClassificationsPath -PathType leaf) -eq $true) {
        [array] $csvIngredientClassifications = Import-Csv -Delimiter ';' $csvIngredientClassificationsPath;
    }
    else {
        [array] $csvIngredientClassifications = $null;
        write-host "Item Properties References - csv not available."
    }
    if ((Test-Path -Path $csvIngredientClassificationCertificatesPath -PathType leaf) -eq $true) {
        [array] $csvIngredientClassificationCertificates = Import-Csv -Delimiter ';' $csvIngredientClassificationCertificatesPath;
    }
    else {
        [array] $csvIngredientClassificationCertificates = $null;
        write-host "Item Properties References - csv not available."
    }
    if ((Test-Path -Path $csvIngredientSpecificationsPath -PathType leaf) -eq $true) {
        [array] $csvIngredientSpecifications = Import-Csv -Delimiter ';' $csvIngredientSpecificationsPath;
    }
    else {
        [array] $csvIngredientSpecifications = $null;
        write-host "Item Properties References - csv not available."
    }

    write-Host 'Preparing data: ' -NoNewline
    $totalRows = $csvIngredients.Count + $csvIngredientAllergens.Count + $csvIngredientClassifications.Count + $csvIngredientClassificationCertificates.Count + $csvIngredientSpecifications.Count;
    
    $ingredientsList = New-Object 'System.Collections.Generic.List[array]'
    $allergensDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $classificationsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $classificationsCertificatesDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $specificationsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
	
    if ($totalRows -gt 1) {
        $total = $totalRows
    }
    else {
        $total = 1
    }

    foreach ($row in $csvIngredients) {
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        $ingredientsList.Add([array]$row);
    }

    foreach ($row in $csvIngredientAllergens) {
        $key = $row.IngredientCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
	
        if ($allergensDict.ContainsKey($key) -eq $false) {
            $allergensDict[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $allergensDict[$key];
		
        $list.Add([array]$row);
    }
    foreach ($row in $csvIngredientClassifications) {
        $key = $row.IngredientCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
	
        if ($classificationsDict.ContainsKey($key) -eq $false) {
            $classificationsDict[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $classificationsDict[$key];
		
        $list.Add([array]$row);
    }
    foreach ($row in $csvIngredientClassificationCertificates) {
        $key = $row.IngredientCode + '___' + $row.ClassificationCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
	
        if ($classificationsCertificatesDict.ContainsKey($key) -eq $false) {
            $classificationsCertificatesDict[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $classificationsCertificatesDict[$key];
		
        $list.Add([array]$row);
    }
    foreach ($row in $csvIngredientSpecifications) {
        $key = $row.IngredientCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
	
        if ($specificationsDict.ContainsKey($key) -eq $false) {
            $specificationsDict[$key] = New-Object System.Collections.Generic.List[array];
        }
        $list = $specificationsDict[$key];
		
        $list.Add([array]$row);
    }

    Write-Host '';
    Write-Host 'Adding/updating data: ' -NoNewline;
    if ($ingredientsList.Count -gt 1) {
        $total = $ingredientsList.Count
    }
    else {
        $total = 1
    }
    $progressItterator = 0;
    $progress = 0;
    $beforeProgress = 0;

    $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")

    foreach ($csvHeader in $ingredientsList) {
        try {
            $progressItterator++;
            $progress = [math]::Round(($progressItterator * 100) / $total);
            if ($progress -gt $beforeProgress) {
                Write-Host $progress"% " -NoNewline
                $beforeProgress = $progress
            }
            $rs.DoQuery([string]::Format("SELECT ""Code"" FROM ""@CT_PF_OIMD"" WHERE ""U_Code"" = N'{0}'", $csvHeader.Code));
	
            #Creating object
            $md = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::IngredientMasterData)
            #Checking if data already exists
            if ($rs.RecordCount -gt 0) {
                $dummy = $md.GetByKey($rs.Fields.Item(0).Value);
                $exists = $true
            }
            else {
                $md.U_Code = $csvHeader.Code;
                $exists = $false
            }
   
            $md.U_Desc = $csvHeader.Description;
            $md.U_UoM = $csvHeader.UoM;
            $md.U_Category = $csvHeader.Category;
            $md.U_AltCode = $csvHeader.AlternativeCode;
	
            $md.U_Remarks = $csvHeader.Remarks;
	

            #Data loading from a csv file 
            [array]$csvAllergens = $allergensDict[$csvHeader.Code];    
            if ($csvAllergens.count -gt 0) {
                #Deleting all existing items
                $count = $md.Allergens.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $md.Allergens.DelRowAtPos(0);
                }
         
                #Adding the new data       
                foreach ($csvAllergen in $csvAllergens) {
                    $md.Allergens.U_AlgCode = $csvAllergen.Code;
                    $dummy = $md.Allergens.Add();
                }
            }

            #Data loading from a csv file 
            [array]$csvClassifications = $classificationsDict[$csvHeader.Code];
            if ($csvClassifications.count -gt 0) {
                #Deleting all existing items
                $count = $md.Classifications.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $md.Classifications.DelRowAtPos(0);
                }
         
                $count = $md.Certificates.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $md.Certificates.DelRowAtPos(0);
                }
        
                #Adding the new data       
                foreach ($csvClassification in $csvClassifications) {
                    $md.Classifications.U_ClassCode = $csvClassification.Code;
                    $dummy = $md.Classifications.Add();

                    $certKey = $csvHeader.Code + '___' + $csvClassification.Code;
                    [array]$csvCertificates = $classificationsCertificatesDict[$certKey];
                    if ($csvCertificates.Count -gt 0) {
                        #Adding the new data       
                        foreach ($csvCertificate in $csvCertificates) {
                            $md.Certificates.U_ClassCode = $csvClassification.Code;
                            $md.Certificates.U_BPCode = $csvCertificate.BusinessPartnerCode;
                            $md.Certificates.U_CertNum = $csvCertificate.CertificateNumber;
                            $md.Certificates.U_CertDate = $csvCertificate.CertificateDate;
                            $md.Certificates.U_Status = $csvCertificate.Status; #NA - not approved, P - pending, A - approved
                            $md.Certificates.U_StatDate = $csvCertificate.StatusDate;
                            $md.Certificates.U_Attachment = $csvCertificate.Attachment;
                            $md.Certificates.U_Remarks = $csvCertificate.Remarks;
                            $dummy = $md.Certificates.Add();
                        }
                    }

                }


            }

            [array]$csvSpecifications = $specificationsDict[$csvHeader.Code];
            if ($csvSpecifications.count -gt 0) {
                #Deleting all existing items
                $count = $md.Specifications.Count
                for ($i = 0; $i -lt $count; $i++) {
                    $dummy = $md.Specifications.DelRowAtPos(0);
                }

                #Adding the new data       
                foreach ($csvSpecification in $csvSpecifications) {
                    $md.Specifications.U_BPCode = $csvSpecification.BusinessPartnerCode;
                    $md.Specifications.U_SpecNum = $csvSpecification.SpecificationNumber
                    $md.Specifications.U_SpecDate = $csvSpecification.SpecificationDate;
                    $md.Specifications.U_Status = $csvSpecification.Status; #NA - not approved, P - pending, A - approved
                    $md.Specifications.U_StatDate = $csvSpecification.StatusDate;
                    $md.Specifications.U_Attachment = $csvSpecification.Attachment;
                    $md.Specifications.U_Remarks = $csvSpecification.Remarks;
                    $dummy = $md.Specifications.Add();
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
            if ($exists -eq $false) {
                $taskMsg = "adding";
            }
            else {
                $taskMsg = "updating"
            }
            $ms = [string]::Format("Error when {0} Igredient with Code {1} Details: {2}", $taskMsg, $csvHeader.Code, $err);
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

