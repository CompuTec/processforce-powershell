#region #Script info
########################################################################
# CompuTec PowerShell Script - Import Bill of Materials Structures
########################################################################
$SCRIPT_VERSION = "3.7"
# Last tested PF version: ProcessForce 10.0 (2.10.13.48), Release 13 PL: MAIN (64-bit)
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
#   https://connect.computec.pl/display/PF930EN/PowerShell+FAQ
#   https://connect.computec.pl/display/PF930EN/PowerShell+FAQ
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
$csvBomAttachmentsFilePath = -join ($csvImportCatalog, "BOM_Attachments.csv")

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

	[array]$csvItems = Import-Csv -Delimiter ';' -Path $csvBomFilePath
	[array]$bomItems = Import-Csv -Delimiter ';' -Path $csvBomItemsFilePath 
	[array]$bomScraps = $null;
	if ((Test-Path -Path $csvBomscrapsFilePath -PathType leaf) -eq $true) {
		[array]$bomScraps = Import-Csv -Delimiter ';' -Path $csvBomscrapsFilePath
	}
	else {
		write-host "BOM Scraps - csv not available."
	}
	[array]$bomCoproducts = $null;
	if ((Test-Path -Path $csvBomCoproductsFilePath -PathType leaf) -eq $true) {
		[array]$bomCoproducts = Import-Csv -Delimiter ';' -Path $csvBomCoproductsFilePath 
	}
	else {
		write-host "BOM Coproducts - csv not available."
	}
	[array]$bomAttachments = $null;
	if ((Test-Path -Path $csvBomAttachmentsFilePath -PathType leaf) -eq $true) {
		[array]$bomAttachments = Import-Csv -Delimiter ';' -Path $csvBomAttachmentsFilePath 
	}
	else {
		write-host "BOM Attachments - csv not available."
	}
	write-Host 'Preparing data: '
	$totalRows = $csvItems.Count + $bomItems.Count + $bomScraps.Count + $bomCoproducts.Count + $bomAttachments.Count;

	$bomList = New-Object 'System.Collections.Generic.List[array]'

	$dictionaryItems = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$dictionaryScraps = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$dictionaryCoproducts = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$dictionaryAttachments = New-Object 'System.Collections.Generic.Dictionary[String,System.Collections.Generic.List[array]]'

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

	foreach ($row in $bomAttachments) {
		$key = $row.BOM_ItemCode + '___' + $row.Revision;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}

		if ($dictionaryAttachments.ContainsKey($key)) {
			$list = $dictionaryAttachments[$key];
		}
		else {
			$list = New-Object System.Collections.Generic.List[array];
			$dictionaryAttachments[$key] = $list;
		}
    
		$list.Add([array]$row);
	}
	Write-Host '';

	Write-Host 'Adding/updating data: ';
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
			$dictionaryKey = $csvItem.BOM_ItemCode + '___' + $csvItem.Revision;
            
			#Check that Item & Item Details & Revision exist
			$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
			$rs.DoQuery([string]::Format("SELECT ""T0"".""ItemCode"", ""T2"".""U_Code"" FROM  ""OITM"" AS ""T0""
            INNER JOIN ""@CT_PF_OIDT"" AS ""T1"" ON ""T0"".""ItemCode"" = ""T1"".""U_ItemCode""
            INNER JOIN ""@CT_PF_IDT1"" AS ""T2"" ON ""T2"".""Code"" = ""T1"".""Code""
            WHERE
            ""T1"".""U_ItemCode"" = '{0}'
            and ""T2"".""U_Code"" = '{1}'", $csvItem.BOM_ItemCode, $csvItem.Revision))
    
			if ($rs.RecordCount -eq 0) {
				$err = [string]::Format("Item: {0} with Revision: {1} not found.", $csvItem.BOM_ItemCode, $csvItem.Revision);
				Throw [System.Exception] ($err)
			}

			#Creating BOM object
			$bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"BillOfMaterial")
    
			#Checking that the BOM already exist
			$retValue = $bom.GetByItemCodeAndRevision($csvItem.BOM_ItemCode, $csvItem.Revision)
    
			if ($retValue -ne 0) {
				$bom.U_ItemCode = $csvItem.BOM_ItemCode
				$bom.U_Revision = $csvItem.Revision
				$exists = $false;
			}
			else {
				$exists = $true;
			}
    
			$bom.U_Quantity = $csvItem.Quantity
			$bom.U_Factor = $csvItem.Factor
			if ([string]::IsNullOrWhiteSpace($csvItem.Yield) -eq $false) {
				$bom.U_Yield = $csvItem.Yield;
			}
			if ([string]::IsNullOrWhiteSpace($csvItem.YieldFormula) -eq $false) {
				$bom.U_YieldFormula = $csvItem.YieldFormula;
			}
			else {
				$bom.U_YieldFormula = [string]::Empty;
			}
			if ([string]::IsNullOrWhiteSpace($csvItem.YieldItemsFormula) -eq $false) {
				$bom.U_ItemFormula = $csvItem.YieldItemsFormula;
			}
			else {
				$bom.U_ItemFormula = [string]::Empty;
			}
			if ([string]::IsNullOrWhiteSpace($csvItem.YieldCoproductsFormula) -eq $false) {
				$bom.U_CoproductFormula = $csvItem.YieldCoproductsFormula;
			}
			else {
				$bom.U_CoproductFormula = [string]::Empty;
			}
			if ([string]::IsNullOrWhiteSpace($csvItem.U_CoproductFormula) -eq $false) {
				$bom.U_ScrapFormula = $csvItem.U_CoproductFormula;
			}
			else {
				$bom.U_ScrapFormula = [string]::Empty;
			}
			$bom.U_WhsCode = $csvItem.Warehouse
			$bom.U_OcrCode = $csvItem.DistRule
			$bom.U_OcrCode2 = $csvItem.DistRule2
			$bom.U_OcrCode3 = $csvItem.DistRule3
			$bom.U_OcrCode4 = $csvItem.DistRule4
			$bom.U_OcrCode5 = $csvItem.DistRule5
			$bom.U_Project = $csvItem.Project
			$bom.U_ProdType = $csvItem.ProdType # I = Internal, E = External
			if (-not [string]::IsNullOrEmpty($csvItem.Instructions)) {
				$bom.U_Instructions = [string] $csvItem.Instructions.Replace("``n", "`n");
			}
			else {
				$bom.U_Instructions = [string]::Empty
			}
			#$bom.UDFItems.Item("U_UDF1").Value = $csvItem.UDF1 # how to import UDF


			$bomItems = $dictionaryItems[$dictionaryKey]

			#Deleting all existing items
			$count = $bom.Items.Count
			for ($i = 0; $i -lt $count; $i++) {
				$dummy = $bom.Items.DelRowAtPos(0);
			}
			if ($bomItems.count -gt 0) {
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
					$bom.Items.U_IssueType = $item.IssueType # M = Manual, B = Backflush, #F = Fixed Backflush
					if ($bom.Items.U_IssueType -eq 'B') {
						if ([string]::IsNullOrWhiteSpace($item.BinCode) -eq $false) {
							$bom.Items.U_BinCode = $item.BinCode;
						}
					}
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
					 
					if ([string]::IsNullOrWhiteSpace($item.Yield) -eq $false) {
						$bom.Items.U_Yield = $item.Yield;
					}

					$bom.Items.U_Remarks = $item.Remarks
					if ([string]::IsNullOrWhiteSpace($item.Formula) -eq $false) {
						$bom.Items.U_Formula = $item.Formula
					}
             
					$dummy = $bom.Items.Add()
				}
			}
    
			#Data loading from a csv file - BOM Coproducts
			[array]$bomCoproducts = @();
			$bomCoproducts = $dictionaryCoproducts[$dictionaryKey];
			#Deleting all existing items
			$count = $bom.Coproducts.Count
			for ($i = 0; $i -lt $count; $i++) {
				$dummy = $bom.Coproducts.DelRowAtPos(0);
			}
			if ($bomCoproducts.Count -gt 0) {
				#Adding the new data       
				foreach ($coproducts in $bomCoproducts) {
					$bom.Coproducts.U_Sequence = $coproducts.Sequence
					$bom.Coproducts.U_ItemCode = $coproducts.ItemCode
					$bom.Coproducts.U_Revision = $coproducts.Item_Revision
					$bom.Coproducts.U_WhsCode = $coproducts.Warehouse
					$bom.Coproducts.U_Factor = $coproducts.Factor
					$bom.Coproducts.U_FactorDescription = $coproducts.FactorDesc
					$bom.Coproducts.U_Quantity = $coproducts.Quantity
					$bom.Coproducts.U_IssueType = $coproducts.IssueType # M = Manual, B = Backflush, #F = Fixed Backflush
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
					
					if ([string]::IsNullOrWhiteSpace($coproducts.Yield) -eq $false) {
						$bom.CoProducts.U_Yield = $coproducts.Yield;
					}
					$dummy = $bom.Coproducts.Add()
				}
			}
      
			#Data loading from a csv file - BOM Scraps
			[array]$bomScraps = @();
			$bomScraps = $dictionaryScraps[$dictionaryKey];
			#Deleting all existing items
			$count = $bom.Scraps.Count
			for ($i = 0; $i -lt $count; $i++) {
				$dummy = $bom.Scraps.DelRowAtPos(0);
			}
			if ($bomScraps.count -gt 0) {
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
					$bom.Scraps.U_IssueType = $scraps.IssueType # M = Manual, B = Backflush, #F = Fixed Backflush
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
					if ([string]::IsNullOrWhiteSpace($scraps.Yield) -eq $false) {
						$bom.Scraps.U_Yield = $scraps.Yield;
					}
					$dummy = $bom.Scraps.Add()
				}
			}
			$bom.U_BatchSize = $csvItem.BatchSize
    
			#Adding attachments to Resources
			[array]$bomAttachmentsData = $dictionaryAttachments[$dictionaryKey];
			if ($bomAttachmentsData.count -gt 0) {
				#Deleting all existing attachments
				$count = $bom.Attachments.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $bom.Attachments.DelRowAtPos(0);
				}
				$bom.Attachments.SetCurrentLine(0);
				#Adding the new data
				foreach ($att in $bomAttachmentsData) {
					if ($bom.Attachments.IsRowFilled()) {
						$dummy = $bom.Attachments.Add()
					}
					# $fileName = [System.IO.Path]::GetFileName($att.AttachmentPath)
					# $bom.Attachments.U_AttFileName = $fileName
					# $bom.Attachments.U_AttDate = [System.DateTime]::Today
					# $bom.Attachments.U_AttPath = $att.AttachmentPath
					$bom.Attachments.U_FullPath = $att.AttachmentPath
				}
			}


			#Adding or updating BOMs depends if it already exists in the database
			$message = 0;
    
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
			if ($exists -eq $false) {
				$taskMsg = "adding";
			}
			else {
				$taskMsg = "updating"
			}
			$ms = [string]::Format("Error when {0} Bill Of Material with ItemCode {1} and Revision {2} Details: {3}", $taskMsg, $csvItem.BOM_ItemCode, $csvItem.Revision, $err);
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