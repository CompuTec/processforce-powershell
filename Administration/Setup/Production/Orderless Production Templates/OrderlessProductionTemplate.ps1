#region #Script info
########################################################################
# CompuTec PowerShell Script - Import Orderless Production Templates
########################################################################
$SCRIPT_VERSION = "3.0"
# Last tested PF version: ProcessForce 9.3 (9.30.200) PL: MAIN (64-bit)
# Description:
#      Import Orderless Production Templates. Script add new Templates or will update existing Templates.    
#      You need to have all requred files for import.)
#      By default all files needs be stored in same catalog as script.
# Warning:
#   It's recommended run script when all users all disconnected.
#   Before running this script please Make Backup of your database.
# Troubleshooting:
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

$csvTemplatesPath = -join ($csvImportCatalog, "OrderlessProductionTemplate.csv")
$csvTemplatesLinesPath = -join ($csvImportCatalog, "OrderlessProductionTemplateLines.csv")
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

#region Connect to company
 
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

#region additionalInfoFunctionsy
$qmBOMCodeInfo = New-Object 'CompuTec.Core.DI.Database.QueryManager';
$qmBOMCodeInfo.CommandText = "SELECT B.""Code"", B.""U_ItemCode"", B.""U_Revision"", 0 AS ""U_LineNum"" FROM ""@CT_PF_OBOM"" B WHERE B.""U_ItemCode"" = @BOMItemCode AND B.""U_Revision"" = @BOMRevision";
$qmBOMCoproductInfo = New-Object 'CompuTec.Core.DI.Database.QueryManager';
$qmBOMCoproductInfo.CommandText = "SELECT BC.""Code"", BC.""U_ItemCode"", BC.""U_Revision"", BC.""U_LineNum"" FROM ""@CT_PF_OBOM"" B INNER JOIN ""@CT_PF_BOM3"" BC ON B.""Code"" = BC.""Code"" WHERE B.""U_ItemCode"" = @BOMItemCode AND B.""U_Revision"" = @BOMRevision AND BC.""U_ItemCode"" = @ItemCode AND BC.""U_Revision"" = @ItemRevision AND BC.""U_Sequence"" = @Sequence";
$qmBOMScrapInfo = New-Object 'CompuTec.Core.DI.Database.QueryManager';
$qmBOMScrapInfo.CommandText = "SELECT BS.""Code"", BS.""U_ItemCode"", BS.""U_Revision"", BS.""U_LineNum"" FROM ""@CT_PF_OBOM"" B INNER JOIN ""@CT_PF_BOM4"" BS ON B.""Code"" = BS.""Code"" WHERE B.""U_ItemCode"" = @BOMItemCode AND B.""U_Revision"" = @BOMRevision AND BS.""U_ItemCode"" = @ItemCode AND BS.""U_Revision"" = @ItemRevision AND BS.""U_Sequence"" = @Sequence";	

function getBOMInfo($token, $BOMItemCode, $BOMRevision) {
	try {
		$qmBOMCodeInfo.ClearParameters();
		$qmBOMCodeInfo.AddParameter("BOMItemCode", $BOMItemCode);
		$qmBOMCodeInfo.AddParameter("BOMRevision", $BOMRevision);
		$result = $qmBOMCodeInfo.Execute($token); 

		if ($result.RecordCount -ne 1) {
			throw [System.Exception]("Bill Of Materials don't exists");
		}

		return [psobject]@{
			Code     = $result.Fields.Item('Code').Value
			ItemCode = $result.Fields.Item('U_ItemCode').Value
			Revision = $result.Fields.Item('U_Revision').Value
			LineNum  = $result.Fields.Item('U_LineNum').Value
		};
	}
 catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Couldn't retrive additional information for Bill Of Materials with ItemCode: {0}, Revision: {1}, Details: {2}", $BOMItemCode, $BOMRevision, $err);
		throw [System.Exception]($msg);
	}
}

function getBOMCoproductInfo($token, $BOMItemCode, $BOMRevision, $ItemCode, $ItemRevision, $Sequence) {
	try {
		$qmBOMCoproductInfo.ClearParameters();
		$qmBOMCoproductInfo.AddParameter("BOMItemCode", $BOMItemCode);
		$qmBOMCoproductInfo.AddParameter("BOMRevision", $BOMRevision);
		$qmBOMCoproductInfo.AddParameter("ItemCode", $ItemCode);
		$qmBOMCoproductInfo.AddParameter("ItemRevision", $ItemRevision);
		$qmBOMCoproductInfo.AddParameter("Sequence", $Sequence);
		$result = $qmBOMCoproductInfo.Execute($token); 

		if ($result.RecordCount -ne 1) {
			throw [System.Exception]("Coproduct in given Bill Of Materials don't exists");
		}

		return [psobject]@{
			Code     = $result.Fields.Item('Code').Value
			ItemCode = $result.Fields.Item('U_ItemCode').Value
			Revision = $result.Fields.Item('U_Revision').Value
			LineNum  = $result.Fields.Item('U_LineNum').Value
		};
	}
 catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Couldn't retrive additional Coproduct information for Bill Of Materials with ItemCode: {0}, Revision: {1}, Coproduct ItemCode: {2}, Coproduct Revision: {3}, Coproduct Sequence: {4}. Details: {5}", $BOMItemCode, $BOMRevision, $ItemCode, $ItemRevision, $Sequence, $err);
		throw [System.Exception]($msg);
	}
}

function getBOMScrapInfo($token, $BOMItemCode, $BOMRevision, $ItemCode, $ItemRevision, $Sequence) {
	try {
		$qmBOMScrapInfo.ClearParameters();
		$qmBOMScrapInfo.AddParameter("BOMItemCode", $BOMItemCode);
		$qmBOMScrapInfo.AddParameter("BOMRevision", $BOMRevision);
		$qmBOMScrapInfo.AddParameter("ItemCode", $ItemCode);
		$qmBOMScrapInfo.AddParameter("ItemRevision", $ItemRevision);
		$qmBOMScrapInfo.AddParameter("Sequence", $Sequence);
		$result = $qmBOMScrapInfo.Execute($token); 

		if ($result.RecordCount -ne 1) {
			throw [System.Exception]("Scrap in given Bill Of Materials don't exists");
		}

		return [psobject]@{
			Code     = $result.Fields.Item('Code').Value
			ItemCode = $result.Fields.Item('U_ItemCode').Value
			Revision = $result.Fields.Item('U_Revision').Value
			LineNum  = $result.Fields.Item('U_LineNum').Value
		};
	}
 catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Couldn't retrive additional Scrap information for Bill Of Materials with ItemCode: {0}, Revision: {1}, Scrap ItemCode: {2}, Scrap Revision: {3}, Scrap Sequence: {4}. Details: {5}", $BOMItemCode, $BOMRevision, $ItemCode, $ItemRevision, $Sequence, $err);
		throw [System.Exception]($msg);
	}
}
#endRegion



try {

	#Data loading from a csv file
	write-host ""

	[array]$csvTemplates = Import-Csv -Delimiter ';' -Path $csvTemplatesPath;
	[array]$csvTemplatesLinesPath = Import-Csv -Delimiter ';' -Path $csvTemplatesLinesPath;
	#region Preparing Data
	write-Host 'Preparing data: '
	$totalRows = $csvTemplates.Count + $csvTemplatesLinesPath.Count;

	$templates = New-Object 'System.Collections.Generic.List[array]'
	$dictionaryTemplateLines = New-Object 'System.Collections.Generic.Dictionary[string,psobject]'

	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;

	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}

	foreach ($row in $csvTemplates) {
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		$templates.Add([array]$row);
	}

	foreach ($row in $csvTemplatesLinesPath) {
		#region progress
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		#endregion
		$tCode = $row.TemplateCode;
		$BOMItemCode = $row.BOMItemCode;
		$BOMRevision = $row.BOMRevision;
		$BOMKey = $BOMItemCode + '___' + $BOMRevision;

		if ($dictionaryTemplateLines.ContainsKey($tCode) -eq $false) {
			$dictionaryTemplateLines[$tCode] = New-Object 'System.Collections.Generic.Dictionary[string,psobject]';
		}
		if ($dictionaryTemplateLines[$tCode].ContainsKey($BOMKey) -eq $false) {
			$dictionaryTemplateLines[$tCode].Add($BOMKey,
				[psobject] @{
					BOMItemCode = $BOMItemCode
					BOMRevision = $BOMRevision
					Lines       = New-Object 'System.Collections.Generic.List[psobject]';
				});
		}
		$dictionaryTemplateLines[$tCode][$BOMKey].Lines.Add(
			[psobject] @{
				TemplateCode = $row.TemplateCode
				BOMItemCode  = $row.BOMItemCode
				BOMRevision  = $row.BOMRevision
				ItemType     = $row.ItemType
				ItemCode     = $row.ItemCode
				ItemRevision = $row.ItemRevision
				Sequence     = $row.Sequence
				RoutingCode  = $row.RoutingCode
				U_BaseLine   = $null
			});
	}
	Write-Host '';
	#endregion
	Write-Host 'Adding/updating data: ';
	if ($templates.Count -gt 1) {
		$total = $templates.Count
	}
	else {
		$total = 1
	}
	$progressItterator = 0;
	$progress = 0;
	$beforeProgress = 0;

	foreach ($csvTempl in $templates) {
		try {
			$progressItterator++;
			$progress = [math]::Round(($progressItterator * 100) / $total);
			if ($progress -gt $beforeProgress) {
				Write-Host $progress"% " -NoNewline
				$beforeProgress = $progress
			}
			$dictionaryKey = $csvTempl.TemplateCode;

			#Creating object
			$template = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::OrderlessProductionTemplate);

			#Checking it entry already exists
			$exists = $false;
			try {
				$retValue = $template.GetByKey($csvTempl.TemplateCode);
				if ($retValue -eq 0) {
					$exists = $true;
				}
			}
			catch {
				$exists = $false;
			}
			if (-not $exists) {
				$template.Code = $csvTempl.TemplateCode;
			}
			
			$template.Name = $csvTempl.TemplateName;
			if ([string]::IsNullOrWhiteSpace($csvTempl.Date) -eq $false) {
				$template.U_Date = $csvTempl.Date;
			}

			#Deleting all existing items
			$count = $template.Lines.Count;
			for ($i = $count - 1; $i -ge 0; $i--) {
				$dummy = $template.Lines.DelRowAtPos($i);
			}
			$template.Lines.SetCurrentLine(0);
			
			if ($dictionaryTemplateLines.ContainsKey($csvTempl.TemplateCode) -eq $false) { 
				throw [System.Exception]("Lines for template not found in csv file.");
			}
			$linesForTemplate = $dictionaryTemplateLines[$csvTempl.TemplateCode];
			$dictionaryBOMCodeBOMKey = New-Object 'System.Collections.Generic.Dictionary[string,string]';
			foreach ($BOMKey in $linesForTemplate.Keys) {
				$BOMInfo = $linesForTemplate[$BOMKey];
				$BOMItemCode = $BOMInfo.BOMItemCode;
				$BOMRevision = $BOMInfo.BOMRevision;

				$addBOMInfo = getBOMInfo -token $pfcCompany.Token -BOMItemCode $BOMItemCode -BOMRevision $BOMRevision;

				if ($template.Lines.IsRowFilled() -eq $true) {
					$template.Lines.Add();
				}
				$template.Lines.U_BomCode = $addBOMInfo.Code;
				$dictionaryBOMCodeBOMKey.Add($addBOMInfo.Code, $BOMKey);

				#Add BOM U_LineNum to lines from CSV

				foreach ($line in $BOMInfo.Lines) {
					switch ($line.ItemType) {
						"H" { 
							$addInfo = $addBOMInfo;
							break;
						}
						"C" { 
							$addInfo = getBOMCoproductInfo -token $pfcCompany.Token -BOMItemCode $line.BOMItemCode -BOMRevision $line.BOMRevision -ItemCode $line.ItemCode -ItemRevision $line.ItemRevision -Sequence $line.Sequence
							break;
						}
						"S" { 
							$addInfo = getBOMScrapInfo -token $pfcCompany.Token -BOMItemCode $line.BOMItemCode -BOMRevision $line.BOMRevision -ItemCode $line.ItemCode -ItemRevision $line.ItemRevision -Sequence $line.Sequence
							break;
						}
						Default {
							$msg = [string]::Format("Incorrect ItemType: {0}. BOMItemCode: {1}, BOMRevision: {2}, ItemCode: {3}, Revision: {4}, Sequence: {5}", [string]$line.ItemType, [string]$line.BOMItemCode, [string]$line.BOMRevision, [string]$line.ItemCode, [string]$line.ItemRevision, [string]$line.Sequence);
							throw [System.Exception]($msg);
						}
					}
					$line.U_BaseLine = $addInfo.LineNum;
				}
			}

			#find lines to update and delete
			$lineNumbersToDel = New-Object 'System.Collections.Generic.List[int]'
			for ($lineNum = 0; $lineNum -lt $template.Lines.Count; $lineNum++) {
				$template.Lines.SetCurrentLine($lineNum);
				
				if ($dictionaryBOMCodeBOMKey.ContainsKey($template.Lines.U_BomCode) -eq $false) {
					$msg = [string]::Format("Couldn't finde mapping for U_BOMCode: {0}.", [string]$template.Lines.U_BomCode);
					throw [System.Exception]($msg);
				}

				$BOMKey = $dictionaryBOMCodeBOMKey[$template.Lines.U_BomCode];
				$linesToCheck = $linesForTemplate[$BOMKey].Lines;

				$lineFound = $false;
				foreach ($ltc in  $linesToCheck) {
					if ($ltc.ItemType -eq $template.Lines.U_ItemType -and $ltc.U_BaseLine -eq $template.Lines.U_BaseLine) {
						if ($ltc.ItemType -eq "H" -and ([string]::IsNullOrWhiteSpace($ltc.RoutingCode) -eq $false)) {
							$template.Lines.U_RtgCode = $ltc.RoutingCode;
						}
						$lineFound = $true;
						break;
					}
				}
				if ($lineFound -eq $false) {
					$lineNumbersToDel.Add($lineNum);
				}
			}

			#remove lines
			for ($i = $lineNumbersToDel.Count - 1; $i -ge 0 ; $i--) {
				$dummy = $template.Lines.DelRowAtPos($lineNumbersToDel[$i]);
			}
			#Adding or updating templates
			$message = 0;
    
			if ($exists -eq $true) {
				$message = $template.Update()
			}
			else {
				$message = $template.Add()
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
			$ms = [string]::Format("Error when {0} Orderless Production Template {1}. Details: {2}", $taskMsg, $csvTempl.TemplateCode, $err);
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