#region #Script info
Clear-Host
########################################################################
# CompuTec PowerShell Script - Import Resources Calendars
########################################################################
$SCRIPT_VERSION = "3.2"
# Last tested PF version: ProcessForce 9.3 (9.30.140) PL: 04 R1 HF1 (64-bit)
# Description:
#      Import Resources Calendars. Script add new or will update existing Resources Calendars.
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

$csvResourcesCalendarPath = -join ($csvImportCatalog, "ResourcesCalendar.csv")
$csvResourcesCalendarWorkingHoursPath = -join ($csvImportCatalog, "ResourcesCalendarWorkingHours.csv")
$csvResourcesCalendarHolidaysPath = -join ($csvImportCatalog, "ResourcesCalendarHolidays.csv")
$csvResourcesCalendarExceptionsPath = -join ($csvImportCatalog, "ResourcesCalendarExceptions.csv")

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
	[array] $csvResCalendar = Import-Csv -Delimiter ';' $csvResourcesCalendarPath;
	
	if ((Test-Path -Path $csvResourcesCalendarWorkingHoursPath -PathType leaf) -eq $true) {
		[array] $csvResCalendarWorkingHours = Import-Csv -Delimiter ';' $csvResourcesCalendarWorkingHoursPath;
	}
	else {
		[array] $csvResCalendarWorkingHours = $null;
		write-host "Resources Working Hours - csv not available."
	}
	
	if ((Test-Path -Path $csvResourcesCalendarExceptionsPath -PathType leaf) -eq $true) {
		[array] $csvResCalendarExceptions = Import-Csv -Delimiter ';' $csvResourcesCalendarExceptionsPath;
	}
	else {
		[array] $csvResCalendarExceptions = $null;
		write-host "Resources Exceptions - csv not available."
	}
	
	if ((Test-Path -Path $csvResourcesCalendarHolidaysPath -PathType leaf) -eq $true) {
		[array] $csvResCalendarHolidays = Import-Csv -Delimiter ';' $csvResourcesCalendarHolidaysPath;
	}
	else {
		[array] $csvResCalendarHolidays = $null;
		write-host "Resources Holidays - csv not available."
	}

	write-Host 'Preparing data: '
	$totalRows = $csvResCalendar.Count + $csvResCalendarWorkingHours.Count + $csvResCalendarExceptions.Count + $csvResCalendarHolidays.Count;
    

	$resCalendarlist = New-Object 'System.Collections.Generic.List[array]'
	$workingHoursDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$exceptionsDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
	$holidaysDict = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

	$progressItterator = 0;
	$progres = 0;
	$beforeProgress = 0;
	
	if ($totalRows -gt 1) {
		$total = $totalRows
	}
	else {
		$total = 1
	}

	foreach ($row in $csvResCalendar) {
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
		$resCalendarlist.Add([array]$row);
	}

	foreach ($row in $csvResCalendarWorkingHours) {
		$key = $row.ResourceCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($workingHoursDict.ContainsKey($key) -eq $false) {
			$workingHoursDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $workingHoursDict[$key];
		
		$list.Add([array]$row);
	}

	foreach ($row in $csvResCalendarExceptions) {
		$key = $row.ResourceCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($exceptionsDict.ContainsKey($key) -eq $false) {
			$exceptionsDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $exceptionsDict[$key];
		
		$list.Add([array]$row);
	}

	foreach ($row in $csvResCalendarHolidays) {
		$key = $row.ResourceCode;
		$progressItterator++;
		$progres = [math]::Round(($progressItterator * 100) / $total);
		if ($progres -gt $beforeProgress) {
			Write-Host $progres"% " -NoNewline
			$beforeProgress = $progres
		}
	
		if ($holidaysDict.ContainsKey($key) -eq $false) {
			$holidaysDict[$key] = New-Object System.Collections.Generic.List[array];
		}
		$list = $holidaysDict[$key];
		
		$list.Add([array]$row);
	}

	Write-Host '';
	Write-Host 'Adding/updating data: ' -NoNewline;

	if ($resCalendarlist.Count -gt 1) {
		$total = $resCalendarlist.Count
	}
	else {
		$total = 1
	}
	$progressItterator = 0;
	$progress = 0;
	$beforeProgress = 0;

	foreach ($csvItem in $resCalendarlist) {
		try {
			$progressItterator++;
			$progress = [math]::Round(($progressItterator * 100) / $total);
			if ($progress -gt $beforeProgress) {
				Write-Host $progress"% " -NoNewline
				$beforeProgress = $progress
			}
			#Creating ResourceCalendar object
			$res = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"ResourceCalendar")
			#Checking that the calendar already exist
			$retVal = $res.GetByResourceCode($csvItem.ResourceCode)
			if ($retVal -ne $true) {
				$err = [string]::Format("Calendar for Resource with code: {0} don't exists", $csvItem.ResourceCode);
				throw [System.Exception]($err);
			}
    
			#Working Hours
			[array]$resWH = $workingHoursDict[$csvItem.ResourceCode];
			if ($resWH.count -gt 0) {
				#Deleting all existing working hours
				$count = $res.WorkingHours.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $res.WorkingHours.DelRowAtPos(0);
				}
				#		$res.WorkingHours.SetCurrentLine($res.WorkingHours.Count-1);
				#Adding the new data       
				foreach ($wh in $resWH) {
					$res.WorkingHours.U_Day = $wh.Day
					$res.WorkingHours.U_FromTime = $wh.FromTime
					$res.WorkingHours.U_ToTime = $wh.ToTime
					$dummy = $res.WorkingHours.Add()
				}
			}
    
			#Adding exceptions to Resources
			[array]$resExc = $exceptionsDict[$csvItem.ResourceCode];
			if ($resExc.count -gt 0) {
				#Deleting all existing exceptions
				$count = $res.WorkingHoursExceptions.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $res.WorkingHoursExceptions.DelRowAtPos(0);
				}
				#        $res.WorkingHoursExceptions.SetCurrentLine($res.WorkingHoursExceptions.Count-1);
				#Adding the new data
				foreach ($whe in $resExc) {
					$res.WorkingHoursExceptions.U_Date = $whe.Date
					$res.WorkingHoursExceptions.U_FromTime = $whe.FromTime
					$res.WorkingHoursExceptions.U_ToTime = $whe.ToTime
					$res.WorkingHoursExceptions.U_Remarks = $whe.Remarks
					$dummy = $res.WorkingHoursExceptions.Add()
				}
			}
 
			#Adding holidays to Resources
			[array]$resHol = $holidaysDict[$csvItem.ResourceCode];
			if ($resHol.count -gt 0) {
				#Deleting all existing exceptions
				$count = $res.Holidays.Count
				for ($i = 0; $i -lt $count; $i++) {
					$dummy = $res.Holidays.DelRowAtPos(0);
				}
				#        $res.Holidays.SetCurrentLine($res.Holidays.Count-1);
				#Adding the new data
				foreach ($hol in $resHol) {
					$res.Holidays.U_Date = $hol.Date
					$res.Holidays.U_Remarks = $hol.Remarks
					$dummy = $res.Holidays.Add()
				}
    
			}
			$message = 0
    
			#Updating Resources calendars depends on exists in the database
			$message = $res.Update()  
			if ($message -lt 0) {    
				$err = $pfcCompany.GetLastErrorDescription()
				Throw [System.Exception] ($err)
			}
		}
		Catch {
			$err = $_.Exception.Message;
			$taskMsg = "updating"
			$ms = [string]::Format("Error when {0} Calendar for Resource with Code {1} Details: {2}", $taskMsg, $csvItem.ResourceCode, $err);
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

