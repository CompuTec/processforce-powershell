[CmdletBinding()]
Param(
	[Parameter(Mandatory = $True)]
	[string]$categories, #in case of cost rollover this is cost category from
	[string]$costCategoryTo = '',
	[Parameter(Mandatory = $True)]
	[string]$databaseName,
	[Parameter(Mandatory = $True)]
	[string]$jobType, #rollUp, rollOver
	[string]$itemFrom = '',
	[string]$itemTo = '',
	[string]$warehouses = ''
)
. $PSScriptRoot\Logger.ps1

#------DEFINITIONS - CHANGE ONLY THIS SECION---------
#Path to folder with file
$PATH_TO_SCRIPTS = "C:\Users\HANA10DEV\powershell-scripts\Costing\Costing Roll Up Roll Over\"
#RollUp script filename
$ROLLUP_SCRIPT_FILENAME = "CostingRollUpParam.ps1"
#Roll over script filename
$ROLLOVER_SCRIPT_FILENAME = "CostingRollOverParam.ps1"


#Authorization data
$USERNAME = "manager"
$PASSWORD = "1234"
$SQL_SERVER = "DEV@hanadev:30013"
$LICENSE_SERVER = "hanadev:40000"
$SERVER_TYPE = "dst_HANADB"

#Debug disabled/enabled - set this option to $true to get more detailed information in Computec event log
$debugFlag = 'enabled'

#---------------END OF DEFINITIONS--------------------


$log = [Logger]::new("CostingLauncher");
function WriteLog($msg) {
	$log.WriteLog($msg);
}


$fullRollUpScriptFileName = [String]::Concat($PATH_TO_SCRIPTS, $ROLLUP_SCRIPT_FILENAME);
$fullRollOverScriptFileName = [String]::Concat($PATH_TO_SCRIPTS, $ROLLOVER_SCRIPT_FILENAME);


if (-not $itemFrom -gt '') {
	$itemFrom = '*'
}
if (-not $itemTo -gt '') {
	$itemTo = '*'
}

if (-not $warehouses -gt '') {
	$warehouses = '*'
}

if ($debugFlag -eq 'enabled') {
	$msg = [string]::Format("Costing procedure started for parameters: categories: {0},costCategoryTo: {1}, databaseName: {2}, jobType: {3}, itemFrom: {4}, itemTo: {5}, warehouses: {6} ",
		$categories, $costCategoryTo, $databaseName, $jobType, $itemFrom, $itemTo, $warehouses);
	WriteLog $msg;
}
try {

	$x = [string]::Empty
	if ($jobType -eq 'RollUp') {
		$x = powershell.exe -File $fullRollUpScriptFileName -categories $categories -databaseName $databaseName -itemFrom $itemFrom -itemTo $itemTo -warehouses $warehouses -USERNAME $USERNAME -PASSWORD $PASSWORD -SQL_SERVER $SQL_SERVER -LICENSE_SERVER $LICENSE_SERVER -SERVER_TYPE $SERVER_TYPE -debugFlag $debugFlag
	}

	if ($jobType -eq 'RollOver') {
		$x = powershell.exe -File $fullRollOverScriptFileName -costCategoryFrom $categories -costCategoryTo $costCategoryTo -databaseName $databaseName -itemFrom $itemFrom -itemTo $itemTo -warehouses $warehouses -USERNAME $USERNAME -PASSWORD $PASSWORD -SQL_SERVER $SQL_SERVER -LICENSE_SERVER $LICENSE_SERVER -SERVER_TYPE $SERVER_TYPE -debugFlag $debugFlag
	}
	if([string]::IsNullOrEmpty($x) -eq $false) {
		throw [Exception]::new($x);
	}
}
catch {
	WriteLog ( "Running procedure failed: " + $_.Exception.Message )
}