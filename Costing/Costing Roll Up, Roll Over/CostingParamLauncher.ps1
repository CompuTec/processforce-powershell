[CmdletBinding()]
Param(
 	[Parameter(Mandatory=$True)]
   	[string]$categories, #in case of cost rollover this is cost category from
   	[string]$costCategoryTo = '',
	[Parameter(Mandatory=$True)]
	[string]$databaseName,
    [Parameter(Mandatory=$True)]
	[string]$jobType, #rollUp, rollOver
	[string]$itemFrom = '',
	[string]$itemTo = '',
	[string]$warehouses = ''
)


#------DEFINITIONS - CHANGE ONLY THIS SECION---------
#Path to folder with file
$PATH_TO_SCRIPTS = "C:\PS\PF\Costing\Schedule\"
#RollUp script filename
$ROLLUP_SCRIPT_FILENAME = "CostingRollUpParam.ps1"
#Roll over script filename
$ROLLOVER_SCRIPT_FILENAME = "CostingRollOverParam.ps1"


#Authorization data
$USERNAME = "manager"
$PASSWORD = "1234"
$SQL_SERVER = "localhost"
$SQL_USERNAME = "sa"
$SQL_PASSWORD = "sa"
$SERVER_TYPE = "dst_MSSQL2012"

#Debug disabled/enabled - set this option to $true to get more detailed information in Computec event log
$debugFlag = 'enabled'

#---------------END OF DEFINITIONS--------------------



$fullRollUpScriptFileName = [String]::Concat($PATH_TO_SCRIPTS, $ROLLUP_SCRIPT_FILENAME);
$fullRollOverScriptFileName = [String]::Concat($PATH_TO_SCRIPTS, $ROLLOVER_SCRIPT_FILENAME);


if(-not $itemFrom -gt '')
{
    $itemFrom = '*'
}
if(-not $itemTo -gt '')
{
    $itemTo = '*'
}

if(-not $warehouses -gt '')
{
    $warehouses = '*'
}

if($debugFlag -eq 'enabled')
{
$msg = [string]::Format("Costing procedure started for parameters: categories: {0},costCategoryTo: {1}, databaseName: {2}, jobType: {3}, itemFrom: {4}, itemTo: {5}, warehouses: {6} ",
$categories,$costCategoryTo,$databaseName,$jobType,$itemFrom, $itemTo, $warehouses);

Write-EventLog -LogName Computec -Source "Computec ProcessForce" -EntryType Information -EventId 1 -Message $msg;
}

if($jobType -eq 'RollUp')
{
   powershell.exe -File $fullRollUpScriptFileName -categories $categories -databaseName $databaseName -itemFrom $itemFrom -itemTo $itemTo -warehouses $warehouses -USERNAME $USERNAME -PASSWORD $PASSWORD -SQL_SERVER $SQL_SERVER -SQL_USERNAME $SQL_USERNAME -SQL_PASSWORD $SQL_PASSWORD -SERVER_TYPE $SERVER_TYPE -debugFlag $debugFlag
}

if($jobType -eq 'RollOver')
{
    powershell.exe -File $fullRollOverScriptFileName -costCategoryFrom $categories -costCategoryTo $costCategoryTo -databaseName $databaseName -itemFrom $itemFrom -itemTo $itemTo -warehouses $warehouses -USERNAME $USERNAME -PASSWORD $PASSWORD -SQL_SERVER $SQL_SERVER -SQL_USERNAME $SQL_USERNAME -SQL_PASSWORD $SQL_PASSWORD -SERVER_TYPE $SERVER_TYPE -debugFlag $debugFlag
}