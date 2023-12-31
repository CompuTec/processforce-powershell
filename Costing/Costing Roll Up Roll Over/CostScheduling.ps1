Import-Module -Name "PSScheduledJob"
$credential = Get-Credential -Credential '';
$namePrefix = 'CTPFCOSTING_'
$searchPrefix = 'CTPFCOSTING_*'

$jobs = Get-ScheduledJob -Name $searchPrefix
#deleting 
foreach ($job in $jobs)
{
	Unregister-ScheduledJob $job
}

# Set Up path to your schedule csv file
[array] $csv = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Costing\Schedule\CostingSchedule.csv"

foreach($csvPos in $csv) 
{	

	$jobname = $namePrefix + $csvPos.JobName;
	$categories = $csvPos.CostCategory;

    $costCategoryTo = ' ';
    if( $csvPos.CostCategoryTo -gt '')
    {
	    $costCategoryTo = $csvPos.CostCategoryTo;
    }
	$databaseName = $csvPos.Database;
	$itemFrom = '*';
	if( $csvPos.ItemFrom -gt '')
	{
		$itemFrom = $csvPos.ItemFrom 
	}
	$itemTo = '*';
	if( $csvPos.ItemTo -gt '')
	{
		$itemTo = $csvPos.ItemTo; 
	}
	
	$warehouses = '*';
	if( $csvPos.Warehouse -gt '')
	{
		$warehouses = $csvPos.Warehouse;
	}
	
	#Creating trigger
	$intervalType = $csvPos.IntervalType;
	$interval = $csvPos.Interval;
	
	$daysOfWeek = @();
	
	
	
	#$timeString = ;
	[DateTime] $time = $csvPos.Time
	
	$trigger;
	switch  ($intervalType) {
		"D"{
			$trigger = New-JobTrigger -Daily -At $time -DaysInterval $interval  
			break
		}
		"W" {
			#$daysString = $csvPos.W_DaysOfWeek;
			$daysString = $csvPos.W_DaysOfWeek.split(',');
			foreach($day in $daysString)
			{
				switch ($day) {
					'Mo' {
						$day = [DayOfWeek]::Monday
						break
					}
					'Tu' {
						$day = [DayOfWeek]::Tuesday
						break
					}
					'We' {
						$day = [DayOfWeek]::Wednesday
						break
					}
					'Th' {
						$day = [DayOfWeek]::Thursday
						break
					}
					'Fr' {
						$day = [DayOfWeek]::Friday
						break
					}
					'Sa' {
						$day = [DayOfWeek]::Saturday
						break
					}
					'Su' {
						$day = [DayOfWeek]::Sunday
						break
					}
					
				}
				$daysOfWeek = $daysOfWeek + $day;
			}
			$trigger = New-JobTrigger -Weekly -At $time -DaysOfWeek $daysOfWeek -WeeksInterval $interval
			break
		}
		"O" {
			
			$trigger = New-JobTrigger -Once -At $time 
			break
		}
	}
	
        #set up path to CostingParamlauncher.ps1
		Register-ScheduledJob -Name $jobname -FilePath "C:\PS\PF\Costing\Schedule\CostingParamLauncher.ps1" -ArgumentList $categories,$costCategoryTo,$databaseName,$csvPos.Type,$itemFrom,$itemTo,$warehouses -Trigger $trigger -Credential $credential -RunAs32;
	
}