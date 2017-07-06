clear
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")


#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLPassword = "sa"
$pfcCompany.SQLServer = "localhost"
$pfcCompany.SQLUserName = "sa"
$pfcCompany.Databasename = "SBODemoPL"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]"dst_MSSQL2008"

       
$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$csvItems = Import-Csv -Delimiter ';' -Path "C:\ResourcesCalendar.csv"

 foreach($csvItem in $csvItems) 
 {
    #Creating ResourceCalendar object
    $res = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"ResourceCalendar")
    #Checking that the calendar already exist
    $retVal = $res.GetByResourceCode($csvItem.ResourceCode)
    if($retVal)
   {
    
    #Data loading from a csv file - Working Hours
    [array]$resWH = Import-Csv -Delimiter ';' -Path "C:\ResourcesCalendar_WorkingHours.csv" | Where-Object {$_.ResourceCode -eq $csvItem.ResourceCode}
    if($resWH.count -gt 0)
    {
        #Deleting all existing working hours
        $count = $res.WorkingHours.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $res.WorkingHours.DelRowAtPos(0);
        }
#		$res.WorkingHours.SetCurrentLine($res.WorkingHours.Count-1);
        #Adding the new data       
        foreach($wh in $resWH) 
        {
			$res.WorkingHours.U_Day = $wh.Day
			$res.WorkingHours.U_FromTime = $wh.FromTime
			$res.WorkingHours.U_ToTime = $wh.ToTime
            $dummy = $res.WorkingHours.Add()
        }
        
        
        
    }
    
    #Adding exceptions to Resources
    [array]$resExc = Import-Csv -Delimiter ';' -Path "C:\ResourcesCalendar_Exceptions.csv" | Where-Object {$_.ResourceCode -eq $csvItem.ResourceCode}
    if($resExc.count -gt 0)
    {
        #Deleting all existing exceptions
        $count = $res.WorkingHoursExceptions.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $res.WorkingHoursExceptions.DelRowAtPos(0);
        }
#        $res.WorkingHoursExceptions.SetCurrentLine($res.WorkingHoursExceptions.Count-1);
        #Adding the new data
        foreach($whe in $resExc) 
        {
            $res.WorkingHoursExceptions.U_Date = $whe.Date
			$res.WorkingHoursExceptions.U_FromTime = $whe.FromTime
			$res.WorkingHoursExceptions.U_ToTime = $whe.ToTime
			$res.WorkingHoursExceptions.U_Remarks = $whe.Remarks
            $dummy = $res.WorkingHoursExceptions.Add()
        }
    }
 
 	#Adding holidays to Resources
    [array]$resHol = Import-Csv -Delimiter ';' -Path "C:\ResourcesCalendar_Holidays.csv" | Where-Object {$_.ResourceCode -eq $csvItem.ResourceCode}
    if($resHol.count -gt 0)
    {
        #Deleting all existing exceptions
        $count = $res.Holidays.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $res.Holidays.DelRowAtPos(0);
        }
#        $res.Holidays.SetCurrentLine($res.Holidays.Count-1);
        #Adding the new data
        foreach($hol in $resHol) 
        {
            $res.Holidays.U_Date = $hol.Date
			$res.Holidays.U_Remarks = $hol.Remarks
            $dummy = $res.Holidays.Add()
        }
    }
    $message = 0
    
    #Updating Resources calendars depends on exists in the database
        [System.String]::Format("Updating Resource: {0}", $csvItem.ResourceCode)
        $message = $res.Update()    
    if($message -lt 0)
    {    
	    $err=$pfcCompany.GetLastErrorDescription()
	    write-host -backgroundcolor red -foregroundcolor white $err
	} 
    else
    {
    write-host "Success"
    }   
  }
  }
}
else
{
write-host "Failure"
}
