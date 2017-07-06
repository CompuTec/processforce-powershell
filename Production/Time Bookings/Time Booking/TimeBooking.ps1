clear
#### DI API path ####
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")


#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLPassword = "sa"
$pfcCompany.SQLServer = "localhost"
$pfcCompany.SQLUserName = "sa"
$pfcCompany.Databasename = "PFDemo"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2012
        
$code = $pfcCompany.Connect()
if($code -eq 1)
{
#Data loading from a csv file
$csvItems = Import-Csv -Delimiter ';' -Path "C:\PS\PF\MO\TimeBookings\TimeBookings.csv"

foreach($csvItem in $csvItems) 
 {
    #Creating BOM object
    $otr = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"OperationTimeRecording")
	$otr.U_DocDate = $csvItem.DocDate;
	$otr.U_Remarks = $csvItem.Remarks;
	$otr.U_Ref2 = $csv.Ref2;
	
    #Data loading from a csv file - lines
    [array]$otrLinesCsv = Import-Csv -Delimiter ';' -Path "C:\PS\PF\MO\TimeBookings\TimeBookingLines.csv" | Where-Object {$_.Key -eq $csvItem.Key}
    foreach($otrLineCsv in $otrLinesCsv) 
    {
		
		$otr.Lines.U_BaseEntry = $otrLineCsv.BaseEntry;
		$otr.Lines.U_BaseDocNum = $ortLineCsv.BaseDocNum;
		$otr.Lines.U_RscCode = $otrLineCsv.ResourceCode;
		$otr.Lines.U_BaseLineNum = $otrLineCsv.BaseLineNum;
		$otr.Lines.U_OprCode = $otrLineCsv.OperationCode;
		$otr.Lines.U_Remarks = $otrLineCsv.Remarks;
		switch ($otrLineCsv.TimeType) {
			"Q" {
				$otr.Lines.U_TimeType = [CompuTec.ProcessForce.API.Enumerators.RecordingTimeType]::QueueTime;
				break
			}
			"S" {
				$otr.Lines.U_TimeType = [CompuTec.ProcessForce.API.Enumerators.RecordingTimeType]::SetupTime;
				break
			}
			"R" {
				$otr.Lines.U_TimeType = [CompuTec.ProcessForce.API.Enumerators.RecordingTimeType]::RunTime;
				break
			}
			"L" {
				$otr.Lines.U_TimeType = [CompuTec.ProcessForce.API.Enumerators.RecordingTimeType]::StockTime;
				break
			}
			default {
				break
			}
		}
		
		if($otrLineCsv.NumberOfResources -gt 1)
		{
			$otr.Lines.U_NrOfResources =  $otrLineCsv.NumberOfResources;
		}
		else
		{
			$otr.Lines.U_NrOfResources =  1;
		}
		
		if ($otrLineCsv.StartDate -ne '') {
			$otr.Lines.U_StartDate = $otrLineCsv.StartDate;
		}
		if ($otrLineCsv.StartTime -ne '') {
	        $otr.Lines.U_StartTime = $otrLineCsv.StartTime;
		}		
        if ($otrLineCsv.EndDate -ne '') {
			$otr.Lines.U_EndDate = $otrLineCsv.EndDate;
		}
		if ($otrLineCsv.EndTime -ne '') {
	        $otr.Lines.U_EndTime = $otrLineCsv.EndTime;
		}
		$otr.Lines.U_WorkingHours = $otrLineCsv.WorkingHours;
        $otr.Lines.Add();
	}	
	
    $message = 0
	[System.String]::Format("Adding TimeBokings with key: '{0}'", $csvItem.Key)
    $message= $otr.Add()
    
    if($message -lt 0)
    {    
	    $err=$pfcCompany.GetLastErrorDescription()
	    write-host -backgroundcolor red -foregroundcolor white $err
        write-host -backgroundcolor red -foregroundcolor white "Fail"
	}
    else
    {
        write-host "Success"
    }
  
}
}
else
{
write-host "Failure"
}