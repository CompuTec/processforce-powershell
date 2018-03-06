Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLPassword = "pass"
$pfcCompany.SQLServer = "localhost"
$pfcCompany.SQLUserName = "sa"
$pfcCompany.Databasename = "PFDemo"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]"dst_MSSQL2012"
    

$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$csvItems = Import-Csv -Delimiter ';' -Path "C:\powershell-scripts\Production\Routings\Resources\Resources.csv"


 foreach($csvItem in $csvItems) 
 {
    #Creating Resource object
    $res = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"Resource")
    #Checking that the resource already exist
    $retVal = $res.GetByRscCode($csvItem.ResourceCode)
    if($retValue -ne 0)
    {
    #Adding the new data
    $res.U_RscType = $csvItem.ResourceType #enum type; Machine = 1 or M, Labour = 2 or L, Tool = 3 or T, Subcontractor = 4 or S 
    $res.U_RscCode = $csvItem.ResourceCode
    $res.U_RscName = $csvItem.ResourceName
    $res.U_RscGrpCode = $csvItem.ResourceGroup
	$res.U_QueueTime = $csvItem.QueTime
    $res.U_QueueRate = $csvItem.QueTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9 
	$res.U_SetupTime = $csvItem.SetupTime
    $res.U_SetupRate = $csvItem.SetupTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9 
	$res.U_RunTime = $csvItem.RunTime
    $res.U_RunRate = $csvItem.RunTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9 
	$res.U_StockTime = $csvItem.StockTime
    $res.U_StockRate = $csvItem.StockTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9     
    $res.U_ResourceCount = $csvItem.ResourceNumber
    $res.U_HasCycles = $csvItem.HasCycle #enum type; 1 = Yes, 2 = No
    $res.U_CycleCap = $csvItem.CycleCapacity
	$res.U_ResActCode = $csvItem.ResourceAccountingCode
	$res.U_Project = $csvItem.Project
	$res.U_OcrCode = $csvItem.Dimension1
	$res.U_OcrCode2 = $csvItem.Dimension2
	$res.U_OcrCode3 = $csvItem.Dimension3
	$res.U_OcrCode4 = $csvItem.Dimension4
	$res.U_OcrCode5 = $csvItem.Dimension5
    $res.U_WhsCode = $csvItem.IssueWhsCode
    $res.U_BinAbs = $csvItem.IssueBinAbs
    $res.U_RWhsCode = $csvItem.ReceiptWhsCode
    $res.U_RBinAbs = $csvItem.ReceiptBinAbs

    
	#$res.UDFItems.Item("U_UDF1").Value = $csvItem.UDF1 ## how to import UDFs
	
	if($res.U_RscType -eq 'S')
	{
		$res.U_VendorCode = $csvItem.VendorCode
		$res.U_ItemCode = $csvItem.ItemCode
	}
   }           
    
    #Data loading from a csv file - Resource Properties
    [array]$resProps = Import-Csv -Delimiter ';' -Path "C:\powershell-scripts\Production\Routings\Resources\Resources_Properties.csv" | Where-Object {$_.ResourceCode -eq $csvItem.ResourceCode}
    if($resProps.count -gt 0)
    {
        #Deleting all existing properties
        $count = $res.Properties.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $res.Properties.DelRowAtPos(0);
        }
        
        #Adding the new data       
        foreach($prop in $resProps) 
        {
		    $res.Properties.U_PrpCode = $prop.PropertiesCode
		    $res.Properties.U_PrpConType = $prop.Condition #enum ConditionType; Equal EQ = 1, NotEqual NE = 2, GratherThan GT = 3, GratherThanOrEqual GE = 4, LessThan LT = 5, LessThanOrEqual LE = 6, Between BT = 7
    		$res.Properties.U_PrpConValue = $prop.Value
    		$res.Properties.U_PrpConValueTo = $prop.ToValue
    		$res.Properties.U_UnitOfMeasure = $prop.UoM
            $dummy = $res.Properties.Add()
        }
        
        
        
    }
    
    #Adding attachments to Resources
    [array]$resAttachments = Import-Csv -Delimiter ';' -Path "C:\powershell-scripts\Production\Routings\Resources\Resources_Attachments.csv" | Where-Object {$_.ResourceCode -eq $csvItem.ResourceCode}
    if($resAttachments.count -gt 0)
    {
        #Deleting all existing attachments
        $count = $res.Attachments.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $res.Attachments.DelRowAtPos(0);
        }
        
        #Adding the new data
        foreach($att in $resAttachments) 
        {
            $fileName = [System.IO.Path]::GetFileName($att.AttachmentPath)
            $res.Attachments.U_FileName = $fileName
		    $res.Attachments.U_AttDate = [System.DateTime]::Today
		    $res.Attachments.U_Path = $att.AttachmentPath
            $dummy = $res.Attachments.Add()
        }
    }
 
    $message = 0
    
    #Adding or updating Resources depends on exists in the database
    if($retVal -eq 0)
    {
        [System.String]::Format("Updating Resource: {0}", $csvItem.ResourceCode)
        $message = $res.Update()
    }
    else
    {
        [System.String]::Format("Adding Resource: {0}", $csvItem.ResourceCode)
        $message= $res.Add()
	}
    
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
else
{
write-host "Failure"
}
