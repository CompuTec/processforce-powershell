clear
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLPassword = "sa"
$pfcCompany.SQLServer = "localhost"
$pfcCompany.SQLUserName = "sa"
$pfcCompany.Databasename = "PFDemo"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]"dst_MSSQL2012"
       
$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$csvItems = Import-Csv -Delimiter ';' -Path "C:\Operations.csv"

 foreach($csvItem in $csvItems) 
 {
    #Creating Operation object
    $operation = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"Operation")
    #Checking that the operation already exist  
    $retValue = $operation.GetByOprCode($csvItem.OperationCode)
    if($retValue -ne 0)
   {     
    #Adding the new data
    $operation.U_OprCode = $csvItem.OperationCode
   }
    #Data loading from a csv file - Operation Properties
    $operation.U_OprName = $csvItem.OperationName
	$operation.U_Remarks = $csvItem.Remarks
	[array]$resProps = Import-Csv -Delimiter ';' -Path "C:\Operations_Properties.csv" | Where-Object {$_.OperationCode -eq $csvItem.OperationCode}
    if($resProps.count -gt 0)
    {
        #Deleting all existing properties
        $count = $operation.OperationProperties.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $operation.OperationProperties.DelRowAtPos(0);
        }
        
        #Adding the new data       
        foreach($prop in $resProps) 
        {
		    $operation.OperationProperties.U_PrpCode = $prop.PropertiesCode
		    $operation.OperationProperties.U_PrpConType = $prop.Condition #enum ConditionType; Equal EQ = 1, NotEqual NE = 2, GratherThan GT = 3, GratherThanOrEqual GE = 4, LessThan LT = 5, LessThanOrEqual LE = 6, Between BT = 7
    		$operation.OperationProperties.U_PrpConValue = $prop.Value
    		$operation.OperationProperties.U_PrpConValueTo = $prop.ToValue
    		$operation.OperationProperties.U_UnitOfMeasure = $prop.UoM
            $dummy = $operation.OperationProperties.Add()
        }
        
    }
    
    #Adding resources to Operations
    [array]$opResource = Import-Csv -Delimiter ';' -Path "C:\Operations_Resources.csv" | Where-Object {$_.OperationCode -eq $csvItem.OperationCode}
    if($opResource.count -gt 0)
    {
        #Deleting all existing resources
        $count = $operation.OperationResources.Count-1
        if($count -gt 1)
        {
        for($i=0; $i -lt $count; $i++)
        {
   
           
            $dummy = $operation.OperationResources.DelRowAtPos(0); 
        }
        }
        $resourcesDict = New-Object 'System.Collections.Generic.Dictionary[String,int]'
        #Adding the new data
        foreach($opRes in $opResource) 
        {
      
            $operation.OperationResources.U_RscCode = $opRes.ResourceCode
            $operation.OperationResources.U_IsDefault = $opRes.Default # Y = Yes, N = No
    		$operation.OperationResources.U_QueueTime = $opRes.QueTime
            $operation.OperationResources.U_QueueRate = $opRes.QueTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
    		$operation.OperationResources.U_SetupTime = $opRes.SetupTime
            $operation.OperationResources.U_SetupRate = $opRes.SetupTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
    		$operation.OperationResources.U_RunTime = $opRes.RunTime
            $operation.OperationResources.U_RunRate = $opRes.RunTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
    		$operation.OperationResources.U_StockTime = $opRes.StockTime		
            $operation.OperationResources.U_StockRate = $opRes.StockTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
            if($operation.OperationResources.U_RscType -eq 'T')
			{
				$operation.OperationResources.U_MachineCode = $opRes.MachineCode
			}
			if($opRes.Cycles -eq 'Y')
			{
				$operation.OperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes
				$operation.OperationResources.U_CycleCap = $opRes.CycleCapacity
			}
			else
			{
				$operation.OperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
			}
			$operation.OperationResources.U_Project = $opRes.Project
			$operation.OperationResources.U_OcrCode = $opRes.Dimension1
			$operation.OperationResources.U_OcrCode2 = $opRes.Dimension2
			$operation.OperationResources.U_OcrCode3 = $opRes.Dimension3
			$operation.OperationResources.U_OcrCode4 = $opRes.Dimension4
			$operation.OperationResources.U_OcrCode5 = $opRes.Dimension5
			
			
			$operation.OperationResources.U_Remarks = $opRes.Remarks
			#$operation.OperationResources.UDFItems.Item("U_UDF1").Value = $opRes.U_UDF1 # how to add UDF
			
			$resourcesDict.Add($operation.OperationResources.U_RscCode,$operation.OperationResources.U_OprRscCode);
            $dummy = $operation.OperationResources.Add()
            
        }
    }
	
	#Adding resources properties to Operations
    [array]$opResourceProperties = Import-Csv -Delimiter ';' -Path "C:\Operations_ResourcesProperties.csv" | Where-Object {$_.OperationCode -eq $csvItem.OperationCode}
    if($opResourceProperties.count -gt 0)
    {
        #Deleting all existing resources
        $count = $operation.OperationResourceProperties.Count-1
        if($count -gt 1)
        {
        for($i=0; $i -lt $count; $i++)
        {
   
           
            $dummy = $operation.OperationResourceProperties.DelRowAtPos(0); 
        }
        }
        
        #Adding the new data
        foreach($opResProp in $opResourceProperties) 
        {
      
            $operation.OperationResourceProperties.U_OprCode = $opResProp.OperationCode
			$operation.OperationResourceProperties.U_OprRscCode = $resourcesDict[$opResProp.ResourceCode]
			$operation.OperationResourceProperties.U_PrpCode = $opResProp.PropertiesCode
			$operation.OperationResourceProperties.U_PrpConType = $opResProp.Condition
			$operation.OperationResourceProperties.U_PrpConValue = $opResProp.Value
			$operation.OperationResourceProperties.U_PrpConValueTo = $opResProp.ToValue
			$operation.OperationResourceProperties.U_UnitOfMeasure = $opResProp.UoM
			
			#$operation.OperationResourceProperties.UDFItems.Item("U_UDF1").Value = $opRes.U_UDF1 # how to add UDF
            $dummy = $operation.OperationResourceProperties.Add()
            
        }
    }
  
    $message = 0
    
    #Adding or updating Operations depends on exists in the database
    if($retValue -eq 0)
    {
        
        [System.String]::Format("Updating Opertion: {0}", $csvItem.OperationCode)
        $message = $operation.Update()
    }
    else
    {
        try
        {
        [System.String]::Format("Adding Operation: {0}", $csvItem.OperationCode)
        $message= $operation.Add()
        }
        catch [Exception]
        {
            Write-Host $_.Exception.InnerException.ToString()
        }
	}
    if($message -lt 0)
    {    
	    $err=$pfcCompany.GetLastErrorDescription()
	    write-host -backgroundcolor red -foregroundcolor white $err
	}    
  }
}