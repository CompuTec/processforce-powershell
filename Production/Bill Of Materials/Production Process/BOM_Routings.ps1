Celar-Host
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
    [array]$csvItems = Import-Csv -Delimiter ';' -Path "C:\BOM_Header.csv"
    [array]$bomRoutings = Import-Csv -Delimiter ';' -Path "C:\BOM_Routings.csv" 
    [array]$bomRoutingsOperations = Import-Csv -Delimiter ';' -Path "C:\BOM_Routings_Operations.csv" 
    [array]$bomRoutingsOperationsProperties = Import-Csv -Delimiter ';' -Path "C:\BOM_Routings_Operations_Properties.csv" 
    [array]$bomRoutingsOperationsResources = Import-Csv -Delimiter ';' -Path "C:\BOM_Routings_Operations_Resources.csv" 
    [array]$opResourceProperties = Import-Csv -Delimiter ';' -Path "C:\BOM_Routings_Operations_Resources_Properties.csv" 


    write-Host 'Preparing data: '
    $totalRows = $csvItems.Count + $bomRoutings.Count + $bomRoutingsOperations.Count + $bomRoutingsOperationsProperties.Count + $bomRoutingsOperationsResources.Count + $opResourceProperties.Count

    $bomList = New-Object 'System.Collections.Generic.List[array]'

    $dictionaryRoutings = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryRoutingsOperations = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryRoutingsOperationsProperties = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryRoutingsOperationsResources = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'
    $dictionaryResourceProperties = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;

    if($totalRows -gt 1) {
        $total = $totalRows
    } else {
        $total = 1
    }

    foreach($row in $csvItems){
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if($progres -gt $beforeProgress)
        {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

  
        $bomList.Add([array]$row);
    }

    foreach($row in $bomRoutings){
        $key = $row.BOM_Header + '___' + $row.Revision;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if($progres -gt $beforeProgress)
        {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if($dictionaryRoutings.ContainsKey($key)){
            $list = $dictionaryRoutings[$key];
        } else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryRoutings[$key] = $list;
        }
    
        $list.Add([array]$row);
    }

    foreach($row in $bomRoutingsOperations){
        $key = $row.BOM_Header + '___' + $row.Revision + '___' + $row.RoutingCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if($progres -gt $beforeProgress)
        {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if($dictionaryRoutingsOperations.ContainsKey($key)){
            $list = $dictionaryRoutingsOperations[$key];
        } else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryRoutingsOperations[$key] = $list;
        }
    
        $list.Add([array]$row);
    }

    foreach($row in $bomRoutingsOperationsProperties){
        $key = $row.BOM_Header + '___' + $row.Revision + '___' + $row.RoutingCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if($progres -gt $beforeProgress)
        {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if($dictionaryRoutingsOperationsProperties.ContainsKey($key)){
            $list = $dictionaryRoutingsOperationsProperties[$key];
        } else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryRoutingsOperationsProperties[$key] = $list;
        }
    
        $list.Add([array]$row);
    }

    
    foreach($row in $bomRoutingsOperationsResources){
        $key = $row.BOM_Header + '___' + $row.Revision + '___' + $row.RoutingCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if($progres -gt $beforeProgress)
        {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if($dictionaryRoutingsOperationsResources.ContainsKey($key)){
            $list = $dictionaryRoutingsOperationsResources[$key];
        } else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryRoutingsOperationsResources[$key] = $list;
        }
    
        $list.Add([array]$row);
    }
    
    foreach($row in $opResourceProperties){
        $key = $row.BOM_Header + '___' + $row.Revision + '___' + $row.RoutingCode;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if($progres -gt $beforeProgress)
        {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if($dictionaryResourceProperties.ContainsKey($key)){
            $list = $dictionaryResourceProperties[$key];
        } else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryResourceProperties[$key] = $list;
        }
    
        $list.Add([array]$row);
    }
    Write-Host '';


 foreach($csvItem in $bomList) 
 {
    $dictionaryKey = $csvItem.BOM_Header + '___' + $csvItem.Revision;
    

    #Creating BOM object
    $bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"BillOfMaterial")
    #Checking that the BOM already exist
    $retValue = $bom.GetByItemCodeAndRevision($csvItem.BOM_Header, $csvItem.Revision)
   if($retValue -ne 0)
   {
    $bom.U_ItemCode = $csvItem.BOM_Header
    $bom.U_Revision = $csvItem.Revision
   }   
    #Data loading from a csv file - Routing
    #[array]$bomRoutings = Import-Csv -Delimiter ';' -Path "C:\BOM_Routings.csv" | Where-Object {$_.BOM_Header -eq $csvItem.BOM_Header -and $_.Revision -eq $csvItem.Revision}
    $bomRoutings = $dictionaryRoutings[$dictionaryKey];
    if($bomRoutings.count -gt 0)
    {
        #Deleting all existing routings, operations, resources
        
        $count = $bom.Routings.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $bom.Routings.DelRowAtPos(0);
        }
        
        $count = $bom.RoutingOperations.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $bom.RoutingOperations.DelRowAtPos(0);
        }
                
        $count = $bom.RoutingOperationResources.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $bom.RoutingOperationResources.DelRowAtPos(0);
        }     
        
		$count = $bom.RoutingOperationProperties.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $bom.RoutingOperationProperties.DelRowAtPos(0);
        }
        
        #Adding a new data - Routings
        foreach($rtg in $bomRoutings) 
        {  
            $bom.Routings.U_RtgCode = $rtg.RoutingCode
		    $bom.Routings.U_IsDefault = $rtg.DefaultMRP # Y = Yes, N = No
            $bom.Routings.U_IsRollUpDefault = $rtg.DefaultCosting # Y = Yes, N = No
		    $dummy = $bom.Routings.Add()
        }
        #Deleting defautl operationes copied from Routing
      
        while($bom.RoutingOperations.Count -ne 1 )
        {  
            $count = $bom.RoutingOperations.Count
            $nextint=0;
        
            for($i=0; $i -lt $count; $i++)
            {
                try{
                    $dummy = $bom.RoutingOperations.DelRowAtPos($nextint);
                    }catch
                    {
                    $nextint++
                    }
            
            }
		}
        $dummy = $bom.RoutingOperations.DelRowAtPos(0);
       
        $drivers = New-Object 'System.Collections.Generic.Dictionary[String,int]'
        #Adding a new data - Operations for Routings
        foreach($rtg in $bomRoutings) 
        {
        
            #[array]$bomRoutingsOperations = Import-Csv -Delimiter ';' -Path "C:\BOM_Routings_Operations.csv" | Where-Object {$_.BOM_Header -eq $csvItem.BOM_Header -and $_.Revision -eq $csvItem.Revision -and $_.RoutingCode -eq $rtg.RoutingCode}
            $dictionaryKeyRt = $csvItem.BOM_Header + '___' + $csvItem.Revision + '___' + $rtg.RoutingCode;
            $bomRoutingsOperations = $dictionaryRoutingsOperations[$dictionaryKeyRt];
            $bom.RoutingOperations.SetCurrentLine($bom.RoutingOperations.Count - 1)
            $overlayDict = New-Object 'System.Collections.Generic.Dictionary[int,int]';
            foreach($rtgOper in $bomRoutingsOperations) 
            {
                $bom.RoutingOperations.U_RtgCode = $rtgOper.RoutingCode   
		        $bom.RoutingOperations.U_OprCode = $rtgOper.OperationCode      
                $bom.RoutingOperations.U_OprSequence = $rtgOper.Sequence

                if($rtgOper.OperationOverlayCode -gt ''){
                    $bom.RoutingOperations.U_OprOverlayCode = $rtgOper.OperationOverlayCode;
                    $bom.RoutingOperations.U_OprOverlayId = $overlayDict[$rtgOper.OperationOverlaySequence];
                    $bom.RoutingOperations.U_OprOverlayQty = $rtgOper.OperationOverlayQty;
                }

                $overlayDict.Add($rtgOper.Sequence,$bom.RoutingOperations.U_LineNum);
                $drivers_key = $rtgOper.RoutingCode + '@#@' + $bom.RoutingOperations.U_OprSequence;
                $drivers.Add($drivers_key,$bom.RoutingOperations.U_RtgOprCode);
                $dummy = $bom.RoutingOperations.Add()
            }
			
			#operation properties
			#[array]$bomRoutingsOperationsProperties = Import-Csv -Delimiter ';' -Path "C:\BOM_Routings_Operations_Properties.csv" | Where-Object {$_.BOM_Header -eq $csvItem.BOM_Header -and $_.Revision -eq $csvItem.Revision -and $_.RoutingCode -eq $rtg.RoutingCode}	
			$bomRoutingsOperationsProperties = $dictionaryRoutingsOperationsProperties[$dictionaryKeyRt];
            if($bomRoutingsOperationsProperties.count -gt 0)
		    {
		        #Deleting all existing properties
		        $count = $bom.RoutingOperationProperties.Count
		        for($i=0; $i -lt $count; $i++)
		        {
		            $dummy = $bom.RoutingOperationProperties.DelRowAtPos(0);
		        }
		        
		        #Adding the new data       
		        foreach($prop in $bomRoutingsOperationsProperties) 
		        {
                    $drivers_key = $prop.RoutingCode + '@#@' + $prop.Sequence;
					$bom.RoutingOperationProperties.U_RtgOprCode = $drivers[$drivers_key];
					$bom.RoutingOperationProperties.U_RtgCode = $prop.RoutingCode
					$bom.RoutingOperationProperties.U_OprCode = $prop.OperationCode
				    $bom.RoutingOperationProperties.U_PrpCode = $prop.PropertiesCode
				    $bom.RoutingOperationProperties.U_PrpConType = $prop.Condition #enum ConditionType; Equal EQ = 1, NotEqual NE = 2, GratherThan GT = 3, GratherThanOrEqual GE = 4, LessThan LT = 5, LessThanOrEqual LE = 6, Between BT = 7
		    		$bom.RoutingOperationProperties.U_PrpConValue = $prop.Value
		    		$bom.RoutingOperationProperties.U_PrpConValueTo = $prop.ToValue
		    		$bom.RoutingOperationProperties.U_UnitOfMeasure = $prop.UoM
		            $dummy = $bom.RoutingOperationProperties.Add()
		        }
		        
		    }
			
        
         
            #Deleting default resources copied from operations   
            $count = $bom.RoutingOperationResources.Count
            for($i=0; $i -lt $count; $i++)
            {
                $dummy = $bom.RoutingOperationResources.DelRowAtPos(0);
            }    
            $count = $bom.RoutingsOperationResourceProperties.Count
            for($i=0; $i -lt $count; $i++)
            {
        
                $dummy = $bom.RoutingsOperationResourceProperties.DelRowAtPos(0);      
            }
		     $driversRtgOprRsc = New-Object 'System.Collections.Generic.Dictionary[String,int]'
             #Adding resources for operations   
        
            #[array]$bomRoutingsOperationsResources = Import-Csv -Delimiter ';' -Path "C:\BOM_Routings_Operations_Resources.csv" | Where-Object {$_.BOM_Header -eq $csvItem.BOM_Header -and $_.Revision -eq $csvItem.Revision -and $_.RoutingCode -eq $rtg.RoutingCode}
            $dictionaryKeyRt = $csvItem.BOM_Header + '___' + $csvItem.Revision + '___' + $rtg.RoutingCode;
            $bomRoutingsOperationsResources = $dictionaryRoutingsOperationsResources[$dictionaryKeyRt];
            if($bomRoutingsOperationsResources.count -gt 0)
            {
                foreach($rtgOperResc in $bomRoutingsOperationsResources) 
                {
                    $drivers_key = $rtgOperResc.RoutingCode + '@#@' + $rtgOperResc.Sequence;
                    $bom.RoutingOperationResources.U_RtgCode = $rtgOperResc.RoutingCode
	                $bom.RoutingOperationResources.U_OprCode = $rtgOperResc.OperationCode
                    $bom.RoutingOperationResources.U_RtgOprCode =$drivers[$drivers_key];
                    $bom.RoutingOperationResources.U_RscCode = $rtgOperResc.ResourceCode
                    $bom.RoutingOperationResources.U_IsDefault = $rtgOperResc.Default
                    $bom.RoutingOperationResources.U_IssueType = $rtgOperResc.IssueType;
                    $bom.RoutingOperationResources.U_QueueTime = $rtgOperResc.QueTime
                    $queTimeUoM = $rtgOperResc.QueTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3
                    switch($queTimeUoM)
                    {
                        "1" { $bom.RoutingOperationResources.U_QueueRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedSeconds }
                        "2" { $bom.RoutingOperationResources.U_QueueRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedMinutes }
                        "3" { $bom.RoutingOperationResources.U_QueueRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedHours }

                    }

                    $bom.RoutingOperationResources.U_SetupTime = $rtgOperResc.SetupTime
                    $setupTimeUoM = $rtgOperResc.SetupTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3
                    switch($setupTimeUoM)
                    {
                        "1" { $bom.RoutingOperationResources.U_SetupRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedSeconds }
                        "2" { $bom.RoutingOperationResources.U_SetupRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedMinutes }
                        "3" { $bom.RoutingOperationResources.U_SetupRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedHours }

                    }

                    

                    $bom.RoutingOperationResources.U_RunTime = $rtgOperResc.RunTime
                    $runtimeUom = $rtgOperResc.RunTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
                    switch($runtimeUom)
                    {
                        "1" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedSeconds }
                        "2" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedMinutes }
                        "3" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedHours }
                        "4" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::SecondsPerPiece }
                        "5" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::MinutesPerPiece }
                        "6" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::HoursPerPiece }
                        "7" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::PiecesPerSecond }
                        "8" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::PiecesPerMinute }
                        "9" { $bom.RoutingOperationResources.U_RunRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::PiecesPerHour }

                    }
                    $bom.RoutingOperationResources.U_StockTime = $rtgOperResc.StockTime
                    $stockTimeUoM = $rtgOperResc.StockTimeUoM #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3 
                    switch($stockTimeUoM)
                    {
                        "1" { $bom.RoutingOperationResources.U_StockRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedSeconds }
                        "2" { $bom.RoutingOperationResources.U_StockRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedMinutes }
                        "3" { $bom.RoutingOperationResources.U_StockRate = [CompuTec.ProcessForce.API.Enumerators.RateType]::FixedHours }

                    }
				    if($rtgOperResc.NumberOfResources -ne '')
				    {
					    $bom.RoutingOperationResources.U_NrOfResources = $rtgOperResc.NumberOfResources
				    }
					
				    if($rtgOperResc.HasCycles -ne '')
				    {
						
					    if($rtgOperResc.HasCycles -eq 'Y')
					    {
							
						    $bom.RoutingOperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes
							
						    if($rtgOperResc.CycleCapacity -ne '')
						    {
							    $bom.RoutingOperationResources.U_CycleCap = $rtgOperResc.CycleCapacity
						    }
					    } else
					    {
						    $bom.RoutingOperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No
					    }
				    }
					
				    $bom.RoutingOperationResources.U_Remarks = $rtgOperResc.Remarks
				    if($rtgOperResc.Project -ne '')
				    {
					    $bom.RoutingOperationResources.U_Project = $rtgOperResc.Project
				    }
					
				    $key = $rtgOperResc.Sequence + '@#@' + $bom.RoutingOperationResources.U_RscCode
					
				    $driversRtgOprRsc.Add($key,$bom.RoutingOperationResources.U_RtgOprRscCode);
                    $dummy = $bom.RoutingOperationResources.Add()
                
			    }
				
			    #Adding resources properties to Operations
			    #[array]$opResourceProperties = Import-Csv -Delimiter ';' -Path "C:\BOM_Routings_Operations_Resources_Properties.csv" | Where-Object {$_.BOM_Header -eq $csvItem.BOM_Header -and $_.Revision -eq $csvItem.Revision -and $_.RoutingCode -eq $rtg.RoutingCode}
                $dictionaryKeyRt = $csvItem.BOM_Header + '___' + $csvItem.Revision + '___' + $rtg.RoutingCode;
                $opResourceProperties = $dictionaryResourceProperties[$dictionaryKeyRt];
			    if($opResourceProperties.count -gt 0)
			    {
			        #Deleting all existing resources
			        $count = $bom.RoutingsOperationResourceProperties.Count-1
			        if($count -gt 1)
			        {
			        for($i=0; $i -lt $count; $i++)
			        {
			   
			           
			            $dummy = $bom.RoutingsOperationResourceProperties.DelRowAtPos(0); 
			        }
			        }
			        
			        #Adding the new data
			        foreach($opResProp in $opResourceProperties) 
			        {
			      
				  	    $key = $opResProp.Sequence + '@#@' + $opResProp.RoutingCode
					    $drivers_key = $opResProp.RoutingCode + '@#@' + $opResProp.Sequence;
					
					    $bom.RoutingsOperationResourceProperties.U_RtgOprCode = $drivers[$drivers_key]
					    $bom.RoutingsOperationResourceProperties.U_RtOpRscCode = $driversRtgOprRsc[$key]
					    $bom.RoutingsOperationResourceProperties.U_RtgCode = $opResProp.RoutingCode
					    $bom.RoutingsOperationResourceProperties.U_OprCode = $opResProp.OperationCode
					    $bom.RoutingsOperationResourceProperties.U_PrpCode = $opResProp.PropertiesCode
					    $bom.RoutingsOperationResourceProperties.U_PrpConType = $opResProp.Condition
					    $bom.RoutingsOperationResourceProperties.U_PrpConValue = $opResProp.Value
					    $bom.RoutingsOperationResourceProperties.U_PrpConValueTo = $opResProp.ToValue
					    $bom.RoutingsOperationResourceProperties.U_UnitOfMeasure = $opResProp.UoM
						
					    #$bom.RoutingsOperationResourceProperties.UDFItems.Item("U_UDF1").Value = $opRes.U_UDF1 # how to add UDF
			            $dummy = $bom.RoutingsOperationResourceProperties.Add()
			            
			        }
			    }
            }
        }
    }
	
    $message = 0
    #Adding or updating BOMs depends on exists in the database
    if($retValue -eq 0)
    {
        [System.String]::Format("Updating BOM: {0}", $csvItem.BOM_Header)
        $message = $bom.Update()
    }
    else
    {
        [System.String]::Format("Adding BOM: {0}", $csvItem.BOM_Header)
        $message= $bom.Add()
	}
    
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
