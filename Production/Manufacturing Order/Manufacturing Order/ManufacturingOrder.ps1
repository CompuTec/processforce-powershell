clear
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLPassword = "sa"
$pfcCompany.SQLServer = "localhost"
$pfcCompany.SQLUserName = "sa"
$pfcCompany.Databasename = "PFDemo915" 
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2012

$code = 0;
$code = $pfcCompany.Connect()
if($code -eq 1)
{
	
	#Data loading from a csv file
	$csvItems = Import-Csv -Delimiter ';' -Path "C:\ManufacturingOrder.csv"
	
	foreach($csvItem in $csvItems) 
 	{
	    #Creating BOM object
	    $mo = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ManufacturingOrder)
		$bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::BillOfMaterial)
		$bom.GetByItemCodeAndRevision($csvItem.ItemCode, $csvItem.Revision);
		$mo.U_BOMCode = $bom.Code;
		$mo.U_RtgCode = $csvItem.Routing
		$mo.U_Warehouse = $csvItem.Warehouse
		$mo.U_Quantity = $csvItem.Quantity
		$mo.U_Factor = $csvItem.Factor
		$mo.U_RequiredDate = $csvItem.RequiredDate
		switch ($csvItem.Status) {
			"RL" {
				$status = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Released
				break
			}
			"ST" {
				$status = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Started
				break
			}
			"FI" {
				$status = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Finished
				break
			}
			default {
				$status = [CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus]::Scheduled
				break
			}
		}
		$mo.U_Status = $status
		$mo.CalculateManufacturingTimes($false);
		$count = $mo.RoutingOperations.Count;
		
		
		[array] $csvOperations = Import-Csv -Delimiter ';' -Path "C:\ManufacturingOrderOperations.csv" | Where-Object {$_.Key -eq $csvItem.Key}
	
		
		foreach($csvOper in $csvOperations) 
		{	
			
			for($i=0; $i -le $count; $i++)
			{
				$mo.RoutingOperations.SetCurrentLine($i);
				if($mo.RoutingOperations.U_OprSequence -eq $csvOper.Sequence)
				{
					$mo.RoutingOperations.U_Status = $csvOper.Status;
					break;
				}
			}
		}
		
		
		
		
		$message = 0
    
	    #Adding Maufacturing Order depends on exists in a database
        [System.String]::Format("Adding Manufacturing Order. CSV Key: {0}", $csvItem.Key)
	    $message = $mo.Add();
	}

	
	
	$pfcCompany.Disconnect()
}
else
{
	Write-Host -BackgroundColor Red -ForegroundColor White 'Connection failed'
}