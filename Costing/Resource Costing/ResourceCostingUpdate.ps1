clear
#### DI API path ####
[void] [Reflection.Assembly]::LoadFrom( "C:\Program Files\SAP\SAP Business One\AddOns\CT\ProcessForce\CompuTec.Core.DLL")
[void] [Reflection.Assembly]::LoadFrom( "C:\Program Files\SAP\SAP Business One\AddOns\CT\ProcessForce\CompuTec.ProcessForce.API.DLL")

#### Before running this script please restore Resource Costing. ####
####  This script allows only to update Resource Costing on categories different than 000 ####

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()

$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLPassword = "sa"
$pfcCompany.SQLServer = "localhost"
$pfcCompany.SQLUserName = "sa"
$pfcCompany.Databasename = "PFDemo"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2012"

$code = $pfcCompany.Connect()
if($code -eq 1)
{

#Data loading from a csv file
$csvResourceCostings = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Costing\ResourceCosting.csv"

 foreach($csvResourceCosting in $csvResourceCostings) 
 {
    #Creating Resource Costing Object
    $rc = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"ResourceCosting")
	if($csvResourceCosting.CostCategory -ne '000')
	{
    #Checking if ResourceCosting exists
    $retValue = $rc.Get($csvResourceCosting.ResourceCode, $csvResourceCosting.CostCategory)
	
   	if($retValue)
   	{
   
	   	#Data loading from the csv file - Costing Details for positions from ResourceCosting.csv file
	    [array]$csvCostingDetails = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Costing\ResourceCostingDetails.csv" | Where-Object {$_.ResourceCode -eq $csvResourceCosting.ResourceCode -and $_.Category -eq $csvResourceCosting.CostCategory}
	    if($csvCostingDetails.count -gt 0)
	    {
			
	        foreach($csvCD in $csvCostingDetails)
	        {
			
				$count = $rc.Costs.Count
				for ($i=0; $i -lt $count ; $i++)
				{
					$rc.Costs.SetCurrentLine($i);
					#QT - Queue Time, ST - Setup Time, RT - Run Time, TT - Stock Time
					if( $rc.Costs.U_CostType -eq $csvCD.CostType)
					{
						$rc.Costs.U_HourRate = $csvCD.HourlyRate
						$rc.Costs.U_FixOH = $csvCD.FixedOH
						$rc.Costs.U_FixOHPrct = $csvCD.FixedOHPrct
						$rc.Costs.U_FixOHOther = $csvCD.FixedOHOther
						$rc.Costs.U_VarOH = $csvCD.VariableOH
						$rc.Costs.U_VarOHPrct = $csvCD.VariableOHPrct
						$rc.Costs.U_VarOHOther = $csvCD.VariableOHOther
						$rc.Costs.U_Remarks = $csvCD.Remarks
						break;
					}
				}
	            
	        }
	    }
  
  		
	    $message = 0
	
	
		[System.String]::Format("Updating Resource Costing Details for Resource: {0}  Category: {1}", $csvResourceCosting.ResourceCode, $csvResourceCosting.CostCategory)
		$message = $rc.Update()
				 
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
  else
  {
  	write-host -backgroundcolor red -foregroundcolor white "Resource Costing Details for Resource: "  $csvResourceCosting.ResourceCode   " Category: " $csvResourceCosting.CostCategory  " don't exists";
  }
  }
  else
  {
  	write-host -backgroundcolor red -foregroundcolor white "Masive update for Cost Category 000 is turned off - please make updates on custom Cost Category and use Roll-Over functionality";
  }
}
}
else
{
write-host "Failure: " $pfcCompany.GetLastErrorDescription()
}
