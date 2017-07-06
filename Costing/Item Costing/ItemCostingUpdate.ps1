clear
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")


#### Before running this script please restore Item Costing Details. ####
#### This script allows only to update Item Costing on categories different than 000 ####

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
$csvItemCostings = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Costing\ItemCosting.csv"

 foreach($csvItemCosting in $csvItemCostings) 
 {
    #Creating Item Costing Object
    $ic = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"ItemCosting")
	if($csvItemCosting.CostCategory -ne '000')
	{
    #Checking if ItemCosting exists
    $retValue = $ic.Get($csvItemCosting.ItemCode, $csvItemCosting.Revision, $csvItemCosting.CostCategory)
	
   	if($retValue)
   	{
   
	   	#Data loading from the csv file - Costing Details for positions from ItemCosting.csv file
	    [array]$csvCostingDetails = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Costing\ItemCostingDetails.csv" | Where-Object {$_.ItemCode -eq $csvItemCosting.ItemCode -and $_.Revision -eq $csvItemCosting.Revision -and $_.Category -eq $csvItemCosting.CostCategory}
	    if($csvCostingDetails.count -gt 0)
	    {
			
	        foreach($csvCD in $csvCostingDetails)
	        {
			
				$count = $ic.CostingDetails.Count;
				for ($i=0; $i -lt $count ; $i++)
				{
					$ic.CostingDetails.SetCurrentLine($i);
					if($ic.CostingDetails.U_WhsCode -eq $csvCD.WhsCode)
					{
						#ML - Manual, MN - Manual no Roll-up, PL - Price List, PN - Price List no Roll-up, AC - Automatic, AN - Automatic no Roll-up
						$ic.CostingDetails.U_Type = $csvCD.Type
						$ic.CostingDetails.U_PriceList = $csvCD.PriceListCode
						$ic.CostingDetails.U_WhenZero = $csvCD.WhenZero
						$ic.CostingDetails.U_ItemCost = $csvCD.ItemCost
						$ic.CostingDetails.U_FixOH = $csvCD.FixedOH
						$ic.CostingDetails.U_FixOHPrct = $csvCD.FixedOHPrct
						$ic.CostingDetails.U_FixOHOther = $csvCD.FixedOHOther
						$ic.CostingDetails.U_VarOH = $csvCD.VariableOH
						$ic.CostingDetails.U_VarOHPrct = $csvCD.VariableOHPrct
						$ic.CostingDetails.U_VarOHOther = $csvCD.VariableOHOther
						$ic.CostingDetails.U_Remarks = $csvCD.Remarks
						break;
					}
				}
	            
	        }
	    }
  
  		
	  	$ic.RecalculateCostingDetails()
		$ic.RecalculateRolledCosts()
	    $message = 0
	
	
		[System.String]::Format("Updating Item Costing Details for Item: {0} Revision: {1} Category: {2}", $csvItemCosting.ItemCode, $csvItemCosting.Revision, $csvItemCosting.CostCategory)
		$message = $ic.Update()
				 
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
  	write-host -backgroundcolor red -foregroundcolor white "Item Costing Details for Item: "  $csvItemCosting.ItemCode   " Revision: "  $csvItemCosting.Revision  " Category: " $csvItemCosting.CostCategory  " don't exists";
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
