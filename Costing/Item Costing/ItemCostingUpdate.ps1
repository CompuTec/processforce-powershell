Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")


#### Before running this script please restore Item Costing Details. ####
#### This script allows only to update Item Costing on categories different than 000 ####

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()

$pfcCompany.UserName = "manager";
$pfcCompany.Password = "1234";
$pfcCompany.SQLPassword = "pass";
$pfcCompany.SQLServer = "10.0.0.1:30115";
$pfcCompany.LicenseServer = "10.0.0.2:40000";
$pfcCompany.SQLUserName = "SYSTEM";
$pfcCompany.Databasename = "PFDEMO";
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_HANADB

$code = $pfcCompany.Connect()
if ($code -eq 1) {

	Write-Host 'Preparing data: '
	#Data loading from a csv file
	[array]$csvItemCostings = Import-Csv -Delimiter ';' -Path ($PSScriptRoot + "\ItemCosting.csv");
	[array]$csvItemCostingDetails = Import-Csv -Delimiter ';' -Path ($PSScriptRoot + "\ItemCostingDetails.csv");
	$dictionaryItemCosting = New-Object 'System.Collections.Generic.Dictionary[string,psobject]'

	$totalRows = $csvItemCostings.Count + $csvItemCostingDetails.Count 
	$progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
    if($totalRows -gt 1) {
        $total = $totalRows
    } else {
        $total = 1
    }

	
    foreach($row in $csvItemCostings){
        $key = $row.ItemCode + '__' + $row.Revision + '__' + $row.CostCategory;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
		}

		if (-Not $dictionaryItemCosting.ContainsKey($key)) {
			$dictionaryItemCosting.Add($key, [psobject]@{
				ItemCode = $row.ItemCode
				Revision = $row.Revision
				CostCategory = $row.CostCategory
				Details = New-Object 'System.Collections.Generic.List[array]'});
        }
	}
	
	foreach ($row in $csvItemCostingDetails) {
        $key = $row.ItemCode + '__' + $row.Revision + '__' + $row.Category;
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }

        if ($dictionaryItemCosting.ContainsKey($key)) {
			$list = $dictionaryItemCosting[$key].Details;
			$list.Add([array]$row);
        }
	}
	
    Write-Host '';
	Write-Host 'Add/Update data to SAP: '

	$totalRows = $dictionaryItemCosting.Count
	$progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
    if($totalRows -gt 1) {
        $total = $totalRows
    } else {
        $total = 1
	}
	
    foreach ($key in $dictionaryItemCosting.Keys) {
		$csvItemCosting = $dictionaryItemCosting[$key];
		$progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if ($progres -gt $beforeProgress) {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        #Creating Item Costing Object
        $ic = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ItemCosting);
        if ($csvItemCosting.CostCategory -ne '000') {
            #Checking if ItemCosting exists
            $retValue = $ic.Get($csvItemCosting.ItemCode, $csvItemCosting.Revision, $csvItemCosting.CostCategory)
	
            if ($retValue) {
   
                #Data loading from the csv file - Costing Details for positions from ItemCosting.csv file
                #[array]$csvCostingDetails = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Costing\ItemCostingDetails.csv" | Where-Object {$_.ItemCode -eq $csvItemCosting.ItemCode -and $_.Revision -eq $csvItemCosting.Revision -and $_.Category -eq $csvItemCosting.CostCategory}
				$csvCostingDetails = $csvItemCosting.Details
                    foreach ($csvCD in $csvCostingDetails) {
			
                        $count = $ic.CostingDetails.Count;
                        for ($i = 0; $i -lt $count ; $i++) {
                            $ic.CostingDetails.SetCurrentLine($i);
                            if ($ic.CostingDetails.U_WhsCode -eq $csvCD.WhsCode) {
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
                
                $ic.RecalculateCostingDetails()
                $ic.RecalculateRolledCosts()
                $message = 0
	
	
                Write-Host -NoNewline ([System.String]::Format("Updating Item Costing Details for Item: {0} Revision: {1} Category: {2} ", $csvItemCosting.ItemCode, $csvItemCosting.Revision, $csvItemCosting.CostCategory));
                $message = $ic.Update()
				 
                if ($message -lt 0) {    
                    $err = $pfcCompany.GetLastErrorDescription()
                    write-host -backgroundcolor red -foregroundcolor white "Failure: " + $err
                } 
                else {
                    write-host -BackgroundColor Yellow -ForegroundColor Black "Success"
                }   
		
            }
            else {
                write-host -backgroundcolor red -foregroundcolor white "Item Costing Details for Item: "  $csvItemCosting.ItemCode   " Revision: "  $csvItemCosting.Revision  " Category: " $csvItemCosting.CostCategory  " don't exists";
            }
        }
        else {
            write-host -backgroundcolor red -foregroundcolor white "Masive update for Cost Category 000 is turned off - please make updates on custom Cost Category and use Roll-Over functionality";
        }
    }
}
else {
    write-host "Failure: " $pfcCompany.GetLastErrorDescription()
}
