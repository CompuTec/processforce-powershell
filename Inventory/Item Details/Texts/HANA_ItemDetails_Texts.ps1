clear
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.SQLServer = "10.0.0.38:30015"
$pfcCompany.SQLUserName = "SYSTEM"
$pfcCompany.SQLPassword = "password"
$pfcCompany.Databasename = "PFDEMO"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"
        
$code = $pfcCompany.Connect()
if($code -eq 1)
{
#Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
$csvItems = Import-Csv -Delimiter ';' -Path "c:\ItemDetails.csv"
 
#Checking that Item Details already exist 
 foreach($csvItem in $csvItems) 
 {
  
        $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
     
        $rs.DoQuery([string]::Format( "SELECT T0.""ItemCode"" FROM OITM T0
            LEFT OUTER JOIN ""@CT_PF_OIDT"" T1 ON T0.""ItemCode"" = T1.""U_ItemCode""
            WHERE
            T1.""U_ItemCode"" = N'{0}'", $csvItem.ItemCode))
        $exists = 0;
        if($rs.RecordCount -gt 0)
        {
            $exists = 1
        }
  
    #Creating Item Details 
    $idt = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"ItemDetails")
     
    #Restoring Item Costs and setting Inherit Batch/Serial to 'Yes'
    if($exists -eq 1)
    {
        $idt.GetByItemCode($csvItem.ItemCode)
    }
    else
    {
        $idt.U_ItemCode = $csvItem.ItemCode;
        $idt.CFG_RestoreItemCosting = "Y";
    }
     
    $idt.U_InheritBatch = 1;	
    $idt.U_InheritSerial = 1;
	
     
    #Data loading from the csv file - Texts for itmes from ItemDetails.csv file
    [array]$Texts = Import-Csv -Delimiter ';' -Path "c:\ItemDetails_Texts.csv" | Where-Object {$_.ItemCode -eq $csvItem.ItemCode}
    if($Texts.count -gt 0)
    {
        #Deleting all exisitng Phrases
        $count = $idt.Texts.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $idt.Texts.DelRowAtPos(0);
        }
        $idt.Texts.SetCurrentLine($idt.Texts.Count - 1);
         
        #Adding Phrases
        foreach($Text in $Texts) 
        {
			$idt.Texts.U_TxtCode = $Text.TextCode;
			if($Text.ProductionOrders -eq 1)
			{
				$idt.Texts.U_ProdOrders = "Y"
			}
			if($Text.ShipmentDocuments -eq 1)
			{
				$idt.Texts.U_ShipDoc =  "Y"
			}	
			if ( $Text.PickLists -eq 1)
			{
				$idt.Texts.U_PickLists =  "Y"
			}
			if ($Text.MSDS -eq 1)
			{
				$idt.Texts.U_MSDS =  "Y"
			}	
			if ( $Text.PurchaseOrders -eq 1)
			{
				$idt.Texts.U_PurOrders =  "Y"
			}	
			if ($Text.Returns -eq 1)
			{
				$idt.Texts.U_Returns ="Y"
			}	
			if ($Text.Other -eq 1)
			{
				$idt.Texts.U_Other =  "Y"
			}
            $dummy = $idt.Texts.Add()
        }
    }
  
    $message = 0
     
    #Adding or updating ItemDetails depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Item Details: {0}", $csvItem.ItemCode)
        $message = $idt.Update()
    }
    else
    {
        [System.String]::Format("Adding Item Details: {0}", $csvItem.ItemCode)
       $message= $idt.Add()
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