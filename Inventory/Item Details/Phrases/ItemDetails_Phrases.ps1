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
$pfcCompany.Databasename = "SBODemoPL"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2012

        
$code = $pfcCompany.Connect()
if($code -eq 1)
{
#Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
$csvItems = Import-Csv -Delimiter ';' -Path "c:\PSDI\ItemDetails.csv"
 
#Checking that Item Details already exist 
 foreach($csvItem in $csvItems) 
 {
  
        $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
     
        $rs.DoQuery([string]::Format( "SELECT T0.ItemCode FROM OITM T0
            LEFT OUTER JOIN [@CT_PF_OIDT] T1 ON T0.ItemCode = T1.U_ItemCode
            WHERE
            T1.U_ItemCode = N'{0}'", $csvItem.ItemCode))
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
	
     
    #Data loading from the csv file - Groups for items from ItemDetails.csv file
    [array]$phrases = Import-Csv -Delimiter ';' -Path "c:\PSDI\ItemDetails_Phrases.csv" | Where-Object {$_.ItemCode -eq $csvItem.ItemCode}
    if($phrases.count -gt 0)
    {
        #Deleting all exisitng Phrases
        $count = $idt.Phrases.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $idt.Phrases.DelRowAtPos(0);
        }
        $idt.Phrases.SetCurrentLine($idt.Phrases.Count - 1);
         
        #Adding Phrases
        foreach($phrase in $phrases) 
        {
			$idt.Phrases.U_PhCode = $phrase.PhraseCode;
			if($phrase.ProductionOrders -eq 1)
			{
				$idt.Phrases.U_ProdOrders = "Y"
			}
			if($phrase.ShipmentDocuments -eq 1)
			{
				$idt.Phrases.U_ShipDoc =  "Y"
			}	
			if ( $phrase.PickLists -eq 1)
			{
				$idt.Phrases.U_PickLists =  "Y"
			}
			if ($phrase.MSDS -eq 1)
			{
				$idt.Phrases.U_MSDS =  "Y"
			}	
			if ( $phrase.PurchaseOrders -eq 1)
			{
				$idt.Phrases.U_PurOrders =  "Y"
			}	
			if ($phrase.Returns -eq 1)
			{
				$idt.Phrases.U_Returns ="Y"
			}	
			if ($phrase.Other -eq 1)
			{
				$idt.Phrases.U_Other =  "Y"
			}
            $dummy = $idt.Phrases.Add()
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