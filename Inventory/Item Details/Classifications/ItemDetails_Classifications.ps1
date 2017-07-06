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
$pfcCompany.Databasename = "localhost"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2012

        
$code = $pfcCompany.Connect()
if($code -eq 1)
{
#Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
$csvItems = Import-Csv -Delimiter ';' -Path "c:\ItemDetails.csv"
 
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
	
     
    #Data loading from the csv file - Classifications for itmes from ItemDetails.csv file
    [array]$classifications = Import-Csv -Delimiter ';' -Path "c:\ItemDetails_Classification.csv" | Where-Object {$_.ItemCode -eq $csvItem.ItemCode}
    if($classifications.count -gt 0)
    {
        #Deleting all exisitng Classification
        $count = $idt.Classifications.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $idt.Classifications.DelRowAtPos(0);
        }
        $idt.Classifications.SetCurrentLine($idt.Classifications.Count - 1);
         
        #Adding Properies
        foreach($classification in $classifications) 
        {
			$idt.Classifications.U_ClsCode = $classification.ClassificationCode;
			if($classification.ProductionOrders -eq 1)
			{
				$idt.Classifications.U_ProdOrders = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
			}
			if($classification.ShipmentDocuments -eq 1)
			{
				$idt.Classifications.U_ShipDoc =  [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
			}	
			if ( $classification.PickLists -eq 1)
			{
				$idt.Classifications.U_PickLists =  [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
			}
			if ($classification.MSDS -eq 1)
			{
				$idt.Classifications.U_MSDS =  [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
			}	
			if ( $classification.PurchaseOrders -eq 1)
			{
				$idt.Classifications.U_PurOrders =  [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
			}	
			if ($classification.Returns -eq 1)
			{
				$idt.Classifications.U_Returns =  [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
			}	
			if ($classification.Other -eq 1)
			{
				$idt.Classifications.U_Other =  [CompuTec.ProcessForce.API.Enumerators.YesNoType]::Yes;
			}
            $dummy = $idt.Classifications.Add()
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