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
$pfcCompany.Databasename = "PFDemo"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2012

        
$code = $pfcCompany.Connect()
if($code -eq 1)
{
#Data loading from a csv file - Items for which Item Details will be added (each of them has to has Item Master Data)
$csvItems = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Item Details\ItemDetailsProperites\ItemDetails.csv"
 
#Checking that Item Details already exist 
 foreach($csvItem in $csvItems) 
 {
  
        $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
     
        $rs.DoQuery([string]::Format( "SELECT T0.""ItemCode"" FROM OITM T0
            INNER JOIN ""@CT_PF_OIDT"" T1 ON T0.""ItemCode"" = T1.""U_ItemCode""
            WHERE
            T1.""U_ItemCode"" = N'{0}'", $csvItem.ItemCode))
        $exists = 0;
        if($rs.RecordCount -gt 0)
        {
        
  
            #Creating Item Details 
            $idt = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"ItemDetails")
            $idt.GetByItemCode($csvItem.ItemCode)
            $idt.U_InheritBatch = 1;	
            $idt.U_InheritSerial = 1;
	
     
            #Data loading from the csv file - Properties for itmes from ItemDetails.csv file
            [array]$properties = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Item Details\ItemDetailsProperites\ItemDetails_Properties.csv" | Where-Object {$_.ItemCode -eq $csvItem.ItemCode}
            if($properties.count -gt 0)
            {
                #Deleting all exisitng Properties
                $count = $idt.Properties.Count
                for($i=0; $i -lt $count; $i++)
                {
                    $dummy = $idt.Properties.DelRowAtPos(0);
                }
                $idt.Properties.SetCurrentLine($idt.Properties.Count - 1);
         
                #Adding Properies
                foreach($prop in $properties) 
                {
			        $idt.Properties.U_PrpCode = $prop.PropertyCode;
			


                   $idt.Properties.U_Expression = $prop.Expression;


			        if($prop.RangeFrom -ne "")
			        {
				        $idt.Properties.U_RangeValueFrom = $prop.RangeFrom;
			        }
			        else
			        {
				        $idt.Properties.U_RangeValueFrom = 0;
			        }
			        $idt.Properties.U_RangeValueTo = $prop.RangeTo;
			        if($prop.ReferenceCode -ne "")
			        {
				        $idt.Properties.U_WordCode = $prop.ReferenceCode;
			        }
                   $idt.Properties.U_Remarks = $prop.Remarks
                   $dummy = $idt.Properties.Add()
                }
            }
  
            $message = 0
     
            #Updating ItemDetails depends on exists in the database
            [System.String]::Format("Updating Item Details: {0}", $csvItem.ItemCode)


            $message = $idt.Update()  
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
          else
          {
                $msg = [System.String]::Format("Item Details for Item Code: {0} don't exists. Please run Restore Item Details.", $csvItem.ItemCode)
                write-host -backgroundcolor red -foregroundcolor white $msg
          }
  }
}
else
{
write-host "Failure"
}