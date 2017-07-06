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
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2012"
        
$code = $pfcCompany.Connect()
if($code -eq 1)
{
$rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
#Data loading from a csv file
$csvItems = Import-Csv -Delimiter ';' -Path "C:\BOM_Header.csv"
 foreach($csvItem in $csvItems)
 {
    #Creating BOM object
    $bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"BillOfMaterial")
    #Checking that the BOM already exist
    $retValue = $bom.GetByItemCodeAndRevision($csvItem.BOM_Header, $csvItem.Revision)
   if($retValue -eq 0)
   {
    
    #Data loading from a csv file - Routing
    [array]$bomBindings = Import-Csv -Delimiter ';' -Path "C:\BOM_Routings_Operations_Bind.csv" | Where-Object {$_.BOM_Header -eq $csvItem.BOM_Header -and $_.Revision -eq $csvItem.Revision}
    if($bomBindings.count -gt 0)
    {
    
        
       
        #Deleting all existing routings, operations, resources
         
        $count = $bom.RoutingsOperationInputOutput.Count;
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $bom.RoutingsOperationInputOutput.DelRowAtPos(0);
        }
            
        $bom.RoutingsOperationInputOutput.SetCurrentLine($bom.RoutingsOperationInputOutput.Count-1);
         
        #Adding a new data - Bind
        foreach($bb in $bomBindings)
        { 
            $errorFlag = 0;


            $rs.DoQuery([string]::Format("SELECT RO.""U_RtgOprCode"" FROM ""@CT_PF_BOM12"" RO
            WHERE RO.""U_RtgCode"" =  N'{0}' AND RO.""U_OprCode"" =  N'{1}' AND RO.""U_OprSequence"" =  N'{2}' 
            AND RO.""U_BomCode"" = N'{3}' AND RO.""U_RevCode"" = N'{4}'
            ",$bb.RoutingCode, $bb.OperationCode, $bb.OperationSequence, $bb.BOM_Header, $bb.Revision));

            if($rs.RecordCount -gt 0)
            {
                
                $bom.RoutingsOperationInputOutput.U_RtgOprCode = $rs.Fields.Item(0).Value
            }
            else
            {
                $errorFlag = 1;
                 [System.String]::Format("Error adding binding Routing: {0}, Operation: {1}, OperationSequence: {2}
                 - RECORD NOT FOUND", $bb.RoutingCode, $bb.OperationCode, $bb.OperationSequence);
            }


            $rs.DoQuery([string]::Format("SELECT ISNULL(BS4.""U_LineNum"",ISNULL(BS3.""U_LineNum"",BS.""U_LineNum"")) 
            FROM ""@CT_PF_OBOM"" B LEFT OUTER JOIN ""@CT_PF_BOM1"" BS ON B.""Code"" = BS.""Code"" AND 'IT' = N'{3}' AND BS.""U_ItemCode"" =  N'{2}' AND BS.""U_Sequence"" =  N'{4}'
            LEFT OUTER JOIN ""@CT_PF_BOM3"" BS3 ON B.""Code"" = BS3.""Code"" AND 'CP' = N'{3}' AND BS3.""U_ItemCode"" =  N'{2}' AND BS3.""U_Sequence"" =  N'{4}'
            LEFT OUTER JOIN ""@CT_PF_BOM4"" BS4 ON B.""Code"" = BS4.""Code"" AND 'SC' = N'{3}' AND BS4.""U_ItemCode"" =  N'{2}' AND BS4.""U_Sequence"" =  N'{4}'
            WHERE B.""U_ItemCode"" =  N'{0}' AND B.""U_Revision"" =  N'{1}' AND ISNULL(BS4.""U_LineNum"",ISNULL(BS3.""U_LineNum"",ISNULL(BS.""U_LineNum"",-1))) != -1 
            ",$bb.BOM_Header, $bb.Revision, $bb.ItemCode, $bb.ItemType ,$bb.ItemSequence));

            

            $bom.RoutingsOperationInputOutput.U_RtgCode = $bb.RoutingCode
            $bom.RoutingsOperationInputOutput.U_OprCode = $bb.OperationCode
            $bom.RoutingsOperationInputOutput.U_ItemCode = $bb.ItemCode
            $bom.RoutingsOperationInputOutput.U_Direction = $bb.Direction

            if($bb.ItemType -eq 'IT')
            {
                $bom.RoutingsOperationInputOutput.U_ItemType = [CompuTec.ProcessForce.API.Enumerators.ManufacturingComponentType]::Item
            }
            if($bb.ItemType -eq 'CP')
            {
                $bom.RoutingsOperationInputOutput.U_ItemType = [CompuTec.ProcessForce.API.Enumerators.ManufacturingComponentType]::Cooproduct
            }
            if($bb.ItemType -eq 'SC')
            {
                $bom.RoutingsOperationInputOutput.U_ItemType = [CompuTec.ProcessForce.API.Enumerators.ManufacturingComponentType]::Scrap
            }


            if($bb.TimeCalc -eq 'Y')
            {
                $bom.RoutingsOperationInputOutput.U_InTimeCalc = 'Y'
            }
            else
            {
                $bom.RoutingsOperationInputOutput.U_InTimeCalc = 'N'
            }

            if($rs.RecordCount -gt 0)
            {
                
                $bom.RoutingsOperationInputOutput.U_BaseLine = $rs.Fields.Item(0).Value
            }
            else
            {
                $errorFlag = 1;
                 [System.String]::Format("Error adding binding BOM: {0}, Revision: {1}, ItemCode: {2},
                 IemType: {3}, Sequence: {4} - RECORD NOT FOUND", $bb.BOM_Header, $bb.Revision, $bb.ItemCode, $bb.ItemType ,$bb.ItemSequence);
            }

           
            if($errorFlag -eq 0)
            {
                $bom.RoutingsOperationInputOutput.Add();
            }
        }
        

        $message = 0

        [System.String]::Format("Updating BOM: {0}, Revision: {1}", $csvItem.BOM_Header, $csvItem.Revision)
        $message = $bom.Update()
       
     
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
  }
}
else
{
write-host "Failure"
}