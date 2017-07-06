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
    Write-Host 'Preparing data: '
    [array]$csvItemsRoutings = Import-Csv -Delimiter ';' -Path "C:\MP_PS\SPROC\ItemDetails_Revisions.csv"
    $dictionaryItemsRoutings = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[array]]'

    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
    if($csvItemsRoutings.Count -gt 1) {
        $total = $csvItemsRoutings.Count
    } else {
        $total = 1
    }

    foreach($row in $csvItemsRoutings){
        $key = $row.Itemcode
        $progressItterator++;
        $progres = [math]::Round(($progressItterator * 100) / $total);
        if($progres -gt $beforeProgress)
        {
            Write-Host $progres"% " -NoNewline
            $beforeProgress = $progres
        }
        

        if($dictionaryItemsRoutings.ContainsKey($key)){
            $list = $dictionaryItemsRoutings[$key];
        } else {
            $list = New-Object System.Collections.Generic.List[array];
            $dictionaryItemsRoutings[$key] = $list;
        }
    
        $list.Add([array]$row);
    }
    Write-Host '';
    Write-Host 'Add/Update data to SAP: '
    $progressItterator = 0;
    $progres = 0;
    $beforeProgress = 0;
    if($dictionaryItemsRoutings.Count -gt 1) {
        $total = $dictionaryItemsRoutings.Count
    } else {
        $total = 1
    }

    #Checking that Item Details already exist
    foreach($key in $dictionaryItemsRoutings.Keys)
    {
        try {
            $progressItterator++;
            $progres = [math]::Round(($progressItterator * 100) / $total);
            if($progres -gt $beforeProgress)
            {
                Write-Host $progres"% " -NoNewline
                $beforeProgress = $progres
            }
            $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]::"BoRecordset")
      
            $dummy = $rs.DoQuery([string]::Format( "SELECT T0.""ItemCode"" AS ""ItemCode"", T1.""U_ItemCode"" AS ""IDT_ItemCode"" FROM OITM T0
                LEFT OUTER JOIN ""@CT_PF_OIDT"" T1 ON T0.""ItemCode"" = T1.""U_ItemCode""
                WHERE
                T0.""ItemCode"" = N'{0}'", $key))
            $exists = 0;
            if($rs.RecordCount -gt 0)
            {
                if($rs.Fields.Item('IDT_ItemCode').Value -eq $key) {
                    $exists = 1
                }


            } else {
                $err = [string]::Format('Item Master Data with ItemCode {0} do not exists.',$key);
                throw $err
            }
   
            #Creating Item Details
            $idt = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::"ItemDetails")
      
            #Restoring Item Costs and setting Inherit Batch/Serial to 'Yes'
            if($exists -eq 1)
            {
                $dummy = $idt.GetByItemCode($key)
            }
            else
            {
                $idt.U_ItemCode = $key;
                $idt.CFG_RestoreItemCosting = "Y";
            }
      
            $idt.U_InheritBatch = 1;
            $idt.U_InheritSerial = 1;
    
            [array]$revisions = $dictionaryItemsRoutings[$key];
            if($revisions.count -gt 0)
            {
                #Deleting all exisitng Revisions
                $count = $idt.Revisions.Count
                for($i=0; $i -lt $count; $i++)
                {
                    $dummy = $idt.Revisions.DelRowAtPos(0);
                }
                $idt.Revisions.SetCurrentLine($idt.Revisions.Count - 1);
          
                #Adding Revisions
                foreach($rev in $revisions)
                {
                    $idt.Revisions.U_Code = $rev.RevisionCode
                    $idt.Revisions.U_Description = $rev.RevisionName
                    $idt.Revisions.U_Status = $rev.Status #enum type; Revision Status, Active ACT = 1, BeingPhasedOut BPO = 2, Engineering ENG = 3, Obsolete OBS = 4
                    if($rev.ValidFrom -gt ''){
                        $idt.Revisions.U_ValidFrom = $rev.ValidFrom
                    }
                    if($rev.ValidTo -gt '') {
                       $idt.Revisions.U_ValidTo = $rev.ValidTo
                    }
                    $idt.Revisions.U_Remarks = $rev.Remarks
                    $idt.Revisions.U_Default = $rev.IsDefault #enum type; 1 = Yes, 2 = No
                    $idt.Revisions.U_IsMRPDefault = $rev.IsMRPDefault #enum type; 1 = Yes, 2 = No
                    $idt.Revisions.U_IsCostingDefault = $rev.DefaultForCosting #enum type; 1 = Yes, 2 = No
              
                    $dummy = $idt.Revisions.Add()
                }
            }
   
            $message = 0
      
            #Adding or updating ItemDetails depends on exists in the database
            if($exists -eq 1)
            {
                #[System.String]::Format("Updating Item Details: {0}", $key)
                $message = $idt.Update()
            }
            else
            {
                #[System.String]::Format("Adding Item Details: {0}", $key)
               $message= $idt.Add()
            }
      
            if($message -lt 0)
            {  
                $err=$pfcCompany.GetLastErrorDescription()
                if($exists -eq 1)
                {
                    $msg = [string]::Format("Error when updating Item Details: {0}. Details: {1}", $key,$err)
                } else {
                   $msg = [string]::Format("Error when adding Item Details: {0}. Details: {1}", $key, $err)
                }
        
                write-host -backgroundcolor red -foregroundcolor white $err
            }
        } Catch {
            $err=$_.Exception.Message;
            $ms = [string]::Format("Error when adding/updating Item Details for ItemCode {0} Details: {1}",$key,$err);
            Write-Host -BackgroundColor DarkRed -ForegroundColor White $ms
            if($pfcCompany.InTransaction){
                $pfcCompany.EndTransaction([CompuTec.ProcessForce.API.StopTransactionType]::Rollback);
            } 
        }
    }
}
else
{
write-host "Connection Failure"
}