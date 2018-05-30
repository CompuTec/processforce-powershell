clear
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")

#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = "manager"
$pfcCompany.Password = "1234"
$pfcCompany.LicenseServer="localhost:30000"
#$pfcCompany.SQLPassword = "sa"
$pfcCompany.SQLServer = "PLPO017"
#$pfcCompany.SQLUserName = "sa"
$pfcCompany.Databasename = "MG"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2014
#[SAPbobsCOM.BoDataServerTypes]"dst_MSSQL2012"
     $SQLQUERY= "select t0.DocEntry,t0.DocNum, t0.ObjType,t1.ItemCode,t1.U_Revision ,t2.U_TestPrclCode , t4.MdAbsEntry , t6.""U_DistNumber"" ,'PG' TransType , t0.""CardCode"" from OPDN t0 
inner join PDN1 t1 on t0.DocEntry=t1.""DocEntry""
inner join [@CT_PF_OTCL] t2 on t2.""U_ItemCode"" = t1.ItemCode and ( t0.CardCode=t2.U_TrsPurGdsRcptPoBp or isnull(t2.U_TrsPurGdsRcptPoBp,'')=''  )
inner join OITL t3 on t0.ObjType=t3.DocType and t0.""DocEntry"" = t3.DocEntry and t3.DocLine=t1.LineNum
left outer join ITL1 t4 on t4.""LogEntry""=t3.LogEntry 
left outer join ""@CT_PF_OABT"" t6 on t6.""U_ItemCode""=t1.""ItemCode"" and t6.U_SysNumber=t4.SysNumber
left outer join  [@CT_PF_OQCT] t5 on t5.""U_TransType""='PG' and t0.DocEntry=t5.U_MnfOrder

where t2.U_TrsPurGdsRcptPo ='Y' and t5.""DocEntry"" is null  "
$code = $pfcCompany.Connect();
if($code -eq 1)
{
    #Data loading from a Query 
 $queryManager = New-Object CompuTec.Core.DI.Database.QueryManager
		$queryManager.CommandText = $SQLQUERY
		$rs = $queryManager.Execute($pfcCompany.Token);
        $ppPositionsCount = $rs.RecordCount
        if ($ppPositionsCount -gt 0) {
            $msg = [string]::format('positions to add: {0}', $ppPositionsCount)
            Write-Host -BackgroundColor Blue $msg
           
            $prevDocEntry = -1;
        
            #[docEntry][U_LineNum]
            $MainList = New-Object 'System.Collections.Generic.List[psobject]'
            #[docEntry][U_LineNum][object-line]
           
            Write-Host 'Przygotowywanie danych o pozycjach'
            while (!$rs.EoF) {



                $TestProtocolCode = $rs.Fields.Item('U_TestPrclCode').Value;
                $TransactionID = $rs.Fields.Item('DocEntry').Value;
                $TransactionType    =$rs.Fields.Item('TransType').Value;
                $CardCode    =$rs.Fields.Item('CardCode').Value;
                $DistNumber    =$rs.Fields.Item('U_DistNumber').Value;
                
                if($TransactionID -eq $prevDocEntry)
                {
                    $qcTest.DistNumber.Add($DistNumber)
                }
                else
                {
                $qcTest =[psobject]@{
                        TestProtocolCode = $TestProtocolCode
                        TransactionID = $TransactionID
                        TransactionType = $TransactionType
                        CardCode = $CardCode
                        DistNumber =New-Object 'System.Collections.Generic.List[string]'
                    
                    }
                    $qcTest.DistNumber.Add($DistNumber)

                 $MainList.Add( $qcTest);
                }
                 $prevDocEntry=$TransactionID
                 
                $rs.MoveNext();
            }
            }
#Checking that Control Test already exist 
 foreach($csvTest in $MainList)
 {
  
    #Creating ControlTest
    $test = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"QualityControlTest")
    
	$test.U_TestProtocolNo = $csvTest.TestProtocolCode;
	 
	$test.U_Status  ;
	 
	$test.U_TransType = $csvTest.TransactionType;
	$test.U_BpCode = $csvTest.CardCode;
	$test.U_MnfOrder = $csvTest.TransactionID;
	 
	
	#Batches
	 #Data loading from the csv file - Batches for Test from Quality_ControlTestTransBatches.csv file
	 #Checks if the file exists
	 
	    if($csvTest.DistNumber.Count -gt 0)
	    {
	        #Deleting all exisitng Items
	       
	        $test.Batches.SetCurrentLine($test.Batches.Count - 1);
	         
		    #Adding Defects
	        foreach($batch in $csvTest.DistNumber)
	        {
				$test.Batches.U_Batch = $batch ;
	            $dummy = $test.Batches.Add();
	        }
	    }
  	
	
	

	
    $message = 0
     
    #Adding or updating Test depends on exists in the database
    
    write-host [System.String]::Format("Adding Item Details: {0}", $csvTest.Key);
    $message = $test.Add()
    
     
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