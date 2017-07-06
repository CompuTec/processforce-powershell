#### DI API path ####
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")


#--------------Change only this section----------------------

#Connection data
$USERNAME = "manager"
$PASSWORD = "1234"
$SQL_SERVER = "localhost"
$SQL_USERNAME = "sa"
$SQL_PASSWORD = "saPass"
$DATABASE_NAME = "PFDemo"
$SERVER_TYPE = "dst_MSSQL2012"

#New Resource Code
$RESOURCE = 'RepTest'

#Sql query that need to return columns in following order: Routing Code, Routing Operation Code, IsDefault, QueTime, QueTimeUom, SetupTime, SetupTimeUoM, RunTime, RunTimeUoM, StockTime, StockTimeUoM
$SQLSTRING = [string]::Format("SELECT ""U_RtgCode"" AS ""Routing Code"" , ""U_RtgOprCode"" AS ""Routing Operation Code"", 'N' AS ""IsDefault"", 1 AS ""QueTime"", 1 AS ""QueTimeUom"",
                         1 AS ""SetupTime"", 1 AS ""SetupTimeUoM"", 1 AS ""RunTime"", 1 AS ""RunTimeUoM"", 1 AS ""StockTime"", 1 AS ""StockTimeUoM""
                        from ""@CT_PF_RTG1"" WHERE ""U_OprCode"" = '01'");

#--------------Change only this section----------------------


#Database connection
$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.UserName = $USERNAME
$pfcCompany.Password = $PASSWORD
$pfcCompany.SQLServer = $SQL_SERVER
$pfcCompany.SQLUserName = $SQL_USERNAME
$pfcCompany.SQLPassword = $SQL_PASSWORD
$pfcCompany.Databasename = $DATABASE_NAME
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::$SERVER_TYPE
        
$code = $pfcCompany.Connect()
if($pfcCompany.IsConnected -eq 1)
{


    $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset");
    $rs.DoQuery($SQLSTRING);
    

    ##-------------------------Adding resources to Routing------------------------------
    

    	
	while(!$rs.EoF)
	{
        $routing = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Routing)
        $routing.GetByRtgCode($rs.Fields.Item(0).Value);
        $routing.OperationResources.U_RtgOprCode = $rs.Fields.Item(1).Value
        $routing.OperationResources.U_RscCode = $RESOURCE
        $routing.OperationResources.U_IsDefault = $rs.Fields.Item(2).Value
        $routing.OperationResources.U_QueueTime = $rs.Fields.Item(3).Value
        $routing.OperationResources.U_QueueRate = $rs.Fields.Item(4).Value #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
        $routing.OperationResources.U_SetupTime = $rs.Fields.Item(5).Value
        $routing.OperationResources.U_SetupRate = $rs.Fields.Item(6).Value #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
        $routing.OperationResources.U_RunTime = $rs.Fields.Item(7).Value
        $routing.OperationResources.U_RunRate = $rs.Fields.Item(8).Value #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
        $routing.OperationResources.U_StockTime = $rs.Fields.Item(9).Value
        $routing.OperationResources.U_StockRate = $rs.Fields.Item(10).Value #enum RateType; FixedSeconds = 1, FixedMinutes = 2, FixedHours = 3, SecondsPerPiece = 4, MinutesPerPiece = 5, HoursPerPiece = 6, PiecesPerSecond = 7, PiecesPerMinute = 8,PiecesPerHour = 9
        $routing.OperationResources.Add()        
		

	    
        
        [System.String]::Format("Adding resource for Routing: {0}, Operation Routing Code: {1}", $rs.Fields.Item(0).Value, $rs.Fields.Item(1).Value)
        $rs.MoveNext();
        $message = 0
        $message = $routing.Update();
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



	};
      
    
    $pfcCompany.Disconnect();	  
}
    
else
{
        write-host "Connection failure. Check log file for more details"
}