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
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2012
        
$code = $pfcCompany.Connect();
if($code -eq 1)
{
#Data loading from a csv file - Header information for Quality Templates
$csvTemplates = Import-Csv -Delimiter ';' -Path "C:\Quality_TemplatesForTestProtocols.csv"
 
#Checking that Template already exist 
 foreach($csvTemplate in $csvTemplates) 
 {
  
        $rs = $pfcCompany.CreateSapObject([SAPbobsCOM.BoObjectTypes]"BoRecordset")
     
        $rs.DoQuery([string]::Format( "SELECT ""U_TemplateCode"", ""Code"" FROM ""@CT_PF_OTPT"" WHERE ""U_TemplateCode"" = N'{0}'", $csvTemplate.TemplateCode))
        $exists = 0;
        if($rs.RecordCount -gt 0)
        {
            $exists = 1
        }
  
    #Creating Template
    $tmpl = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"TestProtocolTemplate")
    $rs.MoveFirst();
    
    if($exists -eq 1)
    {
		$tmpl.getByKey($rs.Fields.Item('Code').Value);
    }
    else
    {
        $tmpl.U_TemplateCode = $csvTemplate.TemplateCode;
		$tmpl.U_TemplateName = $csvTemplate.TemplateName;
		if($csvTemplate.ValidFrom -ne "")
		{
			$tmpl.U_ValidFrom = $csvTemplate.ValidFrom;
		}
        else
		{
			$tmpl.U_ValidFrom = [DateTime]::MinValue 
		}
		if($csvTemplate.ValidTo -ne "")
		{
			$tmpl.U_ValidTo = $csvTemplate.ValidTo;
		}
        else
		{
			$tmpl.U_ValidTo = [DateTime]::MinValue 
		}
		if($csvTemplate.GroupCode -ne "")
		{
			$tmpl.U_GrpCode = $csvTemplate.GroupCode;
		}
		if($csvTemplate.Remarks -ne "")
		{
			$tmpl.U_Remarks = $csvTemplate.Remarks;
		}
	}
     #Data loading from the csv file - Rows for templates from Quality_TemplatesForTestProtocolsProperties.csv file
    [array]$Properties = Import-Csv -Delimiter ';' -Path "C:\Quality_TemplatesForTestProtocolsProperties.csv" | Where-Object {$_.TemplateCode -eq $csvTemplate.TemplateCode}
    if($Properties.count -gt 0)
    {
        #Deleting all exisitng Phrases
        $count = $tmpl.Properties.Count
        for($i=0; $i -lt $count; $i++)
        {
            $dummy = $tmpl.Properties.DelRowAtPos(0);
        }
        $tmpl.Properties.SetCurrentLine($tmpl.Properties.Count - 1);
         
        #Adding Properties
        foreach($prop in $Properties) 
        {
			$tmpl.Properties.U_PrpCode = $prop.PropertyCode;
			$tmpl.Properties.U_Expression = $prop.Expression;
			
			if($prop.RangeFrom -ne "")
			{
				$tmpl.Properties.U_RangeValueFrom = $prop.RangeFrom;
			}
			else
			{
				$tmpl.Properties.U_RangeValueFrom = 0;
			}
			$tmpl.Properties.U_RangeValueTo = $prop.RangeTo;
			
			if($prop.UoM -ne "")
			{
				$tmpl.Properties.U_UnitOfMeasure = $prop.UoM;
			}
			
			if($prop.ReferenceCode -ne "")
			{
				$tmpl.Properties.U_RefCode = $prop.ReferenceCode;
			}
			
			if($prop.ValidFrom -ne "")
			{
				$tmpl.Properties.U_ValidFromDate = $prop.ValidFrom;
			}
            else
            {
	            $tmpl.Properties.U_ValidFromDate = [DateTime]::MinValue 
	        }
			if($prop.ValidTo -ne "")
			{
				$tmpl.Properties.U_ValidToDate = $prop.ValidTo;
			}
            else
            {
	            $tmpl.Properties.U_ValidTDate = [DateTime]::MinValue 
	        }
			
            $dummy = $tmpl.Properties.Add()
        }
    }
  
    $message = 0
     
    #Adding or updating Template depends on exists in the database
    if($exists -eq 1)
    {
        [System.String]::Format("Updating Item Details: {0}", $csvTemplate.TemplateCode)
        $message = $tmpl.Update()
    }
    else
    {
        [System.String]::Format("Adding Item Details: {0}", $csvTemplate.TemplateCode)
       $message= $tmpl.Add()
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