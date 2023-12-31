#region #Script info
########################################################################
# CompuTec PowerShell Script - Removing Item Costing Details
########################################################################
# Version: 2.0
# Last tested PF version: PF 9.1 PL13
# Description: 
# 	Removes Item Costing Data per key: item + revision + cost category included in CSV file
# 	and runs Item Costing Restore after removing data.
# Warning:
# 	All manual data which was set in Item Costing will be removed, and only data structure will be recreated.
# 	It's recommended run script when all users all disconnected.
#   Before running this script please Make Backup of your database.
# Troubleshooting: 
#   https://connect.computec.pl/display/PF930EN/PowerShell+FAQ
########################################################################
#endregion

#region #PF API library usage
clear
# You need to check in what architecture PowerShell ISE is running (x64 or x86), 
# you need run ISE in the same architecture like PF API is installed (check in Windows -> Programs & Features)
# Examples: 
# 	SAP Client + PF x64 installed on DB/Company => PF API x64 => Windows PowerShell ISE
# 	SAP Client + PF x86 installed on DB/Company => PF API x86 => Windows PowerShell ISE x86
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")
#endregion

#region #Datbase/Company connection settings

$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany()
$pfcCompany.LicenseServer = "10.0.0.xx:40000"
$pfcCompany.SQLServer = "10.0.0.xx:30015"
$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::"dst_HANADB"
$pfcCompany.Databasename = "MICHALB_PFDEMOGB"
$pfcCompany.UserName = "michalb"
$pfcCompany.Password = "1234"

# where:

# LicenseServer = Server name or IP Address with port number, should be the same like in SLD ( see https://[SLD server]:40000/ControlCenter/ )
# SQLServer 	= Server name or IP Address with port number, sometimes best is use IP Address for resolve connection problems
#
# DbServerType 	= 	[SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2012" 	# For MsSQL Server 2012
#                	[SAPbobsCOM.BoDataServerTypes]::"dst_MSSQL2014" 	# For MsSQL Server 2014
#                	[SAPbobsCOM.BoDataServerTypes]::"dst_HANADB" 		# For HANA
#
# Databasename 	= Database / schema name (check in SAP Company select form/window, or in MsSQL Management Studio or in HANA Studio)
# UserName 		= SAP user name ex. manager
# Password 		= SAP user password

#endregion

#region #Connect to company

write-host -backgroundcolor yellow -foregroundcolor black  "Trying connect..."
$version = [CompuTec.Core.CoreConfiguration+DatabaseSetup]::AddonVersion
write-host -backgroundcolor green -foregroundcolor black "PF API Library:" $version';' 'Host:'(Get-WmiObject Win32_OperatingSystem).CSName';' 'OSArchitecture:' (Get-WmiObject Win32_OperatingSystem).OSArchitecture

try
{
    $code = $pfcCompany.Connect()

    write-host -backgroundcolor green -foregroundcolor black "Connected to:" $pfcCompany.SapCompany.CompanyName "/ " $pfcCompany.SapCompany.CompanyDB"" "Sap Company version: " $pfcCompany.SapCompany.Version
}
catch
{
	#Show error messages & stop the script
     write-host "Connection Failure: " -backgroundcolor red -foregroundcolor white $_.Exception.Message

     write-host "LicenseServer:" $pfcCompany.LicenseServer
     write-host "SQLServer:" $pfcCompany.SQLServer
     write-host "DbServerType:" $pfcCompany.DbServerType
     write-host "Databasename" $pfcCompany.Databasename
     write-host "UserName:" $pfcCompany.UserName
   
     return
}

#If company is not connected - stops the script
if(-not $pfcCompany.IsConnected)
{
     write-host "Company is not connected"
     return
}

#endregion

try
{

	#Data loading from a csv file (you need put correct path to CSV file)
	$csvItemCostings = Import-Csv -Delimiter ';' -Path "C:\PS\PF\Costing\ItemCosting.csv"
	
	 foreach($csvItemCosting in $csvItemCostings) 
	 {
			try
			{
				#Creating Item Costing Object
				$ic = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]"ItemCosting")
		
				#Checking if ItemCosting exists
				$retValue = $ic.Get($csvItemCosting.ItemCode, $csvItemCosting.Revision, $csvItemCosting.CostCategory)
		
				if($retValue)
				{
	   
					$message = 0
		
		
					[System.String]::Format("Removing Item Costing Details for Item: {0} Revision: {1} Category: {2}", $csvItemCosting.ItemCode, $csvItemCosting.Revision, $csvItemCosting.CostCategory)
					$message = $ic.Delete()
					 
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
		  catch
		  {
			write-host -backgroundcolor red -foregroundcolor white "Item Costing Details for Item: "  $csvItemCosting.ItemCode   " Revision: "  $csvItemCosting.Revision  " Category: " $csvItemCosting.CostCategory  " don't exists";
			write-host "Error: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
			write-host "Failure: " $pfcCompany.GetLastErrorDescription()			
		  }
	 }
	 }catch
	 {
			write-host "Error: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
			write-host "Failure: " $pfcCompany.GetLastErrorDescription()
	 }

    try
    {
       #region Restore Item Costing Data
       write-host "Restoring Item Costing Data"
       $restore = New-Object CompuTec.ProcessForce.API.Documents.Costing.Restoration.ItemCostingRestorer($pfcCompany.Token)
       $result = $restore.Restore();

       if($result -eq 1)
       {
            write-host "Success"
       }else 
       {
            write-host "Failure: " $pfcCompany.GetLastErrorDescription()
       }
       #endregion

    }catch
    {
      write-host "Error: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
      write-host "Failure: " $pfcCompany.GetLastErrorDescription()
    }

#region Close connection
if($pfcCompany.IsConnected)
{
    $pfcCompany.Disconnect()
    
    write-host " "
    write-host  –backgroundcolor green –foregroundcolor black "Disconnected from the company"
}

#endregion


