Clear-Host
#path to proper WMS server installation, please be sure that you use same powershell atchitecture as wms server
[System.Reflection.Assembly]::LoadFrom("c:\Program Files\CompuTec\CompuTec WMS Server\Dlls\CompuTec.Base.dll");
[System.Reflection.Assembly]::LoadFrom("c:\Program Files\CompuTec\CompuTec WMS Server\Dlls\CompuTec.Core2.dll");
[System.Reflection.Assembly]::LoadFrom("c:\Program Files\CompuTec\CompuTec WMS Server\Dlls\CompuTec.Connection.dll");
[System.Reflection.Assembly]::LoadFrom("c:\Program Files\CompuTec\CompuTec WMS Server\Dlls\CompuTec.Server.Tools.dll");
[System.Reflection.Assembly]::LoadFrom("c:\Program Files\CompuTec\CompuTec WMS Server\Dlls\CompuTec.Configuration.dll");
[System.Reflection.Assembly]::LoadFrom("c:\Program Files\CompuTec\CompuTec WMS Server\OpenNETCF.IoC.dll");
[System.Reflection.Assembly]::LoadFrom("c:\net45\Newtonsoft.Json.dll");
[System.Reflection.Assembly]::LoadFrom("c:\Program Files\CompuTec\CompuTec WMS Server\Dlls\CompuTec.WMS.API.dll");

#add username and pass for SAP user in next line as arguments
$wmsConfig = [CompuTec.ConfigurationWMS.WMSConfiguration]::Manager;
$wmsConfig.LicenseServerSAP = "hanadev:40000";
$wmsConfig.DbServerType = 9; #HANA - 9


$connection = New-Object CompuTec.Connection.WMSConnection([CompuTec.WMS.API.BusinessObjects.WMSBaseInfo]::Manager, "manager", "1234", "DEV@hanadev:30013", "PFDEMOGB", "1", $true);
[CompuTec.Connection.CoreInitializer]::Initialize($connection.Token);

if ($connection.Company.Connected)
{
    Write-Host "---------------"
	Write-Host "connected"
    Write-Host "---------------"
}

$sManager = New-Object CompuTec.Base.SU.SUManager($connection);
$irManager = New-Object CompuTec.Base.SU.IRManager($connection);

#proper path to csv file
$csvSUHeader = Import-Csv -Delimiter ';' -Path "C:\CompuTec-Workfolder-DO NOT DELETE\powershell-ImportSU\SUHeader.csv"
$csvSULines = Import-Csv -Delimiter ';' -Path "C:\CompuTec-Workfolder-DO NOT DELETE\powershell-ImportSU\SULines.csv"

$bindict = @{}
$suTypes = @{}

$su = $null

foreach($csvItem in $csvSUHeader) 
{
	$su = $sManager.NewSU($false); 
	$su["SUCode"] = $csvItem.U_Code;
	$su["WhsCode"] = $csvItem.U_WhsCode;
	$su["BinCode"] = $csvItem.U_BinCode;
	$su["BinAbs"] = $csvItem.U_BinAbs;
	$su["Status"] = $csvItem.U_Status;
	$su["GrossWgt"] = $csvItem.U_GrossWgt;
	$su["NetWgt"] = $csvItem.U_NetWgt;
	$su["Type"] = $csvItem.U_Type;
	$su["Parent"] = $csvItem.U_Parent;
	$su["SSCC"] = $csvItem.U_SSCC;
	$su["Attr1"] = $csvItem.U_Attribute1;
	$su["Attr2"] = $csvItem.U_Attribute2;
	$su["Attr3"] = $csvItem.U_Attribute3;
	$su["Attr4"] = $csvItem.U_Attribute4;
	$su["Attr5"] = $csvItem.U_Attribute5;
	$su["Attr6"] = $csvItem.U_Attribute6;
	$su["Attr7"] = $csvItem.U_Attribute7;
	$su["Attr8"] = $csvItem.U_Attribute8;
	$su["Attr9"] = $csvItem.U_Attribute9;
	$su["Attr10"] = $csvItem.U_Attribute10;
	$su["Remarks"] = $csvItem.U_Remarks;
	$su["BPPackageNo"] = $csvItem.U_BPPackageNo;
	$su["CardCode"] = $csvItem.U_CardCode;
	
	$suLines = $csvSULines | Where {$_.Code -eq $csvItem.U_Code}
	
	foreach($suLine in $suLines) 
	{
		$itemCode = $suLine.U_ItemCode;
		
		$qty = [Double]::Parse($suLine.U_Quantity,[Globalization.CultureInfo]::InvariantCulture.NumberFormat );
		
		$itemManageType = [CompuTec.WebAPI.Tools.Extensions.ItemsExtensions]::GetManageType($connection, $suLine.U_ItemCode);
	
		[CompuTec.Base.Tools.BusinessItemTools]::AddLineToBusinessItem($su.Childs[0], 
				$suLine.U_ItemCode, 
				"", 
				$su["WhsCode"], 
				$su["WhsCode"], 
				"", 
				"", 
				"", 
				"IT", 
				-1, 
				-1, 
				-1, 
				-1, 
				$qty, 
				0, 
				"1", 
				$itemManageType);
				
	
		$subLine = New-Object CompuTec.WebApi.Model.BusinessItem.SubBusinessItemContainer
	
		if($itemManageType -eq "B")
		{
			$subLine["DistNumber"] = $suLine.U_DistNumber;
			$subLine["AbsEntry"] = [CompuTec.Base.Extensions.Warehouse.WarehouseExtensions]::BatchAbs($connection, $suLine.U_ItemCode, $suLine.U_DistNumber);
		}
		else 
		{
			if($itemManageType -eq "S")
			{
				$subLine["DistNumber"] = $suLine.U_DistNumber;
				$subLine["AbsEntry"] = [CompuTec.Base.Extensions.Warehouse.WarehouseExtensions]::SerialAbs($connection, $suLine.U_ItemCode, $suLine.U_DistNumber);
			}
			else
			{
				$subLine["DistNumber"] = "";
				$subLine["AbsEntry"] = -1;
			}
		
		}
		$subLine["Quantity"] = $qty
		$su.Childs[0].Properties[$su.Childs[0].Properties.Count - 1].SubLines.Add($subLine);
	}
    
    Write-Host "Save SU " $su["Code"];
	$irLines = $sManager.SaveSU($su)
	$ircode = $irManager.SaveInventoryRegister($irLines, -1, -1, $TRUE)
	
	$dummy = 0;
}

Write-Host  
Write-Host "---------------"
Write-Host "Finished" 
Write-Host "---------------"

$dummy = 0
