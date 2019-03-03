Clear-Host
#path to proper WMS server installation, please be sure that you use same powershell atchitecture as wms server
[System.Reflection.Assembly]::LoadFrom("c:\Program Files\CompuTec\CompuTec WMS Server\CompuTec.Base.dll");
[System.Reflection.Assembly]::LoadFrom("c:\Program Files\CompuTec\CompuTec WMS Server\CompuTec.Connection.dll");
[System.Reflection.Assembly]::LoadFrom("c:\Program Files\CompuTec\CompuTec WMS Server\CompuTec.Server.Tools.dll");

#add username and pass for SAP user in next line as arguments
$connection = [CompuTec.Connection.ConnectionFactory]::CreateInstance(3.2026).CreateConnection("maciejw", "1234");

#$connection = New-Object CompuTec.Connection.SAPConnection("maciejw", "1234");


if ($connection.Company.Connected)
{
	Write-Host "connected"
}
$sManager = New-Object CompuTec.Base.SU.SUManager($connection);
$irManager = New-Object CompuTec.Base.SU.IRManager($connection);

#proper path to csv file
$csvSUHeader = Import-Csv -Delimiter ';' -Path "d:\Projects\test projects\importSU\SUHeader.csv"
$csvSULines = Import-Csv -Delimiter ';' -Path "d:\Projects\test projects\importSU\SULines.csv"

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
	$irLines = $sManager.SaveSU($su, $true)
	$ircode = $irManager.SaveInventoryRegister($irLines, -1, -1, $TRUE)
	
	$dummy = 0;
}




$dummy = 0