using module .\lib\ItemMasterData.psm1;
using module .\lib\BillOfMaterials.psm1;
using module .\lib\ProductionOrder.psm1;
using module .\lib\Result.psm1;
add-type -Path "C:\Projects\Playground\SAP\DLL\SAPHana\x64\Interop.SAPbobsCOM.dll"

#region Script parameters
$WHS_CODE_1 = "01"
$WHS_CODE_2 = "02"
$COD_ITEMCODE = "SyncTest_CoD";
$FOD_ITEMCODE = "SyncTest_FoD";
$PH_ITEMCODE = "SyncTest_PH";
$A_ITEMCODE = "SyncTest_A";
$B_ITEMCODE = "SyncTest_B";
$C_ITEMCODE = "SyncTest_C";
$D_ITEMCODE = "SyncTest_D";
$F_ITEMCODE = "SyncTest_F";
$H_ITEMCODE = "SyncTest_H";
$X1_ITEMCODE = "SyncTest_X1";
$X2_ITEMCODE = "SyncTest_X2";
$X3_ITEMCODE = "SyncTest_X3";
$X4_ITEMCODE = "SyncTest_X4";

#region prepare Item Master Data
[ItemMasterData] $CoD = [ItemMasterData]::getNewCoproductDummy($COD_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $FoD = [ItemMasterData]::getNewFinalDummy($FOD_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $PH = [ItemMasterData]::getNewPhantom($PH_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $A = [ItemMasterData]::getNewRegularItem($A_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $B = [ItemMasterData]::getNewRegularItem($B_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $C = [ItemMasterData]::getNewRegularItem($C_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $D = [ItemMasterData]::getNewRegularItem($D_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $F = [ItemMasterData]::getNewRegularItem($F_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $H = [ItemMasterData]::getNewRegularItem($H_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $X1 = [ItemMasterData]::getNewRegularItem($X1_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $X2 = [ItemMasterData]::getNewRegularItem($X2_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $X3 = [ItemMasterData]::getNewRegularItem($X3_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
[ItemMasterData] $X4 = [ItemMasterData]::getNewRegularItem($X4_ITEMCODE, $WHS_CODE_1, $WHS_CODE_2);
#endregion
	
#region prepare Bill Of Materials
[BillOfMaterials] $BOMFoD = New-Object 'BillOfMaterials'($FoD.ItemCode, $FoD.DefaultWarehouseCode, 1);
$BOMFoD.addLine($CoD.ItemCode, $CoD.DefaultWarehouseCode, 1);

[BillOfMaterials] $BOMA = New-Object 'BillOfMaterials'($A.ItemCode, $A.DefaultWarehouseCode, 1);
$BOMA.addLine($B.ItemCode, $B.DefaultWarehouseCode, 1);
$BOMA.addLine($C.ItemCode, $C.DefaultWarehouseCode, 1);
	
[BillOfMaterials] $BOMPH = New-Object 'BillOfMaterials'($PH.ItemCode, $PH.DefaultWarehouseCode, 1);
$BOMPH.addLine($X1.ItemCode, $X1.DefaultWarehouseCode, 1);
	
[BillOfMaterials] $BOMD = New-Object 'BillOfMaterials'($D.ItemCode, $D.DefaultWarehouseCode, 1);
$BOMD.addLine($PH.ItemCode, $PH.DefaultWarehouseCode, 1);
$BOMD.addLine($A.ItemCode, $A.DefaultWarehouseCode, 1);
$BOMD.addLine($F.ItemCode, $F.DefaultWarehouseCode, 1);
$BOMD.addLine($H.ItemCode, $H.DefaultWarehouseCode, 1);
#endregion

$TEST_RESULT = New-Object 'Result';
$csvImportCatalog = $PSScriptRoot + "\"
$TEMP_XML_FILE = -join ($csvImportCatalog, "Temp.xml");
# $csvBomFilePath = -join ($csvImportCatalog, "BOMs.csv")
# $csvBomItemsFilePath = -join ($csvImportCatalog, "BOM_Items.csv")
# $csvBomscrapsFilePath = -join ($csvImportCatalog, "BOM_Scraps.csv")
# $csvBomCoproductsFilePath = -join ($csvImportCatalog, "BOM_Coproducts.csv")

#endregion


#region #Datbase/Company connection settings
#configuration xml
$configurationXMLFilePath = -join ($csvImportCatalog, "configuration.xml");
if (!(Test-Path $configurationXMLFilePath -PathType Leaf)) {
	Write-Host -BackgroundColor Red ([string]::Format("File: {0} don't exists.", $configurationXMLFilePath));
	return;
}
[xml] $configurationXml = Get-Content -Encoding UTF8 $configurationXMLFilePath
$xmlConnection = $configurationXml.SelectSingleNode("/configuration/connection");

$connectionConfirmation = [string]::Format('You are connecting to Database: {0} on Server: {1} as User: {2}. Do you want to continue [y/n]?:', $xmlConnection.Database, $xmlConnection.SQLServer, $xmlConnection.UserName);
Write-Host $connectionConfirmation -backgroundcolor Yellow -foregroundcolor DarkBlue -NoNewline
$confirmation = Read-Host
if (($confirmation -ne 'y') -and ($confirmation -ne 'Y')) {
	return;
}

$sapCompany = new-Object -ComObject SAPbobsCOM.Company
$sapCompany.Server = $xmlConnection.SQLServer;
$sapCompany.LicenseServer = $xmlConnection.LicenseServer;
$sapCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::[string]$xmlConnection.DbServerType;
$sapCompany.CompanyDB = $xmlConnection.Database;
$sapCompany.UserName = $xmlConnection.UserName;
$sapCompany.Password = $xmlConnection.Password;
 
#endregion

#region #Connect to company
 
write-host -backgroundcolor yellow -foregroundcolor black  "Trying connect..."
 
try {
	[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'code')]
	$code = $sapCompany.Connect();
 
	write-host -backgroundcolor green -foregroundcolor black "Connected to:" $sapCompany.CompanyName "/ " $sapCompany.CompanyDB"" "Sap Company version: " $sapCompany.Version
	#If company is not connected - stops the script
	if (-not $sapCompany.Connected) {
		throw [System.Exception] ("Company is not connected");
	}
}
catch {
	#Show error messages & stop the script
	write-host "Connection Failure: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
	write-host "LicenseServer:" $sapCompany.LicenseServer
	write-host "SQLServer:" $sapCompany.Server
	write-host "DbServerType:" $sapCompany.DbServerType
	write-host "Databasename" $sapCompany.CompanyDB
	write-host "UserName:" $sapCompany.UserName
	return
}
#endregion

Enum TransactionTask {
	Add = 1;
	Update = 2;
}

Enum TransactionType {
	DI = 1;
	XML = 2;
}
$canWeChangeHeaderWarehouseWhenCreatingProductionOrder = "Can We Change Header Warehouse When Creating Production Order"
function convertYesNoToBool([SAPbobsCOM.BoYesNoEnum] $value) {
	if ($value -eq [SAPbobsCOM.BoYesNoEnum]::tYES) {
		return $true;
	} 
	return $false;
}

function convertBoolToYesNo([bool] $value) {
	if ($value) {
		return [SAPbobsCOM.BoYesNoEnum]::tYES;
	} 
	return [SAPbobsCOM.BoYesNoEnum]::tNO;
}
function prepareItem([ItemMasterData] $ItemMasterData) {
	try {
		$item = $sapCompany.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oItems);

		$exists = $item.GetByKey($ItemMasterData.ItemCode);
		if ($exists) {
			try {
				if ($ItemMasterData.InventoryItem -ne (convertYesNoToBool($item.InventoryItem))) {
					throw [System.Exception] ([string]::Format("Item setting 'Inventory Item' is set in SAP to {0} when it should be {1}", [string](convertYesNoToBool($item.InventoryItem)), [string]$ItemMasterData.InventoryItem));
				}
				if ($ItemMasterData.SalesItem -ne (convertYesNoToBool($item.SalesItem))) {
					throw [System.Exception] ([string]::Format("Item setting 'Sales Item' is set in SAP to {0} when it should be {1}", [string](convertYesNoToBool($item.SalesItem)), [string]$ItemMasterData.SalesItem));
				}
				if ($ItemMasterData.PurchaseItem -ne (convertYesNoToBool($item.PurchaseItem))) {
					throw [System.Exception] ([string]::Format("Item setting 'Purchase Item' is set in SAP to {0} when it should be {1}", [string](convertYesNoToBool($item.PurchaseItem)), [string]$ItemMasterData.PurchaseItem));
				}
				if ($ItemMasterData.PhantomItem -ne (convertYesNoToBool($item.IsPhantom))) {
					throw [System.Exception] ([string]::Format("Item setting 'Phantom Item' is set in SAP to {0} when it should be {1}", [string](convertYesNoToBool($item.IsPhantom)), [string]$ItemMasterData.PhantomItem));
				}
				if ($ItemMasterData.AssetItem -ne (convertYesNoToBool($item.AssetItem))) {
					throw [System.Exception] ([string]::Format("Item setting 'Fixed Asset Indicator' is set in SAP to {0} when it should be {1}", [string](convertYesNoToBool($item.AssetItem)), [string]$ItemMasterData.AssetItem));
				}
				if ($ItemMasterData.ManageByBatches -ne (convertYesNoToBool($item.ManageBatchNumbers))) {
					throw [System.Exception] ([string]::Format("Item setting 'Manage Batch Numbers' is set in SAP to {0} when it should be {1}", [string](convertYesNoToBool($item.ManageBatchNumbers)), [string]$ItemMasterData.ManageByBatches));
				}
				if ($ItemMasterData.ManageBySerialNumbers -ne (convertYesNoToBool($item.ManageSerialNumbers))) {
					throw [System.Exception] ([string]::Format("Item setting 'Manage By Serial Numbers' is set in SAP to {0} when it should be {1}", [string](convertYesNoToBool($item.ManageSerialNumbers)), [string]$ItemMasterData.ManageBySerialNumbers));
				}
				$sapItemStandardCost = $false;
				if ($item.CostAccountingMethod -eq [SAPbobsCOM.BoInventorySystem]::bis_Standard) {
					$sapItemStandardCost = $true
				}
				if ($ItemMasterData.StandardValuationMethod -ne $sapItemStandardCost) {
					if ($sapItemStandardCost) {
						throw [System.Exception] ([string]::Format("Item Valuation Method should be set to standard"));
					}
					else {
						throw [System.Exception] ([string]::Format("Item Valuation Method shouldn't be set to standard"));
					}
				}
				if ($ItemMasterData.DefaultWarehouseCode -ne $item.DefaultWarehouse) {
					throw [System.Exception] ([string]::Format("Default warehouse is set in SAP to {0} when it should be {1}", [string]$item.DefaultWarehouse, [string]$ItemMasterData.DefaultWarehouseCode));
				}
				if ($ItemMasterData.StandardValuationMethod) {
					for ($wi = 0; $wi -lt $item.WhsInfo.Count; $wi++) {
						$item.WhsInfo.SetCurrentLine($wi);
						$whs = $item.WhsInfo;
						if ($whs.WarehouseCode -eq $ItemMasterData.DefaultWarehouseCode) {
							if ($ItemMasterData.AvgStdPrice -ne $whs.StandardAveragePrice) {
								throw [System.Exception] ([string]::Format("Item Cost is set in SAP to {0} on Warehouse {1} when it should be {2}", [string]$whs.StandardAveragePrice, [string]$ItemMasterData.DefaultWarehouseCode, [string]$ItemMasterData.AvgStdPrice));
							}
						}
						if ($whs.WarehouseCode -eq $ItemMasterData.SecondWarehouseCode) {
							if ($ItemMasterData.AvgStdPrice -ne $whs.StandardAveragePrice) {
								throw [System.Exception] ([string]::Format("Item Cost is set in SAP to {0} on Warehouse {1} when it should be {2}", [string]$whs.StandardAveragePrice, [string]$ItemMasterData.SecondWarehouseCode, [string]$ItemMasterData.AvgStdPrice));
							}
						}
					}
				}
				#region check if required warehouses exists
				$firstWhsExists = $false;
				$secondWhsExists = $false;
				for ($wi = 0; $wi -lt $item.WhsInfo.Count; $wi++) {
					$item.WhsInfo.SetCurrentLine($wi);
					$whs = $item.WhsInfo;
					if ($whs.WarehouseCode -eq $ItemMasterData.DefaultWarehouseCode) {
						$firstWhsExists = $true;
					}
					if ($whs.WarehouseCode -eq $ItemMasterData.SecondWarehouseCode) {
						$secondWhsExists = $true;
					}
				}
				if (-not ($firstWhsExists -and $secondWhsExists)) {
					
					if (-not $firstWhsExists -and -not $secondWhsExists) {
						$missingWarehouses = [string]::Format("Missing warehouses: {0}, {1}.", [string] $ItemMasterData.DefaultWarehouseCode, [string]$ItemMasterData.SecondWarehouseCode);
					}
					elseif (-not $firstWhsExists) {
						$missingWarehouses = [string]::Format("Missing warehouse: {0}", [string] $ItemMasterData.DefaultWarehouseCode);
					}
					else {
						$missingWarehouses = [string]::Format("Missing warehouse: {0}", [string] $ItemMasterData.SecondWarehouseCode);
					}
					throw [System.Exception] ($missingWarehouses);
				}

				#endregion
			}
			catch {
				$err = [string]$_.Exception.Message;
				$msg = [string]::Format("Item already exists but it has wrong parameters. Details: {0}", $err);
				throw [System.Exception]($msg);
			}
		}
		else {
			try {
				$sapCompany.StartTransaction();
				$item.ItemCode = $ItemMasterData.ItemCode;
				$item.ItemName = $ItemMasterData.ItemCode;
				$item.InventoryItem = convertBoolToYesNo($ItemMasterData.InventoryItem);
				$item.SalesItem = convertBoolToYesNo($ItemMasterData.SalesItem);
				$item.PurchaseItem = convertBoolToYesNo($ItemMasterData.PurchaseItem);
				$item.IsPhantom = convertBoolToYesNo($ItemMasterData.PhantomItem);
				$item.AssetItem = convertBoolToYesNo($ItemMasterData.AssetItem);
				$item.ManageBatchNumbers = convertBoolToYesNo($ItemMasterData.ManageByBatches);
				$item.ManageSerialNumbers = convertBoolToYesNo($ItemMasterData.ManageBySerialNumbers);
				if ($ItemMasterData.StandardValuationMethod) {
					$item.CostAccountingMethod = [SAPbobsCOM.BoInventorySystem]::bis_Standard;
					if (-not $ItemMasterData.InventoryItem) {
						$item.AvgStdPrice = $ItemMasterData.AvgStdPrice;
					}
				}


				$firstWhsExists = $false;
				$secondWhsExists = $false;
				for ($wi = 0; $wi -lt $item.WhsInfo.Count; $wi++) {
					$item.WhsInfo.SetCurrentLine($wi);
					$whs = $item.WhsInfo;
					if ($whs.WarehouseCode -eq $ItemMasterData.DefaultWarehouseCode -or $whs.WarehouseCode -eq $ItemMasterData.SecondWarehouseCode) {
						if ($whs.WarehouseCode -eq $ItemMasterData.DefaultWarehouseCode) {
							$firstWhsExists = $true;
						}
						else {
							$secondWhsExists = $true;
						}
						if (-not $ItemMasterData.InventoryItem) {
							$whs.StandardAveragePrice = $ItemMasterData.AvgStdPrice;
						}
					}
				}
				if (-not $firstWhsExists) {
					$item.WhsInfo.SetCurrentLine($item.WhsInfo.Count - 1);
					if (-not [string]::IsNullOrWhiteSpace($item.WhsInfo.WarehouseCode)) {
						$item.WhsInfo.Add();
					}
					$item.WhsInfo.WarehouseCode = $ItemMasterData.DefaultWarehouseCode;
					if (-not $ItemMasterData.InventoryItem) {
						$item.WhsInfo.StandardAveragePrice = $ItemMasterData.AvgStdPrice;
					}
				}
				if (-not $secondWhsExists) {
					$item.WhsInfo.SetCurrentLine($item.WhsInfo.Count - 1);
					if (-not [string]::IsNullOrWhiteSpace($item.WhsInfo.WarehouseCode)) {
						$item.WhsInfo.Add();
					}
					$item.WhsInfo.WarehouseCode = $ItemMasterData.SecondWarehouseCode;
					if (-not $ItemMasterData.InventoryItem) {
						$item.WhsInfo.StandardAveragePrice = $ItemMasterData.AvgStdPrice;
					}
				}
				
				$item.DefaultWarehouse = $ItemMasterData.DefaultWarehouseCode;


				$result = $item.Add();
				if ($result -ne 0) {
					$err = $sapCompany.GetLastErrorDescription();
					throw [System.Exception]($err);
				}

				if ($ItemMasterData.StandardValuationMethod -and $ItemMasterData.InventoryItem) {
					$oMR = $sapCompany.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oMaterialRevaluation);
					$oMR.RevalType = "P";
					$oMR.Lines.ItemCode = $ItemMasterData.ItemCode;
					$oMR.Lines.WarehouseCode = $ItemMasterData.DefaultWarehouseCode;
					$oMR.Lines.Price = $ItemMasterData.AvgStdPrice;
					$oMR.Lines.Add()
					$oMR.Lines.ItemCode = $ItemMasterData.ItemCode;
					$oMR.Lines.WarehouseCode = $ItemMasterData.SecondWarehouseCode;
					$oMR.Lines.Price = $ItemMasterData.AvgStdPrice;
					$oMR.Lines.Add()

					$result = $oMR.Add();
					if ($result -ne 0) {
						$err = [string] $sapCompany.GetLastErrorDescription();
						$msg = [string]::Format("Exception while setting Standard Price. Details: {0}", $err);
						throw [System.Exception] ($msg);
					}
				}
				$sapCompany.EndTransaction([SAPbobsCOM.BoWfTransOpt]::wf_Commit);
			}
			catch {
				if ($sapCompany.InTransaction) {
					$sapCompany.EndTransaction([SAPbobsCOM.BoWfTransOpt]::wf_RollBack);
				}
				$err = [string]$_.Exception.Message;
				$msg = [string]::Format("Item don't exists and adding it to SAP failed. Details: {0}", $err);
				throw [System.Exception]($msg);
			}
		}
	}
 catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception while prepparing Item with ItemCode: {0}. Details: {1}", [string] $ItemMasterData.ItemCode, $err);
		throw ($msg);
	}
}

function prepareBOM([BillOfMaterials] $BillOfMaterials) {
	try {
		$bom = $sapCompany.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oProductTrees);
		$exists = $bom.GetByKey($BillOfMaterials.ItemCode);
		if ($exists) {
			try {
				if ($bom.Warehouse -ne $BillOfMaterials.WarehouseCode) {
					throw [System.Exception] ([string]::Format("Warehouse set in SAP to {0} when it should be {1}", [string]$bom.Warehouse, [string]$BillOfMaterials.WarehouseCode));
				}
				if ($bom.Quantity -ne $BillOfMaterials.Quantity) {
					throw [System.Exception] ([string]::Format("Quantity set in SAP to {0} when it should be {1}", [string]$bom.Quantity, [string]$BillOfMaterials.Quantity));
				}

				if ($bom.Items.Count -ne $BillOfMaterials.Lines.Count) {
					throw [System.Exception] ([string]::Format("Numer of Items set in SAP Tree is {0} when it should be {1}", [string]$bom.Items.Count, [string]$BillOfMaterials.Lines.Count));
				}

				for ($i = 0; $i -lt $BillOfMaterials.Lines.Count; $i++) {
					$bom.Items.SetCurrentLine($i);
					$BOMLine = $BillOfMaterials.Lines[$i];

					if ($bom.Items.ItemCode -ne $BOMLine.ItemCode) {
						throw [System.Exception] ([string]::Format("ItemCode in SAP Tree line number {0} is {1} when it should be {2}", ($i + 1), [string]$bom.Items.ItemCode, [string]$BOMLine.ItemCode));
					}
					if ($bom.Items.Warehouse -ne $BOMLine.WarehouseCode) {
						throw [System.Exception] ([string]::Format("Warehouse in SAP Tree line number {0} is {1} when it should be {2}", ($i + 1), [string]$bom.Items.Warehouse, [string]$BOMLine.WarehouseCode));
					}
					if ($bom.Items.Quantity -ne $BOMLine.Quantity) {
						throw [System.Exception] ([string]::Format("Quantity in SAP Tree line number {0} is {1} when it should be {2}", ($i + 1), [string]$bom.Items.Quantity, [string]$BOMLine.Quantity));
					}
				}
			}
			catch {
				$err = [string]$_.Exception.Message;
				$msg = [string]::Format("Bill Of Materials already exists but it has wrong settings. Details: {0}", $err);
				throw [System.Exception]($msg);
			}
		}
		else {
			try {
				$bom.TreeCode = $BillOfMaterials.ItemCode;
				$bom.Warehouse = $BillOfMaterials.WarehouseCode;
				$bom.Quantity = $BillOfMaterials.Quantity;
				foreach ($BOMLine in $BillOfMaterials.Lines) {
					$bom.Items.ItemCode = $BOMLine.ItemCode;
					$bom.Items.Warehouse = $BOMLine.WarehouseCode;
					$bom.Items.Quantity = $BOMLine.Quantity;
					$bom.Items.Add();
				}

				$result = $bom.Add();
				if ($result -ne 0) {
					$err = [string] $sapCompany.GetLastErrorDescription();
					throw [System.Exception] ($err);
				}
			}
			catch {
				$err = [string]$_.Exception.Message;
				$msg = [string]::Format("Bill Of Materials don't exists and adding it to SAP failed. Details: {0}", $err);
				throw [System.Exception]($msg);
			}
		}
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception while prepparing Bill Of Materials with ItemCode: {0}. Details: {1}", [string] $BillOfMaterials.ItemCode, $err);
		throw ($msg);
	}
}
function createProductionOrderUsingDI($po) {
	try {
		$result = $po.Add();
		if ($result -ne 0) {
			$err = [string] $sapCompany.GetLastErrorDescription();
			throw [System.Exception] ($err);
		}
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception while adding Production Order by DI. Details: {0}", $err);
		throw ($msg);
	}
}

function createProductionOrderUsingXML($po) {
	try {
		if (Test-Path -Path $TEMP_XML_FILE) {
			Remove-Item -Path $TEMP_XML_FILE
		}
		$po.SaveXML($TEMP_XML_FILE);
		$prodOrder = $sapCompany.GetBusinessObjectFromXML($TEMP_XML_FILE, 0);
		$result = $prodOrder.Add();
		
		if ($result -ne 0) {
			$err = [string] $sapCompany.GetLastErrorDescription();
			throw [System.Exception] ($err);
		}
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception while adding Production Order by XML. Details: {0}", $err);
		throw ($msg);
	} finally {
		if (Test-Path -Path $TEMP_XML_FILE) {
			Remove-Item -Path $TEMP_XML_FILE
		}
	}
}

function createProductionOrder([ProductionOrder] $ProductionOrder, $type) {
	try {
		$po = $sapCompany.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oProductionOrders);
		$po.ItemNo = $ProductionOrder.ItemCode;
		$po.PlannedQuantity = $ProductionOrder.PlannedQuantity;
		$po.Warehouse = $ProductionOrder.WarehouseCode;
	
		$linesToDelete = New-Object  'System.Collections.Generic.List[int]';
		for ($i = 0; $i -lt $po.Lines.Count; $i++) {
			if ($ProductionOrder.Lines.Where( { $_.LineNum -eq $i }).Count -eq 0) {
				$linesToDelete = $i;
			}
		}

		for ($i = 0; $i -lt $po.Lines.Count; $i++) {
			$poLines = $ProductionOrder.Lines.Where( { $_.LineNum -eq $i });

			if ($poLines.Count -eq 0) {
				$linesToDelete.Add($i);
				break;
			}
			elseif ($poLines.Count -gt 1) {
				throw [System.Exception] (([string]::Format("Incorrect definition of Production Order. LineNum: {0} occures more than once", $i)));
			}
			$poLine = $poLines[0];
			$po.Lines.SetCurrentLine($i);
			try {
				if ($po.Lines.ItemNo -ne $poLine.ItemCode) {
					$po.Lines.ItemNo = $poLine.ItemCode;
				}
			}
			catch {
				$err = [string]$_.Exception.Message;
				$msg = [string]::Format("Exception while changing Item Code from: {0} to: {1}. Details: {2}", [string] $po.Lines.ItemNo, [string] $poLine.ItemCode, $err);
				throw ($msg);
			}
			try {
				if ($po.Lines.Warehouse -ne $poLine.WarehouseCode) {
					$po.Lines.Warehouse = $poLine.WarehouseCode;
				}
			}
			catch {
				$err = [string]$_.Exception.Message;
				$msg = [string]::Format("Exception while changing Warehouse from: {0} to: {1}. Details: {2}", [string] $po.Lines.Warehouse, [string] $poLine.WarehouseCode, $err);
				throw ($msg);
			}
			try {
				if ($po.Lines.PlannedQuantity -ne $poLine.PlannedQuantity) {
					$po.Lines.PlannedQuantity = $poLine.Quantity;
				}
			}
			catch {
				$err = [string]$_.Exception.Message;
				$msg = [string]::Format("Exception while changing Quantity from: {0} to: {1}. Details: {2}", [string] $po.Lines.PlannedQuantity, [string] $poLine.Quantity, $err);
				throw ($msg);
			}
		}
		$po.Lines.SetCurrentLine($po.Lines.Count - 1);

		foreach ($poLine in $ProductionOrder.Lines.Where( { $_.LineNum -eq -1 })) {
			if (-not [stirng]::IsNullOrWhiteSpace($po.Lines.ItemNo)) {
				$po.Lines.Add();
			}
			$po.Lines.ItemNo = $poLine.ItemCode;
			$po.Lines.Warehouse = $poLine.WarehouseCode;
			$po.Lines.PlannedQuantity = $poLine.Quantity;
		}

		#remove lines
		if ($linesToDelete -gt 0) {
			$linesToDelete.Sort();
			$linesToDelete.Reverse();
			foreach ($LineNum in $linesToDelete) {
				$po.Lines.SetCurrentLine($LineNum);
				try {
					$po.lines.Delete();
				}
				catch {
					$err = [string]$_.Exception.Message;
					throw [System.Exception] (([string]::Format("Couldn't delete line with LineNum: {0}", $LineNum)));
				}
			}
		}
		if ($type -eq [TransactionType]::DI) {
			createProductionOrderUsingDI -po $po;
		}
		elseif ($type -eq [TransactionType]::XML) {
			createProductionOrderUsingXML -po $po;
		}
		else {
			throw [System.Exception](([string]::Format("Transaction Type {0} is not supported.", $type)));
		}
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception while creating production order with ItemCode: {0}. Details: {1}", [string] $ProductionOrder.ItemCode, $err);
		throw ($msg);
	}
}

function createProductionOrderFromBOM([BillOfMaterials] $bom) {
	[ProductionOrder] $ProductionOrder = New-Object 'ProductionOrder'($bom.ItemCode, $bom.WarehouseCode, $bom.Quantity);

	$i = 0;
	foreach ($bomLine in $bom.Lines) {
		$ProductionOrder.addLine($bomLine.ItemCode, $bomLine.WarehouseCode, $bomLine.Quantity, $i);
		$i++;
	}
	return $ProductionOrder;
}
function canWeChangeHeaderWarehouseWhenCreatingProductionOrder([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		$ProductionOrder.WarehouseCode = $WHS_CODE_2;
		createProductionOrder -type $type -ProductionOrder $ProductionOrder
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $canWeChangeHeaderWarehouseWhenCreatingProductionOrder, [string] $type, [string] $err);
		throw ($msg);
	}
}

# setup and check - preapare master data for test, check warehouses, check if it is possible to add standard Production Order and BOM - DI and XML
#check FOD 
function setupSAPMasterData($test) {
	#TODO check warehouses

	#region prepare Item Master Data
	prepareItem -ItemMasterData $CoD;
	prepareItem -ItemMasterData $FoD;
	prepareItem -ItemMasterData $PH;
	prepareItem -ItemMasterData $A;
	prepareItem -ItemMasterData $B;
	prepareItem -ItemMasterData $C;
	prepareItem -ItemMasterData $D;
	prepareItem -ItemMasterData $F;
	prepareItem -ItemMasterData $H;
	prepareItem -ItemMasterData $X1;
	prepareItem -ItemMasterData $X2;
	prepareItem -ItemMasterData $X3;
	prepareItem -ItemMasterData $X4;
	#endregion
	
	#region prepare Bill Of Materials
	prepareBOM -BillOfMaterials $BOMFoD;
	prepareBOM -BillOfMaterials $BOMA;
	prepareBOM -BillOfMaterials $BOMPH;
	prepareBOM -BillOfMaterials $BOMD;
	#endregion
	
}

function runTests() {
	$transactionTypeDI = [TransactionType]::DI;
	$transactionTypeXML = [TransactionType]::XML;
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = canWeChangeHeaderWarehouseWhenCreatingProductionOrder -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}

		try {
			$SuccessXML = canWeChangeHeaderWarehouseWhenCreatingProductionOrder -bom $BOMA -type $transactionTypeXML;
		}
		catch {
			$SuccessXML = $false;
			$errXML = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errXML;
		}

	}
	catch {
		
	}
	finally {
		$TEST_RESULT.AddTestResult($canWeChangeHeaderWarehouseWhenCreatingProductionOrder, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}

	# [TransactionType]::DI
	
	# canWeChangeHeaderWarehouseWhenCreatingProductionOrder -bom $BOMA -type [TransactionType
	#	canWeChangeHeaderWarehouseWhenCreatingProductionOrder -bom
	#canWeChangeHeaderWarehouseWhenCreatingProductionOrder - update, add 
}


setupSAPMasterData
runTests






