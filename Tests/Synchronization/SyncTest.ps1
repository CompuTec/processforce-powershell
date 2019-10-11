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
#region test names
$canWeChangeHeaderWarehouseWhenCreatingProductionOrder = "Can We Change Header Warehouse When Creating Production Order";
$canWeChangeLinesWhenCreatingProducionOrder_ItemCode = "Can We Change Lines when creating Production Order - Item Code";
$canWeChangeLinesWhenCreatingProducionOrder_WarehouseCode = "Can We Change Lines when creating Production Order - Warehouse Code";
$canWeChangeLinesWhenCreatingProducionOrder_Quantity = "Can We Change Lines when creating Production Order - Quantity";
$canWeChangeLinesWhenCreatingProducionOrder_AddLineNotFromOITT = "Can We Change Lines when creating Production Order - Add Line not from OITT";
$canWeChangeLinesWhenCreatingProducionOrder_DeleteLineFromOITT = "Can We Change Lines when creating Production Order - Delete Line from OITT";
$CanWeAddProductionOrderInReleasedStatus = "Can We Add Production Order in Released Status";
$CanWeChangeStatusFromPlannedToClosed = "Can We change status from Planned to Closed";
$CanWeChangeHeaderItemCodeWhenStausIsReleased = "Can We Change Header Item Code when Staus is Released";
$CanWeChangeHeaderWarehouseWhenChangingHeaderItemCode = "Can We Change Header Warehouse when Changing Header Item Code";
$CanWeChangeLinesWhenChangingHeaderItemCode_ItemCode = "Can We Change Lines when Changing Header Item Code - Item Code";
$CanWeChangeLinesWhenChangingHeaderItemCode_WarehouseCode = "Can We Change Lines when Changing Header Item Code - Warehouse Code";
#endregion
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

function getProductionOrder($key) {
	try {
		$po = $sapCompany.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oProductionOrders);

		$result = $po.GetByKey($key);
		if (-not $result) {
			$err = [string] $sapCompany.GetLastErrorDescription();
			throw [System.Exception] ($err);
		}
		return $po;
	}
 catch {
		$err = [string]$_.Exception.Message;
		throw [System.Exception] (([string]::Format("Couldn't get Production Order with key: {0}", [string]$key)));
	}
}
function saveProductionOrderUsingDI($po, $task) {
	try {
		$result = -1;
		
		if ($task -eq [TransactionTask]::Add) {
			$result = $po.Add();
		}
		elseif ($task -eq [TransactionTask]::Update) {
			$result = $po.Update();
		}
		else {
			throw [System.Exception](([string]::Format("Incorrect transaction type: {0}", [string] $task)));
		}
		if ($result -ne 0) {
			$err = [string] $sapCompany.GetLastErrorDescription();
			throw [System.Exception] ($err);
		}
		if ($task -eq [TransactionTask]::Add) {
			$DocEntry = $sapCompany.GetNewObjectKey();
			return $DocEntry;
		}
		return $null;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception while adding/updating Production Order by DI. Details: {0}", $err);
		throw ($msg);
	}
}
function prepareProductionOrderXML() {
	try {
		$nodesToBeRemovedFromOWOR = @("OriginAbs", "OriginNum", "UserSign");
		$nodesToBeRemovedFromWOR1 = @("ResAlloc", "StageId");

		[xml] $ProductionOrderXml = Get-Content -Encoding UTF8 $TEMP_XML_FILE;
		$xmlOWOR = $ProductionOrderXml.SelectSingleNode("/BOM/BO/OWOR/row");

		foreach ($nodeName in $nodesToBeRemovedFromOWOR) {
			$node = $xmlOWOR.SelectSingleNode($nodeName);
			if ($node) {
				$xmlOWOR.RemoveChild($node);
			}
		}
		$xmlWOR1s = $ProductionOrderXml.SelectNodes("/BOM/BO/WOR1/row");
	
		foreach ($xmlWOR1 in $xmlWOR1s) {
			foreach ($nodeName in $nodesToBeRemovedFromWOR1) {
				$node = $xmlWOR1.SelectSingleNode($nodeName);
				if ($node) {
					$xmlWOR1.RemoveChild($node);
				}
			}
		}
		if (Test-Path -Path $TEMP_XML_FILE) {
			Remove-Item -Path $TEMP_XML_FILE
		}
		Add-Content -Path $TEMP_XML_FILE $ProductionOrderXml.OuterXml;	
	}
 catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception while preparing xml file for Production Order. Details: {0}", $err);
		throw ($msg);
	}
}
function saveProductionOrderUsingXML($po, $task) {
	try {
		if (Test-Path -Path $TEMP_XML_FILE) {
			Remove-Item -Path $TEMP_XML_FILE
		}
		$po.SaveXML($TEMP_XML_FILE);
		prepareProductionOrderXML
		$prodOrder = $sapCompany.GetBusinessObjectFromXML($TEMP_XML_FILE, 0);
		$result = -1;
		if ($task -eq [TransactionTask]::Add) {
			$result = $prodOrder.Add();
		}
		elseif ($task -eq [TransactionTask]::Update) {
			$result = $prodOrder.Update();
		}
		else {
			throw [System.Exception](([string]::Format("Incorrect transaction type: {0}", [string] $task)));
		}
		
		if ($result -ne 0) {
			$err = [string] $sapCompany.GetLastErrorDescription();
			throw [System.Exception] ($err);
		}
		if ($task -eq [TransactionTask]::Add) {
			$DocEntry = $sapCompany.GetNewObjectKey();
			return $DocEntry;
		}
		return $null;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception while adding/updating Production Order by XML. Details: {0}", $err);
		throw ($msg);
	}
 finally {
		if (Test-Path -Path $TEMP_XML_FILE) {
			Remove-Item -Path $TEMP_XML_FILE
		}
	}
}

function createProductionOrder([ProductionOrder] $ProductionOrder, $type) {
	$task = [TransactionTask]::Add;
	return changeProductionOrder -ProductionOrder $ProductionOrder -type $type -task $task;
}

function updateProductionOrder([ProductionOrder] $ProductionOrder, $type, $po) {
	$task = [TransactionTask]::Update;
	return changeProductionOrder -ProductionOrder $ProductionOrder -type $type -task $task -po $po;
}

function changeProductionOrder([ProductionOrder] $ProductionOrder, $type, $task, $po = $null) {
	try {
		if ($task -eq [TransactionTask]::Add) {
			$po = $sapCompany.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oProductionOrders);
		}
		elseif ($task -eq [TransactionTask]::Update) {
			if ($null -eq $po) {
				throw [System.Exception]("SAP Production Order document not provided");
			}
		}
		else {
			throw [System.Exception](([string]::Format("Not supported task: {0}", $task)));
		}
		$po.ItemNo = $ProductionOrder.ItemCode;
		$po.PlannedQuantity = $ProductionOrder.Quantity;
		$po.Warehouse = $ProductionOrder.WarehouseCode;
		if ($ProductionOrder.IsReleased) {
			$po.ProductionOrderStatus = [SAPbobsCOM.BoProductionOrderStatusEnum]::boposReleased;
		}
		if ($ProductionOrder.IsClosed) {
			$po.ProductionOrderStatus = [SAPbobsCOM.BoProductionOrderStatusEnum]::boposClosed;
		}
	
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
				if ($po.Lines.PlannedQuantity -ne $poLine.Quantity) {
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
			if (-not [string]::IsNullOrWhiteSpace($po.Lines.ItemNo)) {
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
			$DocEntry = ( saveProductionOrderUsingDI -po $po -task $task );
			return $DocEntry;
		}
		elseif ($type -eq [TransactionType]::XML) {
			$DocEntry = ( saveProductionOrderUsingXML -po $po -task $task );
			return $DocEntry;
		}
		else {
			throw [System.Exception](([string]::Format("Transaction Type {0} is not supported.", $type)));
		}
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception while saving ({0}) production order with ItemCode: {1}. Details: {2}", [string] $task, [string] $ProductionOrder.ItemCode, $err);
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
#region ADD TESTS
function canWeChangeHeaderWarehouseWhenCreatingProductionOrder([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		$ProductionOrder.WarehouseCode = $WHS_CODE_2;
		createProductionOrder -type $type -ProductionOrder $ProductionOrder;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $canWeChangeHeaderWarehouseWhenCreatingProductionOrder, [string] $type, [string] $err);
		throw ($msg);
	}
}
function canWeChangeLinesWhenCreatingProducionOrder_ItemCode([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		$ProductionOrder.Lines[0].ItemCode = $X1.ItemCode;
		createProductionOrder -type $type -ProductionOrder $ProductionOrder;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $canWeChangeLinesWhenCreatingProducionOrder_ItemCode, [string] $type, [string] $err);
		throw ($msg);
	}
}
function canWeChangeLinesWhenCreatingProducionOrder_WarehouseCode([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		$ProductionOrder.Lines[0].WarehouseCode = $WHS_CODE_2;
		createProductionOrder -type $type -ProductionOrder $ProductionOrder;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $canWeChangeLinesWhenCreatingProducionOrder_WarehouseCode, [string] $type, [string] $err);
		throw ($msg);
	}
}
function canWeChangeLinesWhenCreatingProducionOrder_Quantity([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		$ProductionOrder.Lines[0].Quantity = $ProductionOrder.Lines[0].Quantity + 1;
		createProductionOrder -type $type -ProductionOrder $ProductionOrder;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $canWeChangeLinesWhenCreatingProducionOrder_Quantity, [string] $type, [string] $err);
		throw ($msg);
	}
}
function canWeChangeLinesWhenCreatingProducionOrder_AddLineNotFromOITT([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		$ProductionOrder.addLine($X1.ItemCode, $X1.DefaultWarehouseCode, 1, -1);
		createProductionOrder -type $type -ProductionOrder $ProductionOrder;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $canWeChangeLinesWhenCreatingProducionOrder_AddLineNotFromOITT, [string] $type, [string] $err);
		throw ($msg);
	}
}
function canWeChangeLinesWhenCreatingProducionOrder_DeleteLineFromOITT([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		$ProductionOrder.Lines.RemoveAt($ProductionOrder.Lines.Count - 1);
		createProductionOrder -type $type -ProductionOrder $ProductionOrder;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $canWeChangeLinesWhenCreatingProducionOrder_DeleteLineFromOITT, [string] $type, [string] $err);
		throw ($msg);
	}
}
function CanWeAddProductionOrderInReleasedStatus([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		$ProductionOrder.IsReleased = $true;
		createProductionOrder -type $type -ProductionOrder $ProductionOrder;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $CanWeAddProductionOrderInReleasedStatus, [string] $type, [string] $err);
		throw ($msg);
	}
}
#endregion
#region UPDATE TESTS
function CanWeChangeStatusFromPlannedToClosed([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		try {
			$prepareDocType = [TransactionType]::DI;
			$DocEntry = createProductionOrder -type $prepareDocType -ProductionOrder $ProductionOrder;
			$po = getProductionOrder -key $DocEntry;
			$ProductionOrder.IsReleased = $true;
			updateProductionOrder -ProductionOrder $ProductionOrder -type $prepareDocType -po $po;
		}
		catch {
			$err = [string]$_.Exception.Message;
			$msg = [string]::Format("Exception while preparing to Production Order to test. Details: {0}", [string] $err);
			throw ($msg);
		}
		$ProductionOrder.IsClosed = $true;
		updateProductionOrder -ProductionOrder $ProductionOrder -type $type -po $po;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $CanWeChangeStatusFromPlannedToClosed, [string] $type, [string] $err);
		throw ($msg);
	}
}
function CanWeChangeHeaderItemCodeWhenStausIsReleased([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		try {
			$prepareDocType = [TransactionType]::DI;
			$DocEntry = createProductionOrder -type $prepareDocType -ProductionOrder $ProductionOrder;
			$po = getProductionOrder -key $DocEntry;
			$ProductionOrder.IsReleased = $true;
			updateProductionOrder -ProductionOrder $ProductionOrder -type $prepareDocType -po $po;
		}
		catch {
			$err = [string]$_.Exception.Message;
			$msg = [string]::Format("Exception while preparing to Production Order to test. Details: {0}", [string] $err);
			throw ($msg);
		}
		$ProductionOrder.ItemCode = $FoD.ItemCode;
		updateProductionOrder -ProductionOrder $ProductionOrder -type $type -po $po;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $CanWeChangeHeaderItemCodeWhenStausIsReleased, [string] $type, [string] $err);
		throw ($msg);
	}
}
function CanWeChangeHeaderWarehouseWhenChangingHeaderItemCode([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		try {
			$prepareDocType = [TransactionType]::DI;
			$DocEntry = createProductionOrder -type $prepareDocType -ProductionOrder $ProductionOrder;
			$po = getProductionOrder -key $DocEntry;
		}
		catch {
			$err = [string]$_.Exception.Message;
			$msg = [string]::Format("Exception while preparing to Production Order to test. Details: {0}", [string] $err);
			throw ($msg);
		}
		$ProductionOrder.ItemCode = $FoD.ItemCode;
		$ProductionOrder.WarehouseCode = $Fod.SecondWarehouseCode;
		updateProductionOrder -ProductionOrder $ProductionOrder -type $type -po $po;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $CanWeChangeHeaderWarehouseWhenChangingHeaderItemCode, [string] $type, [string] $err);
		throw ($msg);
	}
}
function CanWeChangeLinesWhenChangingHeaderItemCode_ItemCode([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		try {
			$prepareDocType = [TransactionType]::DI;
			$DocEntry = createProductionOrder -type $prepareDocType -ProductionOrder $ProductionOrder;
			$po = getProductionOrder -key $DocEntry;
		}
		catch {
			$err = [string]$_.Exception.Message;
			$msg = [string]::Format("Exception while preparing to Production Order to test. Details: {0}", [string] $err);
			throw ($msg);
		}
		$ProductionOrder.ItemCode = $FoD.ItemCode;
		$ProductionOrder.Lines[0].ItemCode = $CoD.ItemCode;
		updateProductionOrder -ProductionOrder $ProductionOrder -type $type -po $po;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $CanWeChangeLinesWhenChangingHeaderItemCode_ItemCode, [string] $type, [string] $err);
		throw ($msg);
	}
}
function CanWeChangeLinesWhenChangingHeaderItemCode_WarehouseCode([BillOfMaterials] $bom, $type) {
	try {
		[ProductionOrder] $ProductionOrder = createProductionOrderFromBOM($bom);
		try {
			$prepareDocType = [TransactionType]::DI;
			$DocEntry = createProductionOrder -type $prepareDocType -ProductionOrder $ProductionOrder;
			$po = getProductionOrder -key $DocEntry;
		}
		catch {
			$err = [string]$_.Exception.Message;
			$msg = [string]::Format("Exception while preparing to Production Order to test. Details: {0}", [string] $err);
			throw ($msg);
		}
		$ProductionOrder.ItemCode = $FoD.ItemCode;
		$ProductionOrder.Lines[0].WarehouseCode = $WHS_CODE_2;
		updateProductionOrder -ProductionOrder $ProductionOrder -type $type -po $po;
		return $true;
	}
	catch {
		$err = [string]$_.Exception.Message;
		$msg = [string]::Format("Exception at test '{0}' using {1}. Details: {2}", $CanWeChangeLinesWhenChangingHeaderItemCode_WarehouseCode, [string] $type, [string] $err);
		throw ($msg);
	}
}
#endregion

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
	
	#region canWeChangeHeaderWarehouseWhenCreatingProductionOrder
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
	#endregion
	#region canWeChangeLinesWhenCreatingProducionOrder_ItemCode
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = canWeChangeLinesWhenCreatingProducionOrder_ItemCode -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}
		try {
			$SuccessXML = canWeChangeLinesWhenCreatingProducionOrder_ItemCode -bom $BOMA -type $transactionTypeXML;
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
		$TEST_RESULT.AddTestResult($canWeChangeLinesWhenCreatingProducionOrder_ItemCode, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}
	#endregion
	#region canWeChangeLinesWhenCreatingProducionOrder_WarehouseCode
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = canWeChangeLinesWhenCreatingProducionOrder_WarehouseCode -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}
		try {
			$SuccessXML = canWeChangeLinesWhenCreatingProducionOrder_WarehouseCode -bom $BOMA -type $transactionTypeXML;
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
		$TEST_RESULT.AddTestResult($canWeChangeLinesWhenCreatingProducionOrder_WarehouseCode, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}
	#endregion
	#region canWeChangeLinesWhenCreatingProducionOrder_Quantity
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = canWeChangeLinesWhenCreatingProducionOrder_Quantity -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}
		try {
			$SuccessXML = canWeChangeLinesWhenCreatingProducionOrder_Quantity -bom $BOMA -type $transactionTypeXML;
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
		$TEST_RESULT.AddTestResult($canWeChangeLinesWhenCreatingProducionOrder_Quantity, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}
	#endregion
	#region canWeChangeLinesWhenCreatingProducionOrder_AddLineNotFromOITT
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = canWeChangeLinesWhenCreatingProducionOrder_AddLineNotFromOITT -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}
		try {
			$SuccessXML = canWeChangeLinesWhenCreatingProducionOrder_AddLineNotFromOITT -bom $BOMA -type $transactionTypeXML;
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
		$TEST_RESULT.AddTestResult($canWeChangeLinesWhenCreatingProducionOrder_AddLineNotFromOITT, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}
	#endregion
	#region canWeChangeLinesWhenCreatingProducionOrder_DeleteLineFromOITT
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = canWeChangeLinesWhenCreatingProducionOrder_DeleteLineFromOITT -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}
		try {
			$SuccessXML = canWeChangeLinesWhenCreatingProducionOrder_DeleteLineFromOITT -bom $BOMA -type $transactionTypeXML;
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
		$TEST_RESULT.AddTestResult($canWeChangeLinesWhenCreatingProducionOrder_DeleteLineFromOITT, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}
	#endregion
	#region CanWeAddProductionOrderInReleasedStatus
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = CanWeAddProductionOrderInReleasedStatus -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}
		try {
			$SuccessXML = CanWeAddProductionOrderInReleasedStatus -bom $BOMA -type $transactionTypeXML;
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
		$TEST_RESULT.AddTestResult($CanWeAddProductionOrderInReleasedStatus, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}
	#endregion
	#region CanWeChangeStatusFromPlannedToClosed
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = CanWeChangeStatusFromPlannedToClosed -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}
		try {
			$SuccessXML = CanWeChangeStatusFromPlannedToClosed -bom $BOMA -type $transactionTypeXML;
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
		$TEST_RESULT.AddTestResult($CanWeChangeStatusFromPlannedToClosed, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}
	#endregion
	#region CanWeChangeHeaderItemCodeWhenStausIsReleased
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = CanWeChangeHeaderItemCodeWhenStausIsReleased -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}
		try {
			$SuccessXML = CanWeChangeHeaderItemCodeWhenStausIsReleased -bom $BOMA -type $transactionTypeXML;
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
		$TEST_RESULT.AddTestResult($CanWeChangeHeaderItemCodeWhenStausIsReleased, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}
	#endregion
	#region CanWeChangeHeaderWarehouseWhenChangingHeaderItemCode
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = CanWeChangeHeaderWarehouseWhenChangingHeaderItemCode -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}
		try {
			$SuccessXML = CanWeChangeHeaderWarehouseWhenChangingHeaderItemCode -bom $BOMA -type $transactionTypeXML;
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
		$TEST_RESULT.AddTestResult($CanWeChangeHeaderWarehouseWhenChangingHeaderItemCode, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}
	#endregion
	#region CanWeChangeLinesWhenChangingHeaderItemCode_ItemCode
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = CanWeChangeLinesWhenChangingHeaderItemCode_ItemCode -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}
		try {
			$SuccessXML = CanWeChangeLinesWhenChangingHeaderItemCode_ItemCode -bom $BOMA -type $transactionTypeXML;
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
		$TEST_RESULT.AddTestResult($CanWeChangeLinesWhenChangingHeaderItemCode_ItemCode, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}
	#endregion
	#region CanWeChangeLinesWhenChangingHeaderItemCode_WarehouseCode
	try {
		$SuccessDI = $false;
		$SuccessXML = $false;
		$errDI = [string]::Empty;
		$errXML = [string]::Empty;
		try {
			$SuccessDI = CanWeChangeLinesWhenChangingHeaderItemCode_WarehouseCode -bom $BOMA -type $transactionTypeDI;
		}
		catch {
			$SuccessDI = $false;
			$errDI = [string]$_.Exception.Message;
			Write-Host -BackgroundColor Red -ForegroundColor White $errDI;
		}
		try {
			$SuccessXML = CanWeChangeLinesWhenChangingHeaderItemCode_WarehouseCode -bom $BOMA -type $transactionTypeXML;
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
		$TEST_RESULT.AddTestResult($CanWeChangeLinesWhenChangingHeaderItemCode_WarehouseCode, $SuccessDI, $SuccessXML, $errDI, $errXML);
	}
	#endregion









}


setupSAPMasterData
runTests






