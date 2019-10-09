using module .\lib\ItemMasterData.psm1;
using module .\lib\BillOfMaterials.psm1;
add-type -Path "C:\Projects\Playground\SAP\DLL\SAPHana\x64\Interop.SAPbobsCOM.dll"

#region Script parameters

$csvImportCatalog = $PSScriptRoot + "\"

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
	$code = $sapCompany.Connect()
 
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
					foreach ($whs in $item.WhsInfo) {
						if ($whs.WarehouseCode -eq $ItemMasterData.DefaultWarehouseCode) {
							if ($ItemMasterData.AvgStdPrice -ne $whs.StandardAveragePrice) {
								throw [System.Exception] ([string]::Format("Item Cost is set in SAP to {0} on Warehouse {1} when it should be {2}", [string]$whs.StandardAveragePrice, [string]$ItemMasterData.DefaultWarehouseCode, [string]$ItemMasterData.AvgStdPrice));
							}
							break;
						}
					}
				}

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
					if(-not $ItemMasterData.InventoryItem){
						$item.AvgStdPrice = $ItemMasterData.AvgStdPrice;
					}
				}
				
				$whsExists = $false;
				foreach ($whs in $item.WhsInfo) {
					if ($whs.WarehouseCode -eq $ItemMasterData.DefaultWarehouseCode) {
						$whsExists = $true;
						if(-not $ItemMasterData.InventoryItem) {
							$whs.StandardAveragePrice = $ItemMasterData.AvgStdPrice;
						}
					}
				}
				if (-not $whsExists) {
					$item.WhsInfo.SetCurrentLine($item.WhsInfo.Count - 1);
					if (-not [string]::IsNullOrWhiteSpace($item.WhsInfo.WarehouseCode)) {
						$item.WhsInfo.Add();
					}
					$item.WhsInfo.WarehouseCode = $ItemMasterData.DefaultWarehouseCode;
					if(-not $ItemMasterData.InventoryItem) {
						$whs.StandardAveragePrice = $ItemMasterData.AvgStdPrice;
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

function canWeChangeHeaderWarehouseWhenCreatingProductionOrder([bool]OnlyUpdate)

# setup and check - preapare master data for test, check warehouses, check if it is possible to add standard Production Order and BOM - DI and XML
#check FOD 
function setupSAPMasterData() {
	#TODO check warehouses



	#region prepare Item Master Data
	[ItemMasterData] $CoD = [ItemMasterData]::getNewCoproductDummy("SyncTest_CoD", "01");
	prepareItem -ItemMasterData $CoD;
	[ItemMasterData] $FoD = [ItemMasterData]::getNewFinalDummy("SyncTest_FoD", "01");
	prepareItem -ItemMasterData $FoD;
	[ItemMasterData] $PH = [ItemMasterData]::getNewPhantom("SyncTest_PH", "01");
	prepareItem -ItemMasterData $PH;
	[ItemMasterData] $A = [ItemMasterData]::getNewRegularItem("SyncTest_A", "01");
	prepareItem -ItemMasterData $A;
	[ItemMasterData] $B = [ItemMasterData]::getNewRegularItem("SyncTest_B", "01");
	prepareItem -ItemMasterData $B;
	[ItemMasterData] $C = [ItemMasterData]::getNewRegularItem("SyncTest_C", "01");
	prepareItem -ItemMasterData $C;
	[ItemMasterData] $D = [ItemMasterData]::getNewRegularItem("SyncTest_D", "01");
	prepareItem -ItemMasterData $D;
	[ItemMasterData] $F = [ItemMasterData]::getNewRegularItem("SyncTest_F", "01");
	prepareItem -ItemMasterData $F;
	[ItemMasterData] $H = [ItemMasterData]::getNewRegularItem("SyncTest_H", "01");
	prepareItem -ItemMasterData $H;
	[ItemMasterData] $X1 = [ItemMasterData]::getNewRegularItem("SyncTest_X1", "01");
	prepareItem -ItemMasterData $X1;
	[ItemMasterData] $X2 = [ItemMasterData]::getNewRegularItem("SyncTest_X2", "01");
	prepareItem -ItemMasterData $X2;
	[ItemMasterData] $X3 = [ItemMasterData]::getNewRegularItem("SyncTest_X3", "01");
	prepareItem -ItemMasterData $X3;
	[ItemMasterData] $X4 = [ItemMasterData]::getNewRegularItem("SyncTest_X4", "01");
	prepareItem -ItemMasterData $X4;
	#endregion
	
	#region prepare Bill Of Materials
	$BOMFoD = New-Object 'BillOfMaterials'($FoD.ItemCode, $FoD.DefaultWarehouseCode, 1);
	$BOMFoD.addLine($CoD.ItemCode, $CoD.DefaultWarehouseCode, 1);
	prepareBOM -BillOfMaterials $BOMFoD;

	$BOMA = New-Object 'BillOfMaterials'($A.ItemCode, $A.DefaultWarehouseCode, 1);
	$BOMA.addLine($B.ItemCode, $B.DefaultWarehouseCode, 1);
	$BOMA.addLine($C.ItemCode, $C.DefaultWarehouseCode, 1);
	prepareBOM -BillOfMaterials $BOMA;
	
	$BOMPH = New-Object 'BillOfMaterials'($PH.ItemCode, $PH.DefaultWarehouseCode, 1);
	$BOMPH.addLine($X1.ItemCode, $X1.DefaultWarehouseCode, 1);
	prepareBOM -BillOfMaterials $BOMPH;
	
	$BOMD = New-Object 'BillOfMaterials'($D.ItemCode, $D.DefaultWarehouseCode, 1);
	$BOMD.addLine($PH.ItemCode, $PH.DefaultWarehouseCode, 1);
	$BOMD.addLine($A.ItemCode, $A.DefaultWarehouseCode, 1);
	$BOMD.addLine($F.ItemCode, $F.DefaultWarehouseCode, 1);
	$BOMD.addLine($H.ItemCode, $H.DefaultWarehouseCode, 1);
	prepareBOM -BillOfMaterials $BOMD;
	#endregion
	
}

setupSAPMasterData
# perform tests




