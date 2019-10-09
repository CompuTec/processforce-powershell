class ItemMasterData {
	[string] $ItemCode = "";
	[bool] $InventoryItem = $false;
	[bool] $SalesItem = $false;
	[bool] $PurchaseItem = $false;
	[bool] $PhantomItem = $false;
	[bool] $AssetItem = $false;
	[bool] $ManageByBatches = $false;
	[bool] $ManageBySerialNumbers = $false;
	[bool] $StandardValuationMethod = $true;
	[double] $AvgStdPrice = 1;
	[string] $DefaultWarehouseCode = "";

	ItemMasterData(){}

	ItemMasterData([string]$ItemCode, [string] $DefaultWarehouseCode){
		$this.ItemCode = $ItemCode;
		$this.DefaultWarehouseCode = $DefaultWarehouseCode;
	}

	ItemMasterData([string]$ItemCode, [bool] $InventoryItem, [bool] $SalesItem, [bool] $PurchaseItem, [bool] $PhantomItem, [bool] $AssetItem, [bool] $ManageByBatches, [bool] $ManageBySerialNumbers, [bool]$StandardValuationMethod, [double] $AvgStdPrice, [string] $DefaultWarehouseCode) {
		$this.ItemCode = $ItemCode;
		$this.InventoryItem = $InventoryItem;
		$this.SalesItem = $SalesItem;
		$this.PurchaseItem = $PurchaseItem;
		$this.PhantomItem = $PhantomItem;
		$this.AssetItem = $AssetItem;
		$this.ManageByBatches = $ManageByBatches;
		$this.ManageBySerialNumbers = $ManageBySerialNumbers;
		$this.StandardValuationMethod = $StandardValuationMethod;
		$this.AvgStdPrice = $AvgStdPrice;
		$this.DefaultWarehouseCode = $DefaultWarehouseCode;
	}

	static [ItemMasterData] getNewRegularItem([string]$ItemCode, [string] $DefaultWarehouseCode)
	{
		$imd = New-Object ItemMasterData($ItemCode, $DefaultWarehouseCode);
		$imd.InventoryItem = $true;
		$imd.SalesItem = $true;
		$imd.PurchaseItem = $true;
		$imd.PhantomItem = $false;
		$imd.AssetItem = $false;
		$imd.ManageByBatches = $false;
		$imd.ManageBySerialNumbers = $false;
		$imd.StandardValuationMethod = $false;
		$imd.AvgStdPrice = 0;
		return $imd;
	}

	static [ItemMasterData] getNewCoproductDummy([string]$ItemCode, [string] $DefaultWarehouseCode)
	{
		$imd = New-Object ItemMasterData($ItemCode, $DefaultWarehouseCode);
		$imd.InventoryItem = $false;
		$imd.SalesItem = $false;
		$imd.PurchaseItem = $false;
		$imd.PhantomItem = $false;
		$imd.AssetItem = $false;
		$imd.ManageByBatches = $false;
		$imd.ManageBySerialNumbers = $false;
		$imd.StandardValuationMethod = $true;
		$imd.AvgStdPrice = 1;
		return $imd;
	}

	static [ItemMasterData] getNewFinalDummy([string]$ItemCode, [string] $DefaultWarehouseCode)
	{
		$imd = New-Object ItemMasterData($ItemCode, $DefaultWarehouseCode);
		$imd.InventoryItem = $true;
		$imd.SalesItem = $false;
		$imd.PurchaseItem = $false;
		$imd.PhantomItem = $false;
		$imd.AssetItem = $false;
		$imd.ManageByBatches = $false;
		$imd.ManageBySerialNumbers = $false;
		$imd.StandardValuationMethod = $true;
		$imd.AvgStdPrice = 1;
		return $imd;
	}
	static [ItemMasterData] getNewPhantom([string]$ItemCode, [string] $DefaultWarehouseCode)
	{
		$imd = New-Object ItemMasterData($ItemCode, $DefaultWarehouseCode);
		$imd.InventoryItem = $false;
		$imd.SalesItem = $true;
		$imd.PurchaseItem = $true;
		$imd.PhantomItem = $true;
		$imd.AssetItem = $false;
		$imd.ManageByBatches = $false;
		$imd.ManageBySerialNumbers = $false;
		$imd.StandardValuationMethod = $false;
		$imd.AvgStdPrice = 0;
		return $imd;
	}
}