class ProductionOrderLine {
	[int] $LineNum = -1;
	[string] $ItemCode = "";
	[string] $WarehouseCode = "";
	[double] $Quantity = 0;
	

	ProductionOrderLine([string]$ItemCode, [string] $WarehouseCode, [double] $Quantity){
		$this.ItemCode = $ItemCode;
		$this.WarehouseCode = $WarehouseCode;
		$this.Quantity = $Quantity;
	}

	ProductionOrderLine([string]$ItemCode, [string] $WarehouseCode, [double] $Quantity, [int] $LineNum){
		$this.ItemCode = $ItemCode;
		$this.WarehouseCode = $WarehouseCode;
		$this.Quantity = $Quantity;
		$this.LineNum = $LineNum;
	}


}

