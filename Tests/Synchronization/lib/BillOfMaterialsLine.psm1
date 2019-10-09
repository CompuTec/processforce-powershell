class BillOfMaterialsLine {
	[string] $ItemCode = "";
	[string] $WarehouseCode = "";
	[double] $Quantity = 0;
	

	BillOfMaterialsLine([string]$ItemCode, [string] $WarehouseCode, [double] $Quantity){
		$this.ItemCode = $ItemCode;
		$this.WarehouseCode = $WarehouseCode;
		$this.Quantity = $Quantity;
	}


}

