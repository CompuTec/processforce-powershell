using module .\ProductionOrderLine.psm1;
class ProductionOrder {
	[string] $ItemCode = "";
	[string] $WarehouseCode = "";
	[double] $Quantity = 1;
	[bool] $IsReleased = $false;
	[System.Collections.Generic.List[ProductionOrderLine]] $Lines;
	
	ProductionOrder([string]$ItemCode, [string] $WarehouseCode, [double] $Quantity){
		$this.ItemCode = $ItemCode;
		$this.WarehouseCode = $WarehouseCode;
		$this.Quantity = $Quantity;
		$this.Lines = New-Object 'System.Collections.Generic.List[ProductionOrderLine]';
	}

	addLine([string]$ItemCode, [string] $WarehouseCode, [double] $Quantity, [int] $LineNum){
		$this.Lines.Add((New-Object 'ProductionOrderLine'($ItemCode,$WarehouseCode, $Quantity, $LineNum)));
	}


}

