using module .\BillOfMaterialsLine.psm1;
class BillOfMaterials {
	[string] $ItemCode = "";
	[string] $WarehouseCode = "";
	[double] $Quantity = 1;
	[System.Collections.Generic.List[BillOfMaterialsLine]] $Lines;
	
	BillOfMaterials([string]$ItemCode, [string] $WarehouseCode, [double] $Quantity){
		$this.ItemCode = $ItemCode;
		$this.WarehouseCode = $WarehouseCode;
		$this.Quantity = $Quantity;
		$this.Lines = New-Object 'System.Collections.Generic.List[BillOfMaterialsLine]';
	}

	addLine([string]$ItemCode, [string] $WarehouseCode, [double] $Quantity){
		$this.Lines.Add((New-Object 'BillOfMaterialsLine'($ItemCode,$WarehouseCode, $Quantity)));
	}


}

